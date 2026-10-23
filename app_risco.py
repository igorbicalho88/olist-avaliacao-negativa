"""Avaliação de risco: versão executável do prever.py.

Uso: dê dois cliques no programa e escolha o CSV com os pedidos
(ou arraste o CSV para cima do programa). O resultado é salvo ao lado
do arquivo original, em Excel, com o final "_com_risco.xlsx":
  - aba "Precisam de atenção": só os pedidos em alerta, com o motivo e a ação sugerida
  - aba "Todos os pedidos": a lista completa, do maior para o menor risco
Colunas extras do CSV (ex.: nome e telefone do cliente) são copiadas para o resultado.
"""
import os
import sys

import joblib
import pandas as pd

# Importações explícitas: o modelo salvo usa estas partes do scikit-learn.
# Sem elas, o PyInstaller não percebe que precisa incluí-las no .exe.
import sklearn.pipeline  # noqa: F401
import sklearn.compose  # noqa: F401
import sklearn.preprocessing  # noqa: F401
import sklearn.ensemble  # noqa: F401
import sklearn._loss.loss  # noqa: F401
import sklearn.utils.validation  # noqa: F401
import openpyxl  # noqa: F401  (necessário para gravar o Excel)
from openpyxl.styles import Alignment, Font, PatternFill

CORTE = 0.8        # risco mínimo para o pedido entrar em "Precisam de atenção"
CORTE_MEDIO = 0.6  # a partir daqui o risco é considerado médio

CATEGORICAS = ["categoria", "estado_vendedor", "estado_cliente", "forma_pagamento"]
NUMERICAS = ["dias_entrega", "dias_prometidos", "dias_atraso", "atrasado", "horas_aprovacao",
             "dias_ate_transportadora", "mes_compra", "dia_semana_compra",
             "n_itens", "n_produtos", "n_vendedores", "valor_total", "frete_total", "frete_sobre_valor",
             "peso_total_g", "fotos_media", "tam_descricao_media", "n_parcelas", "n_formas_pagamento",
             "mesmo_estado"]


def caminho_recurso(relativo):
    """Acha arquivos embutidos no .exe (PyInstaller) ou ao lado do script."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relativo)


def ler_csv(caminho):
    """Lê CSV separado por vírgula ou ponto e vírgula (padrão do Excel em português)."""
    pedidos = pd.read_csv(caminho, sep=None, engine="python", encoding="utf-8-sig")
    for c in NUMERICAS:
        if c in pedidos.columns and not pd.api.types.is_numeric_dtype(pedidos[c]):
            pedidos[c] = pd.to_numeric(pedidos[c].astype(str).str.replace(",", "."), errors="coerce")
    return pedidos


def motivo_e_acao(p):
    """Explica em linguagem simples por que o pedido tem risco e o que fazer."""
    motivos, acoes = [], []
    atraso = p["dias_atraso"]
    if pd.notna(atraso) and atraso > 0:
        motivos.append(f"Entregue com {int(round(atraso))} dia(s) de atraso")
        acoes.append("pedir desculpas pelo atraso e oferecer um cupom")
    if p["n_vendedores"] >= 2:
        motivos.append(f"{int(p['n_vendedores'])} vendedores: pode ter chegado só parte do pedido")
        acoes.append("confirmar se o cliente recebeu todos os itens")
    elif p["n_itens"] >= 3:
        motivos.append(f"Pedido com {int(p['n_itens'])} itens")
        acoes.append("confirmar se todos os itens chegaram em ordem")
    if not motivos:
        motivos.append("Combinação de fatores (prazo, frete, produto)")
        acoes.append("verificar se está satisfeito com a compra")
    return " | ".join(motivos), "Contatar o cliente: " + "; ".join(acoes)


def processar(entrada):
    modelo = joblib.load(caminho_recurso(os.path.join("models", "modelo_avaliacao_negativa.joblib")))
    pedidos = ler_csv(entrada)

    faltando = [c for c in ["order_id"] + CATEGORICAS + NUMERICAS if c not in pedidos.columns]
    if faltando:
        raise ValueError("O arquivo não tem estas colunas:\n\n" + ", ".join(faltando))

    prob = modelo.predict_proba(pedidos[modelo.feature_names_in_])[:, 1]
    explicacoes = pedidos.apply(motivo_e_acao, axis=1, result_type="expand")

    # Colunas que não são do modelo (ex.: nome, telefone, data) vão junto para o resultado
    extras = [c for c in pedidos.columns if c not in ["order_id"] + CATEGORICAS + NUMERICAS]

    resultado = pd.concat([
        pedidos[["order_id"] + extras],
        pd.DataFrame({
            "risco": prob.round(3),
            "nivel": ["ALTO" if p >= CORTE else "MÉDIO" if p >= CORTE_MEDIO else "BAIXO" for p in prob],
            "motivo": explicacoes[0],
            "acao_sugerida": explicacoes[1],
        }),
        pedidos[["categoria", "estado_cliente", "valor_total", "dias_atraso", "n_itens", "n_vendedores"]],
    ], axis=1).sort_values("risco", ascending=False)

    # Pedidos de risco baixo não precisam de ação
    baixo = resultado["nivel"] == "BAIXO"
    resultado.loc[baixo, "motivo"] = "-"
    resultado.loc[baixo, "acao_sugerida"] = "Nenhuma ação necessária"

    atencao = resultado[resultado["nivel"] == "ALTO"]

    saida = os.path.splitext(entrada)[0] + "_com_risco.xlsx"
    with pd.ExcelWriter(saida, engine="openpyxl") as excel:
        if len(atencao):
            atencao.to_excel(excel, sheet_name="Precisam de atenção", index=False)
        else:
            pd.DataFrame({"mensagem": ["Nenhum pedido com risco alto neste arquivo."]}).to_excel(
                excel, sheet_name="Precisam de atenção", index=False)
        resultado.to_excel(excel, sheet_name="Todos os pedidos", index=False)
        for aba in excel.book.worksheets:
            formatar(aba)
    return saida, len(resultado), len(atencao)


def formatar(aba):
    """Cabeçalho destacado, larguras ajustadas e cores por nível de risco."""
    cores = {"ALTO": "F8CBAD", "MÉDIO": "FFE699", "BAIXO": "C6EFCE"}
    cabecalho = [c.value for c in aba[1]]
    for c in aba[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for i, nome in enumerate(cabecalho, start=1):
        valores = [str(aba.cell(row=r, column=i).value or "") for r in range(1, aba.max_row + 1)]
        aba.column_dimensions[aba.cell(row=1, column=i).column_letter].width = min(max(map(len, valores)) + 2, 60)
    if "nivel" in cabecalho:
        col_nivel = cabecalho.index("nivel") + 1
        for r in range(2, aba.max_row + 1):
            cor = cores.get(aba.cell(row=r, column=col_nivel).value)
            if cor:
                for c in aba[r]:
                    c.fill = PatternFill("solid", fgColor=cor)
    if "risco" in cabecalho:
        col = cabecalho.index("risco") + 1
        for r in range(2, aba.max_row + 1):
            aba.cell(row=r, column=col).number_format = "0%"
    aba.freeze_panes = "B2"


def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    janela = tk.Tk()
    janela.withdraw()

    if len(sys.argv) > 1:          # arquivo arrastado para cima do programa
        entrada = sys.argv[1]
    else:                          # dois cliques: pede o arquivo
        entrada = filedialog.askopenfilename(
            title="Escolha o CSV com os pedidos",
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")])
    if not entrada:
        return

    try:
        saida, total, alertas = processar(entrada)
    except PermissionError:
        messagebox.showerror("Avaliação de risco",
                             "Feche o arquivo de resultado no Excel e tente de novo.")
        return
    except Exception as erro:
        messagebox.showerror("Avaliação de risco", f"Não foi possível processar o arquivo.\n\n{erro}")
        return

    abrir = messagebox.askyesno(
        "Avaliação de risco",
        f"{total} pedidos analisados\n{alertas} precisam de atenção (risco alto)\n\n"
        f"Resultado salvo em:\n{saida}\n\nAbrir o resultado agora?")
    if abrir and hasattr(os, "startfile"):
        os.startfile(saida)


if __name__ == "__main__":
    main()
