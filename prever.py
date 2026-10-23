import argparse
import joblib
import pandas as pd

# Colunas que o modelo precisa receber
COLUNAS = ["categoria", "estado_vendedor", "estado_cliente", "forma_pagamento",
           "dias_entrega", "dias_prometidos", "dias_atraso", "atrasado", "horas_aprovacao",
           "dias_ate_transportadora", "mes_compra", "dia_semana_compra",
           "n_itens", "n_produtos", "n_vendedores", "valor_total", "frete_total", "frete_sobre_valor",
           "peso_total_g", "fotos_media", "tam_descricao_media", "n_parcelas", "n_formas_pagamento",
           "mesmo_estado"]


def main():
    # 1) Lê as opções passadas no terminal
    p = argparse.ArgumentParser(description="Calcula o risco de avaliação negativa de pedidos entregues")
    p.add_argument("--entrada", required=True, help="CSV com os pedidos")
    p.add_argument("--saida", required=True, help="CSV de saída com o risco")
    p.add_argument("--modelo", default="models/modelo_avaliacao_negativa.joblib")
    p.add_argument("--corte", type=float, default=0.8, help="probabilidade mínima para alerta")
    a = p.parse_args()

    # 2) Carrega o modelo salvo e os pedidos
    modelo = joblib.load(a.modelo)
    pedidos = pd.read_csv(a.entrada)

    # 3) Confere se todas as colunas necessárias existem
    faltando = [c for c in COLUNAS if c not in pedidos.columns]
    if faltando:
        raise SystemExit(f"Colunas faltando: {faltando}")

    # 4) Calcula a probabilidade e marca os alertas
    pedidos["prob_negativa"] = modelo.predict_proba(pedidos[modelo.feature_names_in_])[:, 1].round(3)
    pedidos["alerta"] = (pedidos["prob_negativa"] >= a.corte).astype(int)

    # 5) Ordena do maior risco para o menor e salva
    pedidos = pedidos.sort_values("prob_negativa", ascending=False)
    colunas_saida = ["order_id", "prob_negativa", "alerta", "dias_atraso", "n_itens", "n_vendedores"]
    pedidos[colunas_saida].to_csv(a.saida, index=False)

    print(f"{len(pedidos)} pedidos analisados | {pedidos['alerta'].sum()} em alerta (prob >= {a.corte})")


if __name__ == "__main__":
    main()