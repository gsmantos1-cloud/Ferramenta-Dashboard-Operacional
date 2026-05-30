import sqlite3
import os
from datetime import date, timedelta

import holidays

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "pedidos.db")

LIMITES = {
    "Normal": [
        (0, 3,    "NO PRAZO"),
        (4, 5,    "Atraso leve"),
        (6, 6,    "Atraso moderado"),
        (7, 9999, "Atraso crítico"),
    ],
    "Personalizado": [
        (0, 4,    "NO PRAZO"),
        (5, 6,    "Atraso leve"),
        (7, 7,    "Atraso moderado"),
        (8, 9999, "Atraso crítico"),
    ],
    "Internacional": [
        (0, 15,   "NO PRAZO"),
        (16, 20,  "Atraso leve"),
        (21, 25,  "Atraso moderado"),
        (26, 9999,"Atraso crítico"),
    ],
}


def calcular_dias_uteis(data_inicio: date, data_fim: date) -> int:
    feriados = holidays.Brazil(years=range(data_inicio.year, data_fim.year + 2))
    count, atual = 0, data_inicio
    while atual <= data_fim:
        if atual.weekday() < 5 and atual not in feriados:
            count += 1
        atual += timedelta(days=1)
    return count


def determinar_status(dias: int, categoria: str) -> str:
    for minimo, maximo, status in LIMITES.get(categoria, LIMITES["Normal"]):
        if minimo <= dias <= maximo:
            return status
    return "Atraso crítico"


def main():
    if not os.path.exists(DB_PATH):
        print("Banco de dados não encontrado. Inicie o app primeiro.")
        return

    hoje = date.today()
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        pedidos = conn.execute(
            "SELECT id, numero, data_pedido, categoria, status FROM pedidos WHERE ativo = 1"
        ).fetchall()

        atualizados = 0
        for p in pedidos:
            dp   = date.fromisoformat(p["data_pedido"])
            dias = calcular_dias_uteis(dp, hoje)
            novo = determinar_status(dias, p["categoria"])
            if novo != p["status"]:
                conn.execute("UPDATE pedidos SET status = ? WHERE id = ?", (novo, p["id"]))
                print(f"[ATUALIZADO] Pedido {p['numero']}: {p['status']} → {novo} ({dias} dias úteis)")
                atualizados += 1
            else:
                print(f"[OK]         Pedido {p['numero']}: {p['status']} ({dias} dias úteis)")

    print(f"\nFinalizado: {atualizados} pedido(s) atualizado(s) de {len(pedidos)} total.")


if __name__ == "__main__":
    main()
