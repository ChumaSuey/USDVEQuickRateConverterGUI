"""
BCV Dollar Tracker - CLI (Prototipo experimental)

Uso:
    python mainCLI.py [-c {usd,eur}] [-r] [-d N] [--csv ARCHIVO]

Args implementados:
    -c, --currency  {usd,eur}   Moneda a consultar (default: usd)
    -r, --raw                   Mostrar solo días hábiles reales (sin Smart Fill)
    -d, --days      N           Cantidad de días corridos a mostrar (default: 30)
    --csv           ARCHIVO     Exportar resultados a CSV
"""


import argparse
import csv

from backend import format_currency_ve, fetch_all_bcv_history

# La llamada a la API y la lógica de relleno de calendario (Smart Fill)
# están centralizadas en backend.py → fetch_all_bcv_history()


def get_bcv_monthly_data(currency="dolares", days=30, raw_only=False, csv_file=None):
    currency_label = "USD" if currency == "dolares" else "EUR"

    print(f"Conectando a DolarApi para obtener historial ({currency_label})...")
    print("-" * 60)

    all_records = fetch_all_bcv_history(currency=currency)

    if not all_records:
        print("No se encontraron registros o no hay conexión a internet.")
        return []

    if raw_only:
        all_records = [r for r in all_records if not r.get("is_filled")]

    records = all_records[-days:]

    for record in records:
        tag = " (Feriado/FDS)" if record.get("is_filled") else ""
        print(f"Fecha: {record['date']} | Tasa BCV: {format_currency_ve(record['rate'])} Bs.{tag}")

    print("-" * 60)

    rates_only = [r["rate"] for r in records]
    min_rate = min(rates_only)
    max_rate = max(rates_only)
    avg_rate = sum(rates_only) / len(rates_only)

    initial_rate = rates_only[0]
    final_rate = rates_only[-1]
    pct_change = ((final_rate - initial_rate) / initial_rate) * 100

    print(f"RESUMEN DEL PERIODO (últimos {days} días corridos con Smart Fill):")
    print(f"  - Mínimo:    {format_currency_ve(min_rate)} Bs.")
    print(f"  - Máximo:    {format_currency_ve(max_rate)} Bs.")
    print(f"  - Promedio:  {format_currency_ve(avg_rate)} Bs.")
    print(f"  - Variación: {pct_change:+.2f}% (desde {records[0]['date']})")
    print("-" * 60)

    if csv_file:
        try:
            with open(csv_file, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Fecha", "Tasa BCV (Bs.)"])
                for record in records:
                    writer.writerow([record["date"], format_currency_ve(record["rate"])])
            print(f"Exportado a CSV: {csv_file}")
        except Exception as e:
            print(f"Error al exportar CSV: {e}")

    return records


def parse_args():
    parser = argparse.ArgumentParser(
        description="BCV Dollar Tracker - Consulta de tasas oficiales BCV (prototipo CLI)"
    )
    parser.add_argument(
        "-c", "--currency", choices=["usd", "eur"], default="usd",
        help="Moneda a consultar (default: usd)"
    )
    parser.add_argument(
        "-r", "--raw", action="store_true",
        help="Mostrar solo días hábiles reales (sin Smart Fill de feriados/fines de semana)"
    )
    parser.add_argument(
        "-d", "--days", type=int, default=30,
        help="Cantidad de días corridos a mostrar (default: 30)"
    )
    parser.add_argument(
        "--csv", type=str, default=None, metavar="ARCHIVO",
        help="Exportar resultados a archivo CSV"
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    currency = "dolares" if args.currency == "usd" else "euros"
    get_bcv_monthly_data(currency=currency, days=args.days, raw_only=args.raw, csv_file=args.csv)