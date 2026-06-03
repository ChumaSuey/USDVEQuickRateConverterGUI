import requests
import datetime

# ==========================================
# FUNCIÓN AUXILIAR DE FORMATO CONTABLE
# ==========================================
def format_currency_ve(value):
    """
    Formatea un float al estándar contable venezolano/español:
    Puntos para miles, comas para decimales, redondeado exactamente a 2 dígitos.
    Ejemplo: 400.545645 -> "400,55"
             1234.56    -> "1.234,56"
    """
    if value is None:
        return "--,--"
    s = f"{value:,.2f}"
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


# ==========================================
# CAPA DE DATOS (Con Relleno Inteligente de Calendario)
# ==========================================
def fetch_all_bcv_history(currency="dolares"):
    """Busca el histórico de la base de datos de la API y aplica Smart Fill a las brechas."""
    url = f"https://ve.dolarapi.com/v1/historicos/{currency}/oficial"
    try:
        response = requests.get(url, timeout=12)
        response.raise_for_status()
        historical_data = response.json()
        
        cleaned_records = []
        for record in historical_data:
            date = record.get("fecha", "")[:10]
            rate = record.get("promedio") or record.get("precio") or record.get("valor")
            if date and rate:
                cleaned_records.append({"date": date, "rate": float(rate)})
        
        if not cleaned_records:
            return []
            
        cleaned_records.sort(key=lambda x: x["date"])
        
        raw_map = {r["date"]: r["rate"] for r in cleaned_records}
        start_date = datetime.date.fromisoformat(cleaned_records[0]["date"])
        end_date = datetime.date.fromisoformat(cleaned_records[-1]["date"])
        
        filled_records = []
        current_date = start_date
        last_known_rate = cleaned_records[0]["rate"]
        
        while current_date <= end_date:
            date_str = current_date.isoformat()
            if date_str in raw_map:
                last_known_rate = raw_map[date_str]
                filled_records.append({"date": date_str, "rate": last_known_rate, "is_filled": False})
            else:
                filled_records.append({"date": date_str, "rate": last_known_rate, "is_filled": True})
            current_date += datetime.timedelta(days=1)
            
        return filled_records
    except requests.exceptions.RequestException as e:
        print(f"Error de red (Sin Internet o API caída): {e}")
        return []
    except Exception as e:
        print(f"Error general de procesamiento: {e}")
        return []