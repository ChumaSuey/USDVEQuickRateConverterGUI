import requests

def get_bcv_monthly_data():
    # Endpoint for official historical data
    url = "https://ve.dolarapi.com/v1/historicos/dolares/oficial"
    
    print("Connecting to DolarApi to fetch monthly history...")
    print("-" * 60)
    
    try:
        response = requests.get(url)
        response.raise_for_status()
        historical_data = response.json()
        
        # 1. Clean and normalize the records
        cleaned_records = []
        for record in historical_data:
            date = record.get("fecha", "")[:10]
            # Handle potential variable field naming variations
            rate = record.get("promedio") or record.get("precio") or record.get("valor")
            
            if date and rate:
                cleaned_records.append({
                    "date": date,
                    "rate": float(rate)
                })
        
        # 2. Sort chronologically (oldest to newest)
        cleaned_records.sort(key=lambda x: x["date"])
        
        # 3. Slice the last 30 active business entries
        monthly_records = cleaned_records[-30:]
        
        if not monthly_records:
            print("No valid records found in the API response.")
            return
        
        # 4. Print the daily breakdown
        for record in monthly_records:
            print(f"Fecha: {record['date']} | Tasa BCV: {record['rate']:,.4f} Bs.")
            
        print("-" * 60)
        
        # 5. Math Layer (Perfect foundation for GUI metrics)
        rates_only = [r["rate"] for r in monthly_records]
        min_rate = min(rates_only)
        max_rate = max(rates_only)
        avg_rate = sum(rates_only) / len(rates_only)
        
        # Calculate percentage variance over the last 30 points
        initial_rate = rates_only[0]
        final_rate = rates_only[-1]
        pct_change = ((final_rate - initial_rate) / initial_rate) * 100
        
        print(f"📊 SUMMARY FOR THE LAST 30 ACTIVE TRADING DAYS:")
        print(f"  • Mínimo (Low):   {min_rate:,.4f} Bs.")
        print(f"  • Máximo (High):  {max_rate:,.4f} Bs.")
        print(f"  • Promedio (Avg): {avg_rate:,.4f} Bs.")
        print(f"  • Variación (%):  {pct_change:+.2f}% (desde {monthly_records[0]['date']})")
        print("-" * 60)
        
        return monthly_records

    except requests.exceptions.RequestException as e:
        print(f"❌ Connection error: {e}")
        return []
    except ValueError as e:
        print(f"❌ Data processing error (could not parse rates as floats): {e}")
        return []

if __name__ == "__main__":
    get_bcv_monthly_data()