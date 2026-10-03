import pandas as pd
import numpy as np
import random
from pathlib import Path

def create_messy_dataset():
    raw_path = Path("data/raw_superstore.csv")
    if not raw_path.exists():
        print("Raw file not found at data/raw_superstore.csv")
        return
    
    df = pd.read_csv(raw_path, encoding='windows-1252')
    print(f"Original real Superstore dataset loaded: {df.shape[0]} rows, {df.shape[1]} columns")

    # Sample ~3,000 rows for quick testing and demonstration
    df_sample = df.sample(n=3000, random_state=42).copy()
    
    # 1. Inject duplicate rows (simulating duplicate batch dumps)
    duplicates = df_sample.sample(n=120, random_state=101)
    df_messy = pd.concat([df_sample, duplicates], ignore_index=True)
    
    # 2. Messy Text & Rogue Whitespace
    def mess_up_text(val):
        if pd.isna(val):
            return val
        rand = random.random()
        val_str = str(val)
        if rand < 0.2:
            return f"  {val_str}  "
        elif rand < 0.4:
            return val_str.lower()
        elif rand < 0.6:
            return val_str.upper()
        return val_str

    for col in ['Ship Mode', 'Segment', 'Category', 'Region']:
        df_messy[col] = df_messy[col].apply(mess_up_text)

    # 3. Dirty Currency & Numeric Formatting (Sales, Profit)
    def mess_up_currency(val):
        if pd.isna(val):
            return val
        rand = random.random()
        try:
            val_float = float(val)
        except Exception:
            return val
        if rand < 0.25:
            return f"$ {val_float:,.2f}"
        elif rand < 0.45:
            return f"{val_float:,.2f} USD"
        elif rand < 0.55:
            return f"  ${val_float:.2f}  "
        elif rand < 0.60:
            return "N/A"
        elif rand < 0.65:
            return "-"
        return val_float

    df_messy['Sales'] = df_messy['Sales'].apply(mess_up_currency)
    df_messy['Profit'] = df_messy['Profit'].apply(mess_up_currency)

    # 4. Inconsistent Date Formats in 'Order Date'
    dates = pd.to_datetime(df_messy['Order Date'], format='mixed', errors='coerce')
    messy_dates = []
    for d in dates:
        if pd.isna(d):
            messy_dates.append("Invalid Date")
            continue
        rand = random.random()
        if rand < 0.3:
            messy_dates.append(d.strftime("%Y-%m-%d"))  # ISO
        elif rand < 0.6:
            messy_dates.append(d.strftime("%m/%d/%Y"))  # US format
        elif rand < 0.8:
            messy_dates.append(d.strftime("%d-%b-%Y"))  # e.g. 15-Nov-2023
        elif rand < 0.95:
            messy_dates.append(d.strftime("%Y/%m/%d %H:%M:%S"))
        else:
            messy_dates.append("Pending")  # Dirty rogue text
            
    df_messy['Order Date'] = messy_dates

    # 5. Inject Missing Values (Nulls)
    np.random.seed(42)
    mask_customer = np.random.rand(len(df_messy)) < 0.04
    df_messy.loc[mask_customer, 'Customer Name'] = np.nan
    
    mask_postal = np.random.rand(len(df_messy)) < 0.08
    df_messy.loc[mask_postal, 'Postal Code'] = np.nan

    # 6. Save as both messy CSV and messy Excel
    csv_out = Path("data/messy_sales_data.csv")
    xlsx_out = Path("data/messy_sales_data.xlsx")
    
    df_messy.to_csv(csv_out, index=False)
    df_messy.to_excel(xlsx_out, index=False, engine='openpyxl')
    
    print(f"Messy dataset generated successfully!")
    print(f"-> CSV: {csv_out.resolve()} ({df_messy.shape[0]} rows)")
    print(f"-> Excel: {xlsx_out.resolve()} ({df_messy.shape[0]} rows)")

if __name__ == "__main__":
    create_messy_dataset()
