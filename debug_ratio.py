from vnstock import Vnstock
import pandas as pd

try:
    print("Fetching ratio for HPG...")
    stock = Vnstock().stock(symbol='HPG', source='VCI')
    df_ratio = stock.finance.ratio(period='quarter', lang='vi')
    
    if df_ratio is not None and not df_ratio.empty:
        print("Columns found:", df_ratio.columns.tolist())
        print("First row:", df_ratio.iloc[0].to_dict())
    else:
        print("Ratio dataframe is empty.")
        
except Exception as e:
    print(f"Error: {e}")
