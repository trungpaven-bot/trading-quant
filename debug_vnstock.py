from vnstock import stock_historical_data
from datetime import datetime, timedelta

try:
    print("Test fetching HPG data...")
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d')
    
    df = stock_historical_data(symbol='HPG', start_date=start_date, end_date=end_date, resolution='1D', type='stock')
    
    if df.empty:
        print("Dataframe is empty!")
    else:
        print("Columns:", df.columns.tolist())
        print(df.head())
        print("Last row close:", df.iloc[-1]['close'])

except Exception as e:
    print(f"Error fetching: {e}")
    