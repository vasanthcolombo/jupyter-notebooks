import pandas as pd
import yfinance as yf

def get_historical_px(ticker:str, start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
    data = yf.download(ticker, start=start_date, end=end_date + pd.Timedelta(days=1), progress=False, auto_adjust=True)
    assert data is not None, "Failed to retrieve historical price data"
    data.columns = data.columns.get_level_values(0)
    data.columns.name = None
    data.index = pd.to_datetime(data.index, errors='coerce').map(lambda x: x.date() if isinstance(x, pd.Timestamp) else x)
    return data

def get_historical_close_px(ticker:str, start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
    data = get_historical_px(ticker, start_date, end_date)
    data = data[['Close']].round(2)
    return data.rename(columns={'Close': 'close'})

def get_px_for_date(date: pd.Timestamp, price_data: pd.DataFrame) -> float:
    available_dates = price_data.index
    closest_date = min(available_dates, key=lambda x: abs(pd.to_datetime(x) - date))
    return round(price_data.loc[closest_date], 2)