import yfinance as yf
from pykrx import bond
import pandas as pd

def market_data_generate():
    국고채10년 = bond.get_otc_treasury_yields("20230104","20240101","국고채10년").reset_index()
    국고채3년 = bond.get_otc_treasury_yields("20230104","20240101","국고채3년").reset_index()

    국고채3년['종목명'] = "국채3년"
    국고채10년['종목명'] = "국채10년"

    treasury_bond = pd.concat([국고채3년,국고채10년])
    treasury_bond.to_csv("./data/국채데이터.csv",index=False)

    snp_df = yf.download('SPY', start='2023-01-02').reset_index()
    nasdaq_df = yf.download('^IXIC', start='2023-01-02').reset_index()
    kospi_df = yf.download('^KS11', start='2023-01-02').reset_index()
    kosdaq_df = yf.download('^KQ11', start='2023-01-02').reset_index()

    domestic_df =  pd.DataFrame()
    domestic_df['Date'] = kospi_df['Date']
    domestic_df['Kosdaq'] = kosdaq_df['Close']
    domestic_df['Kospi'] = kospi_df['Close']

    nondomestic_df =  pd.DataFrame()
    nondomestic_df['Date'] = snp_df['Date']
    nondomestic_df['S&P500'] = snp_df['Close']
    nondomestic_df['Nasdaq'] = nasdaq_df['Close']

    domestic_df.to_csv("./data/국내지수.csv",index=False)
    nondomestic_df.to_csv("./data/해외지수.csv",index=False)

    wti_df = yf.download('CL=F', '2023-01-02')
    wti_df.reset_index(inplace=True,drop=False)

    wti_df.to_csv('./data/WTI.csv',index=False)

    stock_price = pd.DataFrame()
    target = yf.download("005930.KS","2023-01-01").reset_index()
    stock_price['Date']= target['Date']
    stock_price['삼성전자'] = target['Close']
    stock_price['한국금융지주'] = yf.download("071050.KS","2023-01-01").reset_index()['Close']
    stock_price['SK하이닉스'] = yf.download("000660.KS","2023-01-01").reset_index()['Close']
    stock_price['LG에너지솔루션'] = yf.download("373220.KS","2023-01-01").reset_index()['Close']
    stock_price.to_csv("./data/stock_price.csv",index=False)


market_data_generate()
