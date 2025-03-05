import xlwings as xw
import datetime as dt
import pandas as pd
from datetime import date
import requests
import time
from user_agent import random_user
import stockvn as rpv

import gdown
from datetime import datetime
from datetime import timedelta
from bs4 import BeautifulSoup
import json
import html5lib

global head
head={"User-Agent":random_user()}





def dau_thau_thi_truong_mo():
    url_2='https://www.sbv.gov.vn/webcenter/portal/vi/menu/trangchu/hdtttt;jsessionid=fdqYz-QmDhTURXFqjuyFlAx7t8RTZ7a9_Gonrckd2qAAn5Zf90qy!993639228!-1111593773?_afrLoop=372261968090755'
    response2 = requests.get(url_2, allow_redirects=True)
    df4=pd.read_html(response2.text)[0][9:13]
    df= df4.iloc[:, :4]
    df.columns = df.iloc[0]
    df = df[1:]
    df['Lãi suất trúng thầu (%/năm)'] = (df['Lãi suất trúng thầu (%/năm)'].str.replace('%', '').astype(float))/100
    df = df.set_index(df.columns[0])
    return df

    

@xw.func()
def gia_vang_24money():
    url = 'https://api-finance-t19.24hmoney.vn/v1/ios/world-stock/all?device_id=web1723350utptenhuf4a5wu7r8rvgjjohs1qjvbq8468116'
    r = requests.get(url, head)
    data = r.json()['data']['gold_price']
    
    df = pd.DataFrame(data)[['Last', 'footer', 'text', 'Percent', 'change', 'symbol', 'extra_name']].assign(
        change=lambda x: pd.to_numeric(x['change'], errors='coerce'),
        Percent=lambda x: pd.to_numeric(x['Percent'].str.replace('%', '').str.strip(), errors='coerce') / 100,
        Last=lambda x: x['Last'].str.replace('N', '').str.strip()
    )
    
    return df

@xw.func()
def get_PE_PB_vnindex():
    url = 'https://s.cafef.vn/Ajax/PageNew/FinanceData/GetDataChartPE.ashx'
    r = requests.get(url, head)
    data = pd.DataFrame([r.json()['Data']['NowDataFinance'], r.json()['Data']['PastDataFinance']]).T
    data.columns = ['Hiện tại', 'Năm trước']
    
    return data.apply(pd.to_numeric, errors='coerce').reindex(['PE', 'PB', 'ROA', 'ROE', 'MaketCap'])




@xw.func()
def get_index_stock_world():
    url = 'https://api-finance-t19.24hmoney.vn/v1/ios/world-stock/all?device_id=web1723350utptenhuf4a5wu7r8rvgjjohs1qjvbq8468116'
    r = requests.get(url, head)
    
    df = pd.DataFrame(r.json()['data']['world_stock'])[['name', 'last_price', 'change_price', 'change_percent']]
    df['change_percent'] = pd.to_numeric(df['change_percent']) / 100
    
    return df

    


@xw.func()
def get_data_cp_vn30():
    todate = dt.datetime.now()
    N = 1
    
    while N <= 5:
        try:
            fromdate = todate - timedelta(days=N)
            url2 = f"https://s.cafef.vn/Ajax/PageNew/DataGDNN/GDNuocNgoai.ashx?TradeCenter=VN30&Date={fromdate.strftime('%Y-%m-%d')}"
            r2 = requests.get(url2, headers=head)
            df = pd.DataFrame(r2.json()['Data']['ListDataNN'])
            if not df.empty:
                return df[['Symbol']].sort_values(by='Symbol').set_index('Symbol')
            N += 2
        except:
            print('Lỗi phát sinh')


@xw.func()
def get_data_index():
    # Retrieve index data from the first URL
    re_vni_url = requests.get('https://banggia.cafef.vn/stockhandler.ashx?index=true')
    results_vni = json.loads(re_vni_url.text)
    
    # Update names for different indices
    results_vni[0]['name'] = 'HNX'
    results_vni[3]['name'] = 'UPCOM'
    
    # Convert the results to a DataFrame
    df = pd.DataFrame([results_vni[1], results_vni[4], results_vni[0], results_vni[2], results_vni[3]])
    
    # Convert necessary columns to numeric types
    columns_to_convert = ['change', 'percent']
    df['change'] = df['change'].apply(pd.to_numeric, errors='coerce')
    df['percent'] = df['percent'].apply(pd.to_numeric, errors='coerce')/100
    df['value'] = df['value'].str.replace(',', '').astype(float)
    
    # Define URLs for additional index data
    urls = {
        'vni': 'https://api-finance-t19.24hmoney.vn/v1/web/indices/trading-compare-daily?code=10',
        'vn30': 'https://api-finance-t19.24hmoney.vn/v1/web/indices/trading-compare-daily?code=11',
        'hnx': 'https://api-finance-t19.24hmoney.vn/v1/web/indices/trading-compare-daily?code=02',
        'upcom': 'https://api-finance-t19.24hmoney.vn/v1/web/indices/trading-compare-daily?code=03',
        'hn30': 'https://s.cafef.vn/Ajax/PageNew/DataHistory/PriceHistory.ashx?Symbol=HNX30-INDEX&StartDate=&EndDate=&PageIndex=1&PageSize=20'
    }
    
    # Fetch data from URLs
    data_frames = {}
    for key, url in urls.items():
        try:
            response = requests.get(url)
            response.raise_for_status()
            data = response.json()
            if key == 'hn30':
                df_key = pd.DataFrame(data['Data']['Data'])[['Ngay', 'GiaDieuChinh', 'GiaTriKhopLenh', 'GtThoaThuan']][1:2]
                df_key['value'] =df_key['GiaTriKhopLenh']
                df_key[['value']] = round(df_key[['value']].apply(pd.to_numeric, errors='coerce') / 1000000000,2)
            else:
                df_key = pd.DataFrame(data['data'][1]['data'])[-1:]
            data_frames[key] = df_key
        except (requests.RequestException, KeyError, ValueError) as e:
            print(f"Error fetching data for {key}: {e}")
            data_frames[key] = pd.DataFrame()  # Set to empty DataFrame on error
    
    # Combine the fetched data
    list_data = [
        ('VNINDEX', round(data_frames['vni']['total_value_traded'].values[0], 2)),
        ('VN30', round(data_frames['vn30']['total_value_traded'].values[0], 2)),
        ('HNX', round(data_frames['hnx']['total_value_traded'].values[0], 2)),
        ('HNX30', round(data_frames['hn30']['value'].values[0], 2)),
        ('UPCOM', round(data_frames['upcom']['total_value_traded'].values[0], 2))
    ]
    
    # Create DataFrame for the additional data
    df_t = pd.DataFrame(list_data, columns=['name', 'value_2'])
    df_t['value_2'] = pd.to_numeric(df_t['value_2'], errors='coerce')
    
    # Merge the dataframes and calculate the ratio
    result_df = pd.merge(df, df_t, on='name')
    result_df['value/value'] = ((result_df['value'] - result_df['value_2']) / result_df['value_2'])
    
    # Select relevant columns and set index
    data = result_df[['name', 'change', 'index', 'percent', 'volume', 'value', 'value/value']]
    return data.set_index('name')
    


@xw.func()
def info_company(symbol):
    df=rpv.get_info_cp(symbol)
    return df

#async_mode='threading'

@xw.func()
def momentum_ck(symbol):
    data=rpv.momentum_ck(symbol)
    return data


@xw.func()
def CW_info(symbol):

    pload1={}
    url2=f'https://finance.vietstock.vn/chung-khoan-phai-sinh/{symbol}/cw-tong-quan.htm'
    r=requests.get(url2,headers=head,data=pload1)
    soup=BeautifulSoup(r.text,'html.parser')
    ds=soup.find(class_="table table-hover")
    df=pd.read_html(ds.prettify())[0]
    df.rename(columns={0:'CW',1:f'{symbol[1:].upper()}'},inplace=True)
    return df

@xw.func()
def tinh_du_lieu_cp(symbol):
    # Thiết lập thời gian
    todate = dt.datetime.now()
    fromdate = todate - timedelta(days=200)
    fdate = fromdate.strftime('%Y-%m-%d')
    tdate = todate.strftime('%Y-%m-%d')

    # API URL và header
    url = f'https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:{symbol.upper()}~date:gte:{fdate}~date:lte:{tdate}&size=100000&page=1' 
    payload = {}

    # Gọi API và chuyển đổi dữ liệu
    r = requests.get(url, headers=head, data=payload)
    data = pd.DataFrame(r.json()['data'])

    # Đổi tên cột cho dễ hiểu và thêm cột 'volumn'
    data.rename(columns={
        'nmVolume': 'KLGD Khớp lệnh',
        'nmValue': 'GTGD Khớp lệnh',
        'ptVolume': 'KLGD Thỏa thuận',
        'ptValue': 'GTGD Thỏa thuận',
        'change': 'tăng/giảm',
        'pctChange': '% tăng/giảm'
    }, inplace=True)

    data['volumn'] = data['KLGD Khớp lệnh'] + data['KLGD Thỏa thuận']

    # Lấy các giá trị cần thiết từ DataFrame
    first_row = data.iloc[0]
    gia_close = pd.to_numeric(first_row['close'], errors='coerce')
    KL1000 = pd.to_numeric(first_row['volumn'], errors='coerce') / 1000
    BD_gia = pd.to_numeric(first_row['% tăng/giảm'], errors='coerce') / 100

    # Tính toán các giá trị cần thiết
    KLGD_KLTB21_mean = pd.to_numeric(data['volumn'].iloc[:22].mean(), errors='coerce')
    KLTB_KLTB21 = pd.to_numeric(first_row['volumn'], errors='coerce') / KLGD_KLTB21_mean

    close_mean_5 = pd.to_numeric(data['close'].iloc[:6].mean(), errors='coerce')
    close_mean_21 = pd.to_numeric(data['close'].iloc[:22].mean(), errors='coerce')
    gia_tbgia5 = close_mean_5 / close_mean_21

    KL_KLTB5_mean = pd.to_numeric(data['volumn'].iloc[:6].mean(), errors='coerce')
    KL_KLTB5 = pd.to_numeric(first_row['volumn'], errors='coerce') / KL_KLTB5_mean

    # Tính đỉnh và đáy của 60 ngày đầu
    close_60 = pd.to_numeric(data['close'].iloc[:60], errors='coerce')
    day2t = close_60.min()
    dinh2t = close_60.max()
    dinh_day = (dinh2t - day2t) / day2t
    giam_sdinh = (gia_close - dinh2t) / dinh2t
    tang_sday = (gia_close - day2t) / day2t

    return gia_close, KL1000, BD_gia, KLTB_KLTB21, gia_tbgia5, KL_KLTB5, dinh_day, day2t, dinh2t, tang_sday, giam_sdinh






@xw.func()
def get_price_historical_vnd(symbol,fromdate,todate):
    fromdate, todate = pd.to_datetime(fromdate, dayfirst=True), pd.to_datetime(todate, dayfirst=True)
    fdate, tdate=fromdate.strftime('%Y-%m-%d'), todate.strftime('%Y-%m-%d')
    url=f'https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:{symbol.upper()}~date:gte:{fdate}~date:lte:{tdate}&size=100000&page=1' 
    
    payload={}
    r=requests.get(url,headers=head,data=payload)
    df=pd.DataFrame(r.json()['data'])
    data=df[['code','date','open','high','low','close','nmVolume','nmValue','ptVolume', 'ptValue','change','pctChange']].copy()
    data.rename(columns={'nmVolume':'KLGD Khớp lệnh','nmValue':'GTGD Khớp lệnh','ptVolume':'KLGD Thỏa thuận','ptValue':'GTGD Thỏa thuận','change':'tăng/giảm','pctChange':'% tăng/giảm'}, inplace=True)
    #columns_to_convert = ['open', 'high', 'low', 'close']
    #data[columns_to_convert] = data[columns_to_convert].apply(lambda x: x * 1000)
    data['% tăng/giảm'] = data['% tăng/giảm'].astype(str) + '%'
    return data







@xw.func
def key_id(code):
    day=rpv.key_id(str(code))
    return day

@xw.func()
def giao_dich_tu_doanh(symbol,fromdate,todate):
    fromdate = pd.to_datetime(fromdate)
    fdate = fromdate.strftime('%d/%m/%Y')
    todate = pd.to_datetime(todate)
    tdate = todate.strftime('%d/%m/%Y')
    df,data=rpv.get_proprietary_history_cafef(symbol.upper(),fdate,tdate)
    return data



@xw.func()
def report_finance_vnd(symbol,types,year_f,timely):
    symbol,types, timely=symbol.upper(), types.upper(), timely.upper()
    year_f=int(year_f)
    data=rpv.report_finance_vnd(symbol,types,year_f,timely)
    return data




@xw.func()
def laisuat_vietstock(fromdate,todate):
    fromdate=pd.to_datetime(fromdate, dayfirst=True)
    todate=pd.to_datetime(todate, dayfirst=True)
    data=rpv.laisuat_vietstock(fromdate,todate)
    return data

@xw.func()
def getCPI_vietstock(fromdate,todate):
    fromdate=pd.to_datetime(fromdate, dayfirst=True)
    todate=pd.to_datetime(todate, dayfirst=True)
    data=rpv.getCPI_vietstock(fromdate,todate)
    return data

@xw.func()
def solieu_sanxuat_congnghiep(fromdate,todate):
    fromdate=pd.to_datetime(fromdate, dayfirst=True)
    todate=pd.to_datetime(todate, dayfirst=True)
    data=rpv.solieu_sanxuat_congnghiep(fromdate,todate)
    return data

@xw.func()
def solieu_banle_vietstock(fromdate,todate):
    fromdate=pd.to_datetime(fromdate, dayfirst=True)
    todate=pd.to_datetime(todate, dayfirst=True)
    data=rpv.solieu_banle_vietstock(fromdate,todate)
    return data

@xw.func()
def solieu_XNK_vietstock(fromdate,todate):
    fromdate=pd.to_datetime(fromdate, dayfirst=True)
    todate=pd.to_datetime(todate, dayfirst=True)
    data=rpv.solieu_XNK_vietstock(fromdate,todate)
    return data

@xw.func()
def solieu_FDI_vietstock(fromdate,todate):
    fromdate=pd.to_datetime(fromdate, dayfirst=True)
    todate=pd.to_datetime(todate, dayfirst=True)
    data=rpv.solieu_FDI_vietstock(fromdate,todate)
    return data   

@xw.func()
def tygia_vietstock(fromdate,todate):
    fromdate=pd.to_datetime(fromdate, dayfirst=True)
    todate=pd.to_datetime(todate, dayfirst=True)
    data=rpv.tygia_vietstock(fromdate,todate)
    return data 


@xw.func()
def solieu_tindung_vietstock(fromdate,todate):
    fromdate=pd.to_datetime(fromdate, dayfirst=True)
    todate=pd.to_datetime(todate, dayfirst=True)
    data=rpv.solieu_tindung_vietstock(fromdate,todate)
    return data 


@xw.func()
def solieu_GDP_vietstock(fromyear,fromQ,toyear,toQ):
    fromyear=int(fromyear)
    fromQ=int(fromQ)
    toyear=int(toyear)
    toQ=int(toQ)
    data=rpv.solieu_GDP_vietstock(fromyear,fromQ,toyear,toQ)
    return data 



if __name__ == "__main__":
    xw.Book("func.xlsm").set_mock_caller()
    main()