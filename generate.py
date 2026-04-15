import pandas as pd
import json
import re
from datetime import datetime, timedelta

def excel_date(n):
    return (datetime(1899, 12, 30) + timedelta(days=int(n))).strftime('%Y-%m-%d')

def read_margin(xl):
    df = xl.parse('Margin Borrow Rates')
    df.columns = ['Date','Binance','OKX','Bybit','KuCoin','Gate_IO']
    result = []
    for _, row in df.iterrows():
        try:
            d = pd.to_datetime(row['Date']).strftime('%Y-%m-%d')
            if pd.notna(row['Binance']):
                result.append({
                    'date': d,
                    'Binance': round(float(row['Binance'])*100, 4),
                    'OKX': round(float(row['OKX'])*100, 4) if pd.notna(row['OKX']) else None,
                    'Bybit': round(float(row['Bybit'])*100, 4) if pd.notna(row['Bybit']) else None,
                    'KuCoin': round(float(row['KuCoin'])*100, 4) if pd.notna(row['KuCoin']) else None,
                    'Gate_IO': round(float(row['Gate_IO'])*100, 4) if pd.notna(row['Gate_IO']) else None,
                })
        except: pass
    return result

def read_loan(xl):
    df = xl.parse('Daily Data', header=None)
    data_rows = df.iloc[3:].reset_index(drop=True)
    def extract(col_date, col_flex, label):
        dates = pd.to_datetime(data_rows.iloc[:, col_date], errors='coerce')
        vals = pd.to_numeric(data_rows.iloc[:, col_flex], errors='coerce')
        d = pd.DataFrame({'date': dates, label: vals}).dropna()
        d['date'] = d['date'].dt.normalize()
        return d
    from functools import reduce
    dfs = [extract(0,15,'Binance'), extract(8,33,'OKX'), extract(44,51,'ByBit'), extract(62,69,'Gate_IO')]
    merged = reduce(lambda a,b: pd.merge(a,b,on='date',how='outer'), dfs).sort_values('date')
    result = []
    for _, row in merged.iterrows():
        entry = {'date': row['date'].strftime('%Y-%m-%d')}
        for col in ['Binance','OKX','ByBit','Gate_IO']:
            val = row.get(col)
            entry[col] = round(float(val)*100, 4) if pd.notna(val) else None
        result.append(entry)
    return result

def read_term(xl):
    sheets = [s for s in xl.sheet_names if '.' in s and '2026' in s]
    term_data = {}
    exchanges = ['ByBit','OKX','Gate I.O','XT','Bitget','Binance*','KuCoin**']
    for sheet in sheets:
        try:
            df = xl.parse(sheet, header=None)
            entry = {}
            for i, row in df.iterrows():
                exch = str(row.iloc[0]).strip()
                if exch in exchanges:
                    d = {}
                    for j, tk in enumerate(['t7','t14','t30','t60','t90','t180']):
                        v = row.iloc[j+1]
                        d[tk] = round(float(v)*100,4) if pd.notna(v) and isinstance(v,(int,float)) and v>0 else None
                    flex = row.iloc[7]
                    d['flexible'] = round(float(flex)*100,4) if pd.notna(flex) and isinstance(flex,(int,float)) else None
                    spread = row.iloc[10] if len(row)>10 else None
                    d['spread'] = round(float(spread)*100,4) if pd.notna(spread) and isinstance(spread,(int,float)) else None
                    entry[exch] = d
            if entry:
                term_data[sheet] = entry
        except: pass
    return term_data

xl = pd.ExcelFile('Daily Earn-Borrow Rates.xlsx')
margin = read_margin(xl)
loan = read_loan(xl)
term = read_term(xl)

with open('crypto_loan_analysis (2).html', 'r', encoding='utf-8') as f:
    html = f.read()

html = re.sub(r'const DATA = {.*?};', f'const DATA = {json.dumps({"margin":margin,"loan":loan})};', html, flags=re.DOTALL)
html = re.sub(r'const TERM_ALL = {.*?};', f'const TERM_ALL = {json.dumps(term)};', html, flags=re.DOTALL)

with open('crypto_loan_analysis (2).html', 'w', encoding='utf-8') as f:
    f.write(html)

print('Dashboard güncellendi.')
