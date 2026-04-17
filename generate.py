"""
Crypto Loan Dashboard - Veri Güncelleme Script'i
=================================================
Bu script Excel dosyasını okur ve HTML dashboard'u otomatik günceller.

Güncellenen şeyler:
- DATA (margin, loan, btc, earn)
- TENOR_DATA (t7, t14, t30, t60, t90, t180 günlük veriler)
- TERM_ALL (tarihli borsa verileri)
- LAST_DATE
- "Son güncelleme" metni
- "Tümü" dropdown metni

Excel yapısı beklentileri:
- 'Margin Borrow Rates' sheet: Date, Binance, OKX, Bybit, KuCoin, Gate I.O, BTC Price $
- 'Daily Data' sheet: Borsa bazlı günlük Earn/Borrow oranları
- Günlük tarihli sheet'ler: 'GG.AA.2026' veya 'G.A.2026' formatında (term verileri)
"""

import pandas as pd
import json
import re
from datetime import datetime
from functools import reduce


# ==================== VERİ OKUMA FONKSİYONLARI ====================

def read_margin_and_btc(xl):
    """
    'Margin Borrow Rates' sheet'inden margin rate + BTC price verisini çeker.
    Döndürdüğü: (margin_listesi, btc_listesi)
    """
    df = xl.parse('Margin Borrow Rates')
    # Esneklik için: BTC sütunu varsa al, yoksa yok say
    has_btc = df.shape[1] >= 7
    df.columns = ['Date','Binance','OKX','Bybit','KuCoin','Gate_IO','BTC'] if has_btc \
                 else ['Date','Binance','OKX','Bybit','KuCoin','Gate_IO']

    margin_result = []
    btc_result = []
    for _, row in df.iterrows():
        try:
            d = pd.to_datetime(row['Date']).strftime('%Y-%m-%d')
            # Margin: Binance doluysa satırı al
            if pd.notna(row['Binance']):
                margin_result.append({
                    'date': d,
                    'Binance': round(float(row['Binance'])*100, 4),
                    'OKX': round(float(row['OKX'])*100, 4) if pd.notna(row['OKX']) else None,
                    'Bybit': round(float(row['Bybit'])*100, 4) if pd.notna(row['Bybit']) else None,
                    'KuCoin': round(float(row['KuCoin'])*100, 4) if pd.notna(row['KuCoin']) else None,
                    'Gate_IO': round(float(row['Gate_IO'])*100, 4) if pd.notna(row['Gate_IO']) else None,
                })
                # BTC: Excel'de bin olarak girilmiş (74.810 gibi), ×1000 ile ham dolara çevir
                if has_btc:
                    btc_val = row['BTC']
                    if pd.notna(btc_val):
                        btc_result.append({'date': d, 'btc': round(float(btc_val)*1000, 2)})
                    else:
                        btc_result.append({'date': d, 'btc': None})
        except Exception:
            pass
    return margin_result, btc_result


def read_loan_and_earn(xl):
    """
    'Daily Data' sheet'inden loan (Flexible Borrow) ve earn (Flexible Earn) verilerini çeker.
    Kolon indeksleri:
      - Tarih: Binance=0, OKX=8, ByBit=44, Gate=62
      - Earn Flex: Binance=7, OKX=25, ByBit=43, Gate=61
      - Borrow Flex: Binance=15, OKX=33, ByBit=51, Gate=69
    Döndürdüğü: (loan_listesi, earn_listesi)
    """
    df = xl.parse('Daily Data', header=None)
    data_rows = df.iloc[3:].reset_index(drop=True)

    def extract(col_date, col_flex, label):
        dates = pd.to_datetime(data_rows.iloc[:, col_date], errors='coerce')
        vals = pd.to_numeric(data_rows.iloc[:, col_flex], errors='coerce')
        d = pd.DataFrame({'date': dates, label: vals}).dropna()
        d['date'] = d['date'].dt.normalize()
        return d

    # Loan (Borrow Flexible)
    loan_dfs = [
        extract(0, 15, 'Binance'),
        extract(8, 33, 'OKX'),
        extract(44, 51, 'ByBit'),
        extract(62, 69, 'Gate_IO'),
    ]
    loan_merged = reduce(lambda a,b: pd.merge(a,b,on='date',how='outer'), loan_dfs).sort_values('date')

    # Earn (Earn Flexible)
    earn_dfs = [
        extract(0, 7,  'Binance'),
        extract(8, 25, 'OKX'),
        extract(44, 43, 'ByBit'),
        extract(62, 61, 'Gate_IO'),
    ]
    earn_merged = reduce(lambda a,b: pd.merge(a,b,on='date',how='outer'), earn_dfs).sort_values('date')

    def to_records(merged):
        result = []
        for _, row in merged.iterrows():
            entry = {'date': row['date'].strftime('%Y-%m-%d')}
            for col in ['Binance','OKX','ByBit','Gate_IO']:
                val = row.get(col)
                entry[col] = round(float(val)*100, 4) if pd.notna(val) else None
            result.append(entry)
        return result

    return to_records(loan_merged), to_records(earn_merged)


def read_term_sheets(xl):
    """
    Günlük tarihli sheet'lerden (ör. '17.04.2026') vadeli (term) verileri çeker.
    Döndürdüğü: {tarih_str: {borsa: {t7, t14, t30, t60, t90, t180, flexible, spread}}}
    """
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
                        d[tk] = round(float(v)*100, 4) if pd.notna(v) and isinstance(v,(int,float)) and v>0 else None
                    flex = row.iloc[7]
                    d['flexible'] = round(float(flex)*100, 4) if pd.notna(flex) and isinstance(flex,(int,float)) else None
                    spread = row.iloc[10] if len(row)>10 else None
                    d['spread'] = round(float(spread)*100, 4) if pd.notna(spread) and isinstance(spread,(int,float)) else None
                    entry[exch] = d
            if entry:
                term_data[sheet] = entry
        except Exception:
            pass
    return term_data


def parse_sheet_date(s):
    """'17.04.2026' veya '7.4.2026' gibi sheet adlarını datetime objesine çevirir."""
    try:
        parts = s.split('.')
        return datetime(int(parts[2]), int(parts[1]), int(parts[0]))
    except Exception:
        return datetime(1900, 1, 1)


def build_tenor_data(term_data):
    """
    TERM_ALL yapısından TENOR_DATA yapısına dönüştürür.
    Tarihler gerçek tarih sırasına göre sıralanır (string değil).
    """
    tenor_keys = ['t7','t14','t30','t60','t90','t180']
    exchanges = ['ByBit','OKX','Gate I.O','XT','Bitget','Binance*','KuCoin**']
    tenor_data = {tk: [] for tk in tenor_keys}

    sorted_dates = sorted(term_data.keys(), key=parse_sheet_date)

    for tk in tenor_keys:
        for date_str in sorted_dates:
            entry = {'date': date_str}
            for exch in exchanges:
                val = term_data.get(date_str, {}).get(exch, {}).get(tk)
                entry[exch] = val
            tenor_data[tk].append(entry)

    return tenor_data


# ==================== HTML GÜNCELLEME ====================

def update_html(html, data_obj, tenor_data, term_data, last_date_raw, last_date_display):
    """HTML içindeki tüm veri bloklarını ve metinleri günceller."""

    # 1) const DATA = {...}
    html = re.sub(
        r'const DATA = \{.*?\};',
        lambda m: f'const DATA = {json.dumps(data_obj)};',
        html, flags=re.DOTALL
    )

    # 2) const TENOR_DATA = {...}
    html = re.sub(
        r'const TENOR_DATA = \{.*?\};',
        lambda m: f'const TENOR_DATA = {json.dumps(tenor_data)};',
        html, flags=re.DOTALL
    )

    # 3) const TERM_ALL = {...}
    html = re.sub(
        r'const TERM_ALL = \{.*?\};',
        lambda m: f'const TERM_ALL = {json.dumps(term_data)};',
        html, flags=re.DOTALL
    )

    # 4) const LAST_DATE = "...";  (TERM_ALL key'i formatında)
    html = re.sub(
        r'const LAST_DATE = "[^"]*";',
        f'const LAST_DATE = "{last_date_raw}";',
        html
    )

    # 5) "Son güncelleme: DD.MM.YYYY" metni
    html = re.sub(
        r'Son güncelleme: \d{1,2}\.\d{1,2}\.\d{4}',
        f'Son güncelleme: {last_date_display}',
        html
    )

    # 6) "Tümü (01 Oca – DD.MM.YYYY)" dropdown metni
    html = re.sub(
        r'Tümü \(01 Oca – \d{1,2}\.\d{1,2}\.\d{4}\)',
        f'Tümü (01 Oca – {last_date_display})',
        html
    )

    return html


# ==================== ANA AKIŞ ====================

def main():
    excel_path = 'Daily Earn-Borrow Rates.xlsx'
    html_path = 'crypto_loan_analysis (2).html'

    print(f'Excel okunuyor: {excel_path}')
    xl = pd.ExcelFile(excel_path)

    print('Margin & BTC verisi okunuyor...')
    margin, btc = read_margin_and_btc(xl)

    print('Loan & Earn verisi okunuyor...')
    loan, earn = read_loan_and_earn(xl)

    print('Term (günlük) sheetleri okunuyor...')
    term = read_term_sheets(xl)

    print('TENOR_DATA yapısı oluşturuluyor...')
    tenor = build_tenor_data(term)

    # En son tarih (TERM_ALL key formatında, örn '17.04.2026')
    if term:
        last_date_raw = max(term.keys(), key=parse_sheet_date)
        last_date_display = parse_sheet_date(last_date_raw).strftime('%d.%m.%Y')
    else:
        last_date_raw = ''
        last_date_display = ''

    print(f'HTML güncelleniyor: {html_path}')
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()

    data_obj = {'margin': margin, 'loan': loan, 'btc': btc, 'earn': earn}
    html = update_html(html, data_obj, tenor, term, last_date_raw, last_date_display)

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print('─' * 50)
    print('✓ Dashboard başarıyla güncellendi')
    print(f'  Margin: {len(margin)} gün')
    print(f'  Loan:   {len(loan)} gün')
    print(f'  Earn:   {len(earn)} gün')
    print(f'  BTC:    {len(btc)} gün')
    print(f'  Term:   {len(term)} gün')
    print(f'  Son tarih: {last_date_display}')


if __name__ == '__main__':
    main()
