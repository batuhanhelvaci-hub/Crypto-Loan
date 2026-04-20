"""
Crypto Loan Dashboard - Veri Güncelleme Script'i
=================================================
Bu script Excel dosyasını okur ve HTML dashboard'u otomatik günceller.

Akış:
  1. _template.html  → şifresiz template (veriler bu dosyada güncellenir)
  2. Script template'i günceller
  3. Güncellenmiş template otomatik olarak ŞİFRELENİR
  4. crypto_loan_analysis (2).html → şifreli sürüm (yayınlanan dosya)

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
import os
import base64
import hashlib
from datetime import datetime
from functools import reduce


# ==================== ŞİFRELEME AYARLARI ====================

ENCRYPTION_PASSWORD = "Btcturkglobal2026?"
PBKDF2_ITERATIONS = 600000


# ==================== VERİ OKUMA FONKSİYONLARI ====================

def read_margin_and_btc(xl):
    """
    'Margin Borrow Rates' sheet'inden margin rate + BTC price verisini çeker.
    Döndürdüğü: (margin_listesi, btc_listesi)
    """
    df = xl.parse('Margin Borrow Rates')
    has_btc = df.shape[1] >= 7
    df.columns = ['Date','Binance','OKX','Bybit','KuCoin','Gate_IO','BTC'] if has_btc \
                 else ['Date','Binance','OKX','Bybit','KuCoin','Gate_IO']

    margin_result = []
    btc_result = []
    for _, row in df.iterrows():
        try:
            d = pd.to_datetime(row['Date']).strftime('%Y-%m-%d')
            if pd.notna(row['Binance']):
                margin_result.append({
                    'date': d,
                    'Binance': round(float(row['Binance'])*100, 4),
                    'OKX': round(float(row['OKX'])*100, 4) if pd.notna(row['OKX']) else None,
                    'Bybit': round(float(row['Bybit'])*100, 4) if pd.notna(row['Bybit']) else None,
                    'KuCoin': round(float(row['KuCoin'])*100, 4) if pd.notna(row['KuCoin']) else None,
                    'Gate_IO': round(float(row['Gate_IO'])*100, 4) if pd.notna(row['Gate_IO']) else None,
                })
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
    """
    df = xl.parse('Daily Data', header=None)
    data_rows = df.iloc[3:].reset_index(drop=True)

    def extract(col_date, col_flex, label):
        dates = pd.to_datetime(data_rows.iloc[:, col_date], errors='coerce')
        vals = pd.to_numeric(data_rows.iloc[:, col_flex], errors='coerce')
        d = pd.DataFrame({'date': dates, label: vals}).dropna()
        d['date'] = d['date'].dt.normalize()
        return d

    loan_dfs = [
        extract(0, 15, 'Binance'),
        extract(8, 33, 'OKX'),
        extract(44, 51, 'ByBit'),
        extract(62, 69, 'Gate_IO'),
    ]
    loan_merged = reduce(lambda a,b: pd.merge(a,b,on='date',how='outer'), loan_dfs).sort_values('date')

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
    """Günlük tarihli sheet'lerden vadeli (term) verileri çeker."""
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
    """TERM_ALL yapısından TENOR_DATA yapısına dönüştürür."""
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

    html = re.sub(
        r'const DATA = \{.*?\};',
        lambda m: f'const DATA = {json.dumps(data_obj)};',
        html, flags=re.DOTALL
    )
    html = re.sub(
        r'const TENOR_DATA = \{.*?\};',
        lambda m: f'const TENOR_DATA = {json.dumps(tenor_data)};',
        html, flags=re.DOTALL
    )
    html = re.sub(
        r'const TERM_ALL = \{.*?\};',
        lambda m: f'const TERM_ALL = {json.dumps(term_data)};',
        html, flags=re.DOTALL
    )
    html = re.sub(
        r'const LAST_DATE = "[^"]*";',
        f'const LAST_DATE = "{last_date_raw}";',
        html
    )
    html = re.sub(
        r'Son güncelleme: \d{1,2}\.\d{1,2}\.\d{4}',
        f'Son güncelleme: {last_date_display}',
        html
    )
    html = re.sub(
        r'Tümü \(01 Oca – \d{1,2}\.\d{1,2}\.\d{4}\)',
        f'Tümü (01 Oca – {last_date_display})',
        html
    )
    return html


# ==================== ŞİFRELEME ====================

def encrypt_html(plain_html, password):
    """
    Şifresiz HTML'i AES-256-GCM ile şifreler ve parola soran tek dosyalık
    bir HTML wrapper döndürür.
    """
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    except ImportError:
        raise ImportError(
            "cryptography kütüphanesi gerekli. Workflow dosyasına ekleyin:\n"
            "  pip install cryptography"
        )

    salt = os.urandom(16)
    iv = os.urandom(12)
    key = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS, dklen=32)
    ciphertext = AESGCM(key).encrypt(iv, plain_html.encode("utf-8"), None)

    salt_b64 = base64.b64encode(salt).decode("ascii")
    iv_b64 = base64.b64encode(iv).decode("ascii")
    ct_b64 = base64.b64encode(ciphertext).decode("ascii")

    template = r"""<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Crypto Loan - Protected</title>
<meta name="robots" content="noindex, nofollow">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  html, body { height: 100%; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: linear-gradient(135deg, #1e2130 0%, #2d3450 100%);
    color: #e6edf3;
    display: flex; align-items: center; justify-content: center;
    min-height: 100vh; padding: 20px;
  }
  .lock-box {
    background: #262c3d; border: 1px solid rgba(255,255,255,0.08);
    border-radius: 12px; padding: 40px 36px; width: 100%; max-width: 420px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.4);
  }
  .lock-icon { font-size: 48px; text-align: center; margin-bottom: 16px; }
  h1 { text-align: center; font-size: 20px; font-weight: 600; margin-bottom: 8px; }
  .subtitle { text-align: center; font-size: 13px; color: #8b949e; margin-bottom: 28px; }
  .field { margin-bottom: 16px; }
  label { display: block; font-size: 12px; color: #8b949e; margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.5px; }
  input[type="password"] {
    width: 100%; padding: 12px 14px; background: #1e2130;
    border: 1px solid rgba(255,255,255,0.12); border-radius: 8px;
    color: #e6edf3; font-size: 15px; font-family: inherit; transition: border-color 0.15s;
  }
  input[type="password"]:focus { outline: none; border-color: #58a6ff; }
  .remember {
    display: flex; align-items: center; gap: 8px; font-size: 13px;
    color: #8b949e; margin-bottom: 20px; user-select: none; cursor: pointer;
  }
  .remember input { cursor: pointer; }
  button {
    width: 100%; padding: 12px; background: #58a6ff; color: #0d1117;
    border: none; border-radius: 8px; font-size: 15px; font-weight: 600;
    cursor: pointer; transition: background 0.15s, transform 0.05s; font-family: inherit;
  }
  button:hover:not(:disabled) { background: #79c0ff; }
  button:active:not(:disabled) { transform: scale(0.98); }
  button:disabled { opacity: 0.6; cursor: wait; }
  .error {
    margin-top: 14px; padding: 10px 12px; background: rgba(248,81,73,0.1);
    border: 1px solid rgba(248,81,73,0.3); border-radius: 6px;
    color: #ff7b72; font-size: 13px; display: none;
  }
  .error.show { display: block; }
  .loader {
    display: inline-block; width: 14px; height: 14px;
    border: 2px solid rgba(13,17,23,0.3); border-top-color: #0d1117;
    border-radius: 50%; animation: spin 0.7s linear infinite;
    vertical-align: middle; margin-right: 8px;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>
<div class="lock-box">
  <div class="lock-icon">&#128274;</div>
  <h1>Crypto Loan Analysis</h1>
  <p class="subtitle">Bu icerik sifrelidir. Erismek icin parola girin.</p>
  <div class="field">
    <label for="pw">Parola</label>
    <input type="password" id="pw" autofocus autocomplete="current-password" />
  </div>
  <label class="remember">
    <input type="checkbox" id="remember" checked />
    Bu cihazda beni 30 gun hatirla
  </label>
  <button id="unlockBtn">Kilidi Ac</button>
  <div class="error" id="err"></div>
</div>
<script>
(function() {
  const PAYLOAD = { salt: "__SALT__", iv: "__IV__", ct: "__CT__", iter: __ITER__ };
  const STORAGE_KEY = "cla_unlock_v1";
  const REMEMBER_DAYS = 30;
  const pwInput = document.getElementById("pw");
  const btn = document.getElementById("unlockBtn");
  const errBox = document.getElementById("err");
  const remember = document.getElementById("remember");

  function b64ToBytes(b64) {
    const bin = atob(b64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return arr;
  }
  async function deriveKey(password, saltBytes) {
    const enc = new TextEncoder();
    const baseKey = await crypto.subtle.importKey("raw", enc.encode(password), { name: "PBKDF2" }, false, ["deriveKey"]);
    return crypto.subtle.deriveKey(
      { name: "PBKDF2", salt: saltBytes, iterations: PAYLOAD.iter, hash: "SHA-256" },
      baseKey, { name: "AES-GCM", length: 256 }, false, ["decrypt"]
    );
  }
  async function decrypt(password) {
    const key = await deriveKey(password, b64ToBytes(PAYLOAD.salt));
    const plainBuf = await crypto.subtle.decrypt({ name: "AES-GCM", iv: b64ToBytes(PAYLOAD.iv) }, key, b64ToBytes(PAYLOAD.ct));
    return new TextDecoder().decode(plainBuf);
  }
  function renderHtml(html) { document.open(); document.write(html); document.close(); }
  function showError(msg) { errBox.textContent = msg; errBox.classList.add("show"); }
  function clearError() { errBox.classList.remove("show"); errBox.textContent = ""; }

  async function tryUnlock(password, fromRemembered) {
    clearError();
    btn.disabled = true;
    btn.innerHTML = '<span class="loader"></span>Kilit aciliyor...';
    try {
      const html = await decrypt(password);
      if (remember && remember.checked && !fromRemembered) {
        const expires = Date.now() + REMEMBER_DAYS * 24 * 60 * 60 * 1000;
        try { localStorage.setItem(STORAGE_KEY, JSON.stringify({ p: password, e: expires })); } catch(e) {}
      }
      renderHtml(html);
    } catch (e) {
      btn.disabled = false;
      btn.textContent = "Kilidi Ac";
      if (fromRemembered) {
        try { localStorage.removeItem(STORAGE_KEY); } catch(e) {}
      } else {
        showError("Parola yanlis. Lutfen tekrar deneyin.");
        pwInput.select();
      }
    }
  }
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const obj = JSON.parse(raw);
      if (obj && obj.e && obj.e > Date.now() && obj.p) { tryUnlock(obj.p, true); return; }
      else { localStorage.removeItem(STORAGE_KEY); }
    }
  } catch(e) {}
  btn.addEventListener("click", () => {
    const pw = pwInput.value;
    if (!pw) { showError("Parola bos olamaz."); return; }
    tryUnlock(pw, false);
  });
  pwInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") btn.click();
    else clearError();
  });
})();
</script>
</body>
</html>
"""

    return (template
            .replace("__SALT__", salt_b64)
            .replace("__IV__", iv_b64)
            .replace("__CT__", ct_b64)
            .replace("__ITER__", str(PBKDF2_ITERATIONS)))


# ==================== ANA AKIŞ ====================

def main():
    excel_path = 'Daily Earn-Borrow Rates.xlsx'
    template_path = '_template.html'          # şifresiz template (script bunu günceller)
    output_path = 'crypto_loan_analysis (2).html'  # şifreli yayın dosyası

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

    if term:
        last_date_raw = max(term.keys(), key=parse_sheet_date)
        last_date_display = parse_sheet_date(last_date_raw).strftime('%d.%m.%Y')
    else:
        last_date_raw = ''
        last_date_display = ''

    print(f'Template okunuyor: {template_path}')
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    data_obj = {'margin': margin, 'loan': loan, 'btc': btc, 'earn': earn}
    html = update_html(html, data_obj, tenor, term, last_date_raw, last_date_display)

    # Template'i de güncelle (bir sonraki çalıştırmada güncel veriden başlasın diye)
    with open(template_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'✓ Template güncellendi: {template_path}')

    # ŞİFRELE ve yayın dosyasını yaz
    print('Şifreleme yapılıyor (AES-256-GCM)...')
    encrypted_html = encrypt_html(html, ENCRYPTION_PASSWORD)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(encrypted_html)
    print(f'✓ Şifreli sürüm yazıldı: {output_path}')

    print('─' * 50)
    print('✓ Dashboard başarıyla güncellendi ve şifrelendi')
    print(f'  Margin: {len(margin)} gün')
    print(f'  Loan:   {len(loan)} gün')
    print(f'  Earn:   {len(earn)} gün')
    print(f'  BTC:    {len(btc)} gün')
    print(f'  Term:   {len(term)} gün')
    print(f'  Son tarih: {last_date_display}')
    print(f'  Şifreli dosya boyutu: {os.path.getsize(output_path):,} bytes')


if __name__ == '__main__':
    main()
