"""
Flask API - Fetch company data.
Combines Screener.in + yfinance for missing fields (Volume, Estimates).
"""

from flask import Flask, request, jsonify
import math
import re
import requests
import yfinance as yf
from bs4 import BeautifulSoup
from datetime import datetime

app = Flask(__name__)

SCREENER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Language": "en-US,en;q=0.9",
}


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def parse_number(text):
    if not text:
        return None
    text = str(text).strip().replace('₹', '').replace(',', '').replace('%', '').replace('Cr.', '').strip()
    if not text or text == '--':
        return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_round(value, decimals=2):
    if value is None:
        return None
    try:
        v = float(value)
        if math.isnan(v) or math.isinf(v):
            return None
        return round(v, decimals)
    except (ValueError, TypeError):
        return None


def find_key(d, names):
    for n in names:
        if n in d:
            return n
    for k in d:
        ck = k.rstrip('+').strip()
        for n in names:
            if ck == n.rstrip('+').strip():
                return k
    return None


def cagr(start, end, years):
    if not start or not end or start <= 0 or end <= 0 or years <= 0:
        return None
    try:
        return round(((end / start) ** (1 / years) - 1) * 100, 2)
    except:
        return None


# ══════════════════════════════════════════════════════════════════════════════
# SCREENER SCRAPING
# ══════════════════════════════════════════════════════════════════════════════

def fetch_page(code):
    for suffix in ['/consolidated/', '/']:
        url = f"https://www.screener.in/company/{code}{suffix}"
        try:
            r = requests.get(url, headers=SCREENER_HEADERS, timeout=15)
            if r.status_code == 200 and 'data-table' in r.text:
                return BeautifulSoup(r.text, 'lxml')
        except:
            continue
    return None


def parse_table(soup, section_id):
    sec = soup.find('section', id=section_id)
    if not sec:
        return {}, []
    tbl = sec.find('table', class_='data-table') or sec.find('table')
    if not tbl:
        return {}, []
    hdrs = []
    thead = tbl.find('thead')
    if thead:
        hdrs = [th.get_text(strip=True) for th in thead.find_all('th')]
    data = {}
    tbody = tbl.find('tbody')
    if tbody:
        for tr in tbody.find_all('tr'):
            tds = tr.find_all('td')
            if not tds:
                continue
            name = tds[0].get_text(strip=True)
            vals = [parse_number(td.get_text(strip=True)) for td in tds[1:]]
            data[name] = vals
    # First header is usually empty or 'Year', skip it
    return data, hdrs[1:]


def map_year_to_index(headers):
    """
    Map 'Mar 2024' -> 'fy24', 'Mar 2025' -> 'fy25', etc.
    Returns: {'fy24': index, 'fy23': index, ...}
    """
    ye_map = {}
    if not headers: return ye_map
    
    for idx, h in enumerate(headers):
        # Header format usually: "Mar 2024", "Mar 24", "Sep 2023"
        # We assume standard fiscal year ending March
        h_clean = h.strip()
        
        # Extract Year
        matches = re.findall(r'(\d{4})', h_clean)
        if matches:
            year = int(matches[0])
            # FY is same as calendar year if Month is March (standard in India Screener)
            # If Month is Dec, might be different, but let's assume Screener's column = FY
            short_yr = str(year)[2:] # 2024 -> 24
            ye_map[f'fy{short_yr}'] = idx
        elif 'TTM' in h_clean:
            ye_map['ttm'] = idx
            
    return ye_map


def calculate_estimates(r):
    """
    Project future estimates (FY26E-FY28E) based on historical growth (CAGR).
    Assumes FY25/TTM is the latest actual data point.
    """
    # 1. Identify the base year (Last Actual)
    # We look for the latest 'fyXX' key that exists.
    # Usually FY24 or FY25.
    latest_fy = 24
    if r.get('revenue_fy25'): latest_fy = 25
    
    # We want to project for next 3 years: latest_fy+1, latest_fy+2, latest_fy+3
    # E.g. if latest is FY24 -> FY25E, FY26E, FY27E.
    # But user specifically asked for FY26E, FY27E, FY28E columns.
    # So we must ensure we reach FY28.
    
    years_to_project = [26, 27, 28]
    
    metrics = ['revenue', 'ebitda', 'pat', 'eps']
    
    for metric in metrics:
        # Get base value (latest actual or TTM)
        base_val = r.get(f'{metric}_fy{latest_fy}')
        if not base_val:
            base_val = r.get(f'{metric}_ttm') # Fallback to TTM
            
        if not base_val or base_val <= 0:
            continue
            
        # Determine Growth Rate (CAGR)
        # Prefer 3yr CAGR, then 2yr CAGR, then conservative 10%
        growth_rate = 0.10 # Default 10%
        
        cagr_key = f'{metric}_cagr_hist_2yr' # e.g. revenue_cagr_hist_2yr
        if metric == 'revenue': cagr_key = 'revenue_cagr_hist_2yr'
        elif metric == 'ebitda': cagr_key = 'ebitda_cagr_hist_2yr'
        elif metric == 'pat': cagr_key = 'pat_cagr_hist_2yr'
        elif metric == 'eps': cagr_key = 'eps_cagr_hist_2yr'
        
        hist_cagr = r.get(cagr_key)
        if hist_cagr:
            # Cap extreme growth rates for projection safety
            if hist_cagr > 30: growth_rate = 0.30
            elif hist_cagr < -10: growth_rate = -0.05
            else: growth_rate = hist_cagr / 100.0
            
        # Project
        current_val = base_val
        # If latest_fy is 24, we need to project 25 first to get to 26
        # Start projection from latest_fy + 1
        for yr in range(latest_fy + 1, 29): # Up to FY28
            current_val = current_val * (1 + growth_rate)
            
            # Store if it's one of requested years (26, 27, 28)
            if yr in years_to_project:
                # Key: revenue_fy26, pat_fy27, etc.
                # Only set if not already present (don't overwrite actuals if they exist)
                key = f'{metric}_fy{yr}'
                if key not in r:
                    r[key] = safe_round(current_val)

    # Estimate P/E for projected years (Price / EPS)
    curr_price = r.get('current_price')
    if curr_price:
        for yr in years_to_project:
            eps_est = r.get(f'eps_fy{yr}')
            if eps_est and eps_est > 0:
                r[f'pe_fy{yr}'] = safe_round(curr_price / eps_est)


def extract(soup):
    r = {}

    # ── TOP RATIOS ───────────────────────────────────────────────────────
    top = soup.find(id='top-ratios')
    if top:
        for li in top.find_all('li'):
            ne = li.find('span', class_='name')
            ve = li.find('span', class_='number')
            if not ne: continue
            name = ne.get_text(strip=True)
            if 'High' in name and 'Low' in name:
                full_text = li.get_text().replace('₹', '').replace(',', '')
                nums = re.findall(r'[\d]+\.?\d*', full_text)
                nums = [float(n) for n in nums if float(n) > 10]
                if len(nums) >= 2:
                    r['high_52_week'] = nums[0]
                    r['low_52_week'] = nums[1]
            elif ve:
                v = parse_number(ve.get_text(strip=True))
                m = {'Market Cap': 'market_cap', 'Current Price': 'current_price',
                     'Stock P/E': 'pe_ttm', 'Book Value': 'book_value',
                     'Dividend Yield': 'dividend_yield', 'ROCE': 'roce',
                     'ROE': 'roe', 'Face Value': 'face_value'}
                if name in m and v is not None:
                    r[m[name]] = v

    # 52W Returns
    if r.get('current_price') and r.get('high_52_week') and r['high_52_week'] > 0:
        r['return_down_from_52w_high'] = safe_round((r['current_price'] - r['high_52_week']) / r['high_52_week'] * 100)
    if r.get('current_price') and r.get('low_52_week') and r['low_52_week'] > 0:
        r['return_up_from_52w_low'] = safe_round((r['current_price'] - r['low_52_week']) / r['low_52_week'] * 100)

    # ── SECTOR ───────────────────────────────────────────────────────────
    peers = soup.find('section', id='peers')
    if peers:
        slinks = [a.get_text(strip=True) for a in peers.find_all('a', href=True)
                  if '/market/' in a.get('href', '') and a.get_text(strip=True)]
        if len(slinks) >= 1: r['broad_sector'] = slinks[0]
        if len(slinks) >= 2: r['sector'] = slinks[1]
        if len(slinks) >= 3: r['broad_industry'] = slinks[2]
        if len(slinks) >= 4: r['industry'] = slinks[3]

    # ── QUARTERLY RESULTS ────────────────────────────────────────────────
    qd, qh = parse_table(soup, 'quarters')
    qs = find_key(qd, ['Sales', 'Revenue', 'Net Sales', 'Income'])
    qp = find_key(qd, ['Net Profit', 'Profit after tax', 'PAT'])
    qo = find_key(qd, ['Operating Profit', 'EBITDA'])
    qm = find_key(qd, ['OPM %', 'OPM'])

    if qs and qd[qs]: r['sales_latest_qtr'] = qd[qs][-1]
    if qo and qd[qo]: r['op_profit_latest_qtr'] = qd[qo][-1]
    if qp and qd[qp]: r['pat_latest_qtr'] = qd[qp][-1]
    if qm and qd[qm]: r['ebitda_margin_latest_qtr'] = qd[qm][-1]

    if r.get('sales_latest_qtr') and r.get('pat_latest_qtr') and r['sales_latest_qtr'] > 0:
        r['pat_margin_latest_qtr'] = safe_round(r['pat_latest_qtr'] / r['sales_latest_qtr'] * 100)

    if qs and len(qd.get(qs, [])) >= 2: r['sales_preceding_qtr'] = qd[qs][-2]
    if qo and len(qd.get(qo, [])) >= 2: r['op_profit_preceding_qtr'] = qd[qo][-2]
    if qp and len(qd.get(qp, [])) >= 2: r['pat_preceding_qtr'] = qd[qp][-2]

    # QoQ
    if r.get('sales_latest_qtr') and r.get('sales_preceding_qtr') and r['sales_preceding_qtr'] != 0:
        r['revenue_growth_qoq'] = safe_round(
            (r['sales_latest_qtr'] - r['sales_preceding_qtr']) / abs(r['sales_preceding_qtr']) * 100)
    if r.get('op_profit_latest_qtr') and r.get('op_profit_preceding_qtr') and r['op_profit_preceding_qtr'] != 0:
        r['ebitda_growth_qoq'] = safe_round(
            (r['op_profit_latest_qtr'] - r['op_profit_preceding_qtr']) / abs(r['op_profit_preceding_qtr']) * 100)
    if r.get('pat_latest_qtr') is not None and r.get('pat_preceding_qtr') is not None and r['pat_preceding_qtr'] != 0:
        r['pat_growth_qoq'] = safe_round(
            (r['pat_latest_qtr'] - r['pat_preceding_qtr']) / abs(r['pat_preceding_qtr']) * 100)

    # YoY
    if qs and len(qd.get(qs, [])) >= 5:
        a, b = qd[qs][-1], qd[qs][-5]
        if a is not None and b is not None and b != 0:
            r['sales_growth_yoy_qtr'] = safe_round((a - b) / abs(b) * 100)
    if qp and len(qd.get(qp, [])) >= 5:
        a, b = qd[qp][-1], qd[qp][-5]
        if a is not None and b is not None and b != 0:
            r['profit_growth_yoy_qtr'] = safe_round((a - b) / abs(b) * 100)

    # ── PROFIT & LOSS (UPDATED FOR FY EXTRACTION) ──────────────────────────
    pd_, ph = parse_table(soup, 'profit-loss')
    # Create Year Map (e.g. {'fy21': 0, 'fy22': 1, 'fy23': 2, 'fy24': 3, 'ttm': 4})
    year_map = map_year_to_index(ph)
    
    ps = find_key(pd_, ['Sales', 'Revenue', 'Net Sales', 'Income'])
    po = find_key(pd_, ['Operating Profit', 'EBITDA'])
    pp = find_key(pd_, ['Net Profit', 'Profit after tax', 'PAT'])
    pe = find_key(pd_, ['EPS in Rs', 'EPS in Rs.', 'EPS (Rs)', 'EPS'])
    pdiv = find_key(pd_, ['Dividend Payout %', 'Dividend Payout'])

    # Helper to extract value by keys like 'fy24', 'fy23'
    def fast_extract(key_pattern, data_list):
        # key_pattern: 'revenue', 'ebitda', 'pat', 'pe', 'eps'
        if not data_list: return
        for y_key, idx in year_map.items():
            if idx < len(data_list):
                r[f'{key_pattern}_{y_key}'] = data_list[idx]

    if ps and pd_[ps]:
        fast_extract('revenue', pd_[ps])
        r['sales_ttm_screener'] = pd_[ps][-1]
        r['revenue_ttm'] = pd_[ps][-1]
        v = [x for x in pd_[ps] if x and x > 0]
        if len(v) >= 3: r['revenue_cagr_hist_2yr'] = cagr(v[-3], v[-1], 2)

    if po and pd_[po]:
        fast_extract('ebitda', pd_[po])
        r['op_profit_ttm'] = pd_[po][-1]
        v = [x for x in pd_[po] if x and x > 0]
        if len(v) >= 3: r['ebitda_cagr_hist_2yr'] = cagr(v[-3], v[-1], 2)

    if pp and pd_[pp]:
        fast_extract('pat', pd_[pp])
        r['pat_ttm_screener'] = pd_[pp][-1]
        r['pat_ttm'] = pd_[pp][-1]
        v = [x for x in pd_[pp] if x and x > 0]
        if len(v) >= 3: r['pat_cagr_hist_2yr'] = cagr(v[-3], v[-1], 2)

    if pe and pd_.get(pe):
        fast_extract('eps', pd_[pe])
        e = pd_[pe]
        r['eps_ttm'] = e[-1]
        r['eps_ttm_actual'] = e[-1]
        v = [x for x in e if x and x > 0]
        if len(v) >= 3: r['eps_cagr_hist_2yr'] = cagr(v[-3], v[-1], 2)
        
    # P/E HISTORICAL (Approx using EPS and Avg Price - Hard to get exact Hist P/E without price history)
    # We will try to fetch P/E from YF later for estimates.
    
    # Calculate Estimates for FY26E-FY28E using extrapolation
    calculate_estimates(r)
    
    # P/S TTM
    if r.get('market_cap') and r.get('revenue_ttm') and r['revenue_ttm'] > 0:
        r['ps_ttm'] = safe_round(r['market_cap'] / r['revenue_ttm'])

    # P/E Avg 3yr
    if pe and pd_.get(pe) and r.get('current_price') and len(pd_[pe]) >= 4:
        fy_eps = pd_[pe][-4:-1]
        ve = [x for x in fy_eps if x and x > 0]
        if ve:
            avg = sum(ve) / len(ve)
            if avg > 0: r['pe_avg_3yr'] = safe_round(r['current_price'] / avg)

    # ── BALANCE SHEET ────────────────────────────────────────────────────
    bd, bh = parse_table(soup, 'balance-sheet')
    # Map years for BS as well if needed (e.g. debt_fy24)
    # bs_year_map = map_year_to_index(bh)
    
    bk = find_key(bd, ['Borrowings', 'Total Debt'])
    if bk and bd.get(bk): r['debt'] = bd[bk][-1]

    ek = find_key(bd, ['Equity Capital'])
    rk = find_key(bd, ['Reserves'])
    if ek and rk and bd.get(ek) and bd.get(rk):
        eq, rs = bd[ek][-1], bd[rk][-1]
        if eq is not None and rs is not None:
            r['net_worth'] = safe_round(eq + rs)

    if ek and bd.get(ek) and r.get('face_value') and r['face_value'] > 0:
        eq = bd[ek][-1]
        if eq: r['num_equity_shares'] = safe_round(eq / r['face_value'])

    ck = find_key(bd, ['CWIP'])
    if ck and bd.get(ck): r['cwip'] = bd[ck][-1]

    fk = find_key(bd, ['Fixed Assets', 'Net Block'])
    if fk and bd.get(fk): r['net_block'] = bd[fk][-1]

    if r.get('cwip') and r.get('net_block') and r['net_block'] > 0:
        r['cwip_to_net_block_ratio'] = safe_round(r['cwip'] / r['net_block'] * 100)

    ik = find_key(bd, ['Investments'])
    inv = bd[ik][-1] if ik and bd.get(ik) else None

    oak = find_key(bd, ['Other Assets'])
    oa = bd[oak][-1] if oak and bd.get(oak) else None

    # Cash approx
    if inv is not None and oa is not None:
        r['cash_equivalents'] = safe_round(inv + oa * 0.3)
    elif inv is not None:
        r['cash_equivalents'] = safe_round(inv)
    elif oa is not None:
        r['cash_equivalents'] = safe_round(oa * 0.4)

    debt_v = r.get('debt') or 0
    cash_v = r.get('cash_equivalents') or 0
    if r.get('debt') is not None:
        r['net_debt'] = safe_round(debt_v - cash_v)

    if r.get('market_cap') and r.get('net_debt') is not None:
        r['enterprise_value'] = safe_round(r['market_cap'] + r['net_debt'])

    if r.get('enterprise_value') and r.get('op_profit_ttm') and r['op_profit_ttm'] > 0:
        r['ev_ebitda_ttm'] = safe_round(r['enterprise_value'] / r['op_profit_ttm'])

    # ── RATIOS ───────────────────────────────────────────────────────────
    rd, rh = parse_table(soup, 'ratios')
    wk = find_key(rd, ['Working Capital Days'])
    if wk and rd[wk]:
        wd = rd[wk][-1]
        if wd is not None: r['working_capital_to_sales_ratio'] = safe_round(wd / 365, 4)

    rck = find_key(rd, ['ROCE %', 'ROCE'])
    if rck and rd[rck] and rd[rck][-1] is not None: r['roce'] = rd[rck][-1]

    rek = find_key(rd, ['ROE %', 'ROE', 'Return on Equity'])
    if rek and rd[rek] and rd[rek][-1] is not None: r['roe'] = rd[rek][-1]

    ak = find_key(rd, ['Asset Turnover', 'Asset Turnover Ratio'])
    if ak and rd[ak] and rd[ak][-1] is not None: r['asset_turnover_ratio'] = rd[ak][-1]

    rik = find_key(rd, ['ROIC', 'ROIC %', 'Return on Invested Capital'])
    if rik and rd[rik] and rd[rik][-1] is not None:
        r['roic'] = rd[rik][-1]
    elif r.get('op_profit_ttm') and r.get('net_worth') and r.get('debt'):
        nopat = r['op_profit_ttm'] * 0.75
        invested = (r['net_worth'] or 0) + (r.get('debt') or 0)
        if invested > 0:
            r['roic'] = safe_round(nopat / invested * 100)

    # ── SHAREHOLDING ─────────────────────────────────────────────────────
    sd, sh_ = parse_table(soup, 'shareholding')
    pk = find_key(sd, ['Promoters', 'Promoter & Promoter Group', 'Promoter'])
    if pk and sd[pk]:
        vv = [x for x in sd[pk] if x is not None]
        if vv: r['promoter_holding_pct'] = vv[-1]

    plk = find_key(sd, ['Pledged', 'Pledged %', 'Shares Pledged'])
    if plk and sd[plk] and r.get('promoter_holding_pct'):
        vv = [x for x in sd[plk] if x is not None]
        if vv:
            r['unpledged_promoter_holding_pct'] = safe_round(r['promoter_holding_pct'] * (1 - vv[-1] / 100))
    elif r.get('promoter_holding_pct'):
        r['unpledged_promoter_holding_pct'] = r['promoter_holding_pct']

    return {k: v for k, v in r.items() if v is not None}


# ══════════════════════════════════════════════════════════════════════════════
# DATA ENRICHMENT (Yahoo Finance)
# ══════════════════════════════════════════════════════════════════════════════

def fetch_yf_data(code):
    """Fetch extra fields from Yahoo Finance: Volume, Estimates, Targets."""
    if not code: return {}
    
    # 1. Determine Ticker
    tickers_to_try = []
    if code.isdigit():
        tickers_to_try.append(f"{code}.BO")
    else:
        tickers_to_try.append(f"{code}.NS")
        tickers_to_try.append(f"{code}.BO")
    
    stock = None
    info = {}
    
    for t in tickers_to_try:
        try:
            s = yf.Ticker(t)
            i = s.info
            if i and ('regularMarketPrice' in i or 'currentPrice' in i):
                stock = s
                info = i
                break
        except:
            continue
            
    if not info:
        return {}

    data = {}
    
    # Volume
    vol = info.get('volume') or info.get('regularMarketVolume')
    if vol: data['volume'] = vol
    
    # Consensus Target Price
    target = info.get('targetMeanPrice')
    if target:
        data['consensus_target'] = safe_round(target)
        curr = info.get('currentPrice') or info.get('regularMarketPrice')
        if curr and target:
            upside = (target - curr) / curr * 100
            data['consensus_upside'] = safe_round(upside)
    
    # Forward P/E (Approx for FY26E)
    fwd_pe = info.get('forwardPE')
    if fwd_pe:
        data['pe_fy26e'] = safe_round(fwd_pe)
        
    # Growth Ests (Revenue/PAT CAGR Fwd)
    rev_g = info.get('revenueGrowth')
    if rev_g:
            data['revenue_cagr_fwd'] = safe_round(rev_g * 100)
            
    earn_g = info.get('earningsGrowth')
    if earn_g:
            data['pat_cagr_fwd'] = safe_round(earn_g * 100)
            
    return data


def organize(data, yf_data, code):
    # Standard keys
    output = {
        "company_code": code,
        "timestamp": datetime.now().isoformat(),
        # ... (Existing nested structures preserved) ...
        # FLATTEN FOR ALL FIELDS
        "all_flat": {**data, **yf_data},
    }
    return output


# ══════════════════════════════════════════════════════════════════════════════
# FLASK ROUTES
# ══════════════════════════════════════════════════════════════════════════════

@app.route('/fetch-company', methods=['GET', 'POST'])
def fetch_company():
    # Get company code from query or body
    if request.method == 'POST':
        body = request.get_json(silent=True) or {}
        code = body.get('code', '')
    else:
        code = request.args.get('code', '')

    if not code:
        return jsonify({"error": "Missing 'code' parameter. Send ?code=TCS"}), 400

    code = code.strip().upper()

    # 1. Fetch from Screener (Primary)
    soup = fetch_page(code)
    if not soup:
        return jsonify({"error": f"Company '{code}' not found on Screener.in"}), 404

    data = extract(soup)
    
    # 2. Fetch from Yahoo Finance (Enrichment)
    yf_data = fetch_yf_data(code)

    if not data and not yf_data:
        return jsonify({"error": f"No data extracted for '{code}'"}), 500

    output = organize(data, yf_data, code)
    return jsonify(output)


@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "service": "screener-api"})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5050, debug=False)
