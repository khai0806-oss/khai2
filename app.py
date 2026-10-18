"""
HOSE QUANTITATIVE PORTFOLIO STRATEGY - STREAMLIT WEB APP
=========================================================
Kiểm định tính hiệu quả của chiến lược định lượng cổ phiếu HOSE:
- Multi-factor Technical Analysis (TA Score 0-10)
- Nested Walk-Forward Stock Selection (Top N)
- Dynamic Portfolio Optimization (EW / MPT-LW / MinVar-LW / HRP / Blend50)
- Market Regime & Risk Overlay (EMA200 + Slope Confirmation)
- Thí nghiệm chính sách tiền mặt (Cash Policy: Cash40 / Original / Full)
"""

import os
import re
import math
import io
import warnings
from pathlib import Path
from itertools import product
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

from scipy.optimize import minimize
from scipy.cluster.hierarchy import linkage
from scipy.spatial.distance import squareform
from sklearn.covariance import LedoitWolf

warnings.filterwarnings("ignore")

# ==========================================================
# 0. STREAMLIT PAGE CONFIGURATION & STYLING
# ==========================================================
st.set_page_config(
    page_title="HOSE Quant Strategy - Backtesting & Verification",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for polished financial dashboard
st.markdown("""
<style>
    /* Metric Card Styling */
    div[data-testid="stMetricValue"] {
        font-size: 1.7rem !important;
        font-weight: 700;
        color: #1E3A8A;
    }
    div[data-testid="stMetricDelta"] {
        font-size: 0.85rem !important;
    }
    .metric-card {
        background-color: #f8fafc;
        border-radius: 10px;
        padding: 15px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .section-header {
        font-size: 1.25rem;
        font-weight: 600;
        color: #0f172a;
        margin-top: 15px;
        margin-bottom: 10px;
        border-bottom: 2px solid #3b82f6;
        padding-bottom: 5px;
    }
    .badge-champion {
        background-color: #10b981;
        color: white;
        padding: 4px 10px;
        border-radius: 15px;
        font-weight: bold;
        font-size: 0.85rem;
        display: inline-block;
    }
    .badge-benchmark {
        background-color: #64748b;
        color: white;
        padding: 4px 10px;
        border-radius: 15px;
        font-weight: bold;
        font-size: 0.85rem;
        display: inline-block;
    }
    /* Tabs font size */
    button[data-baseweb="tab"] {
        font-size: 1rem !important;
        font-weight: 600 !important;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================================
# 1. CONSTANTS & DICTIONARIES
# ==========================================================
TRADING_DAYS = 252
BENCHMARK_ALIASES = {"VNINDEX", "VN-INDEX", "VNI", "^VNINDEX", "VNINDEXINDEX"}

FA_ALIASES = {
    "ROE": ["roe", "return_on_equity", "returnonequity"],
    "ROA": ["roa", "return_on_assets", "returnonassets"],
    "RevenueGrowth": ["revenue_growth", "sales_growth", "revenuegrowth", "tangtruongdoanhthu"],
    "EPSGrowth": ["eps_growth", "epsgrowth", "tangtruongeps"],
    "GrossMargin": ["gross_margin", "grossmargin", "bienloinhuangop"],
    "OperatingMargin": ["operating_margin", "operatingmargin"],
    "NetMargin": ["net_margin", "netmargin", "bienloinhuanrong"],
    "CurrentRatio": ["current_ratio", "currentratio"],
    "DebtToEquity": ["debt_to_equity", "de", "d_e", "debttoequity"],
    "PE": ["pe", "p_e", "price_to_earnings", "pricetoearnings"],
    "PB": ["pb", "p_b", "price_to_book", "pricetobook"],
    "EVEBITDA": ["ev_ebitda", "evebitda"],
}

VI_COLUMN_MAP = {
    "Mã": "Ticker",
    "Ngày": "Date",
    "Giá mở cửa": "Open",
    "Giá cao nhất": "High",
    "Giá thấp nhất": "Low",
    "Giá đóng cửa": "Close",
    "Giá điều chỉnh": "adj_close",
    "Khối lượng khớp lệnh": "Volume"
}


# ==========================================================
# 2. CORE HELPER FUNCTIONS & DATA STANDARDIZATION
# ==========================================================
def _canon(s):
    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

def find_col(df, candidates):
    lookup = {_canon(c): c for c in df.columns}
    for c in candidates:
        if _canon(c) in lookup:
            return lookup[_canon(c)]
    return None

def parse_dates_safely(s):
    s = pd.Series(s)
    x = pd.to_datetime(s, format="%m/%d/%Y", errors="coerce")
    if x.notna().mean() < 0.80:
        x = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if x.notna().mean() < 0.80:
        x = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return x

def find_default_data_file():
    """Tự động tìm kiếm file dữ liệu có sẵn trong thư mục làm việc."""
    candidates = [
        "HOSE - 01.01.2021 - 31.12.2025 (98 ma bao gồm VNINDEX).xlsx",
        "HOSE - 01.01.2021 - 31.12.2025 (98 ma bao gom VNINDEX).xlsx",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    for f in os.listdir("."):
        if f.endswith((".xlsx", ".xls")) and not f.startswith("~$"):
            return f
    for f in os.listdir("."):
        if f.endswith(".csv"):
            return f
    return None

@st.cache_data(show_spinner=False)
def load_and_standardize_data(file_source, use_adjusted=True):
    """Đọc dữ liệu từ uploaded file hoặc đường dẫn file cục bộ và chuẩn hóa OHLCV."""
    if isinstance(file_source, str):
        if not os.path.exists(file_source):
            raise FileNotFoundError(f"Không tìm thấy file: {file_source}")
        if file_source.endswith(('.xlsx', '.xls')):
            # Thử đọc sheet LichSuGia nếu có, nếu không đọc sheet đầu tiên
            try:
                raw = pd.read_excel(file_source, sheet_name="LichSuGia")
            except Exception:
                raw = pd.read_excel(file_source)
        else:
            raw = pd.read_csv(file_source)
    else:
        # Streamlit UploadedFile
        fname = file_source.name.lower()
        if fname.endswith(('.xlsx', '.xls')):
            try:
                raw = pd.read_excel(file_source, sheet_name="LichSuGia")
            except Exception:
                file_source.seek(0)
                raw = pd.read_excel(file_source)
        else:
            raw = pd.read_csv(file_source)

    # Ánh xạ cột tiếng Việt có dấu sang tên chuẩn
    raw = raw.rename(columns=VI_COLUMN_MAP)
    raw.columns = [str(c).strip() for c in raw.columns]

    ticker_col = find_col(raw, ["ticker", "symbol", "code", "stockcode", "mack", "ma"])
    date_col = find_col(raw, ["date", "datetime", "tradingdate", "time", "ngay"])
    volume_col = find_col(raw, ["volume", "vol", "tradingvolume", "khoiluong", "khoiluongkhoplenh"])

    raw_cols = {
        "Open": find_col(raw, ["open", "openprice", "giamocua"]),
        "High": find_col(raw, ["high", "highprice", "giacaonhat"]),
        "Low": find_col(raw, ["low", "lowprice", "giathapnhat"]),
        "Close": find_col(raw, ["close", "closeprice", "giadongcua"]),
    }
    adj_cols = {
        "Open": find_col(raw, ["adjopen", "adjustedopen"]),
        "High": find_col(raw, ["adjhigh", "adjustedhigh"]),
        "Low": find_col(raw, ["adjlow", "adjustedlow"]),
        "Close": find_col(raw, ["adjclose", "adjustedclose", "giadieuchinh"]),
    }

    if ticker_col is None or date_col is None:
        raise ValueError(f"File thiếu cột Ticker hoặc Date. Các cột tìm thấy: {list(raw.columns)}")

    use_adj = bool(use_adjusted and adj_cols["Close"] is not None)
    pcols = {k: adj_cols[k] if (use_adj and adj_cols[k] is not None) else raw_cols[k] for k in ["Open", "High", "Low", "Close"]}

    if not all(v is not None for v in pcols.values()):
        # Fallback nếu thiếu cột Open/High/Low nhưng có Close
        if pcols["Close"] is not None:
            for k in ["Open", "High", "Low"]:
                if pcols[k] is None:
                    pcols[k] = pcols["Close"]
        else:
            raise ValueError(f"Thiếu các cột giá OHLC cần thiết.")

    out = pd.DataFrame({
        "Ticker": raw[ticker_col].astype(str).str.strip().str.upper(),
        "Date": parse_dates_safely(raw[date_col]),
        "Open": pd.to_numeric(raw[pcols["Open"]], errors="coerce"),
        "High": pd.to_numeric(raw[pcols["High"]], errors="coerce"),
        "Low": pd.to_numeric(raw[pcols["Low"]], errors="coerce"),
        "Close": pd.to_numeric(raw[pcols["Close"]], errors="coerce"),
        "Volume": pd.to_numeric(raw[volume_col], errors="coerce") if volume_col else 0.0,
    })

    out = out.dropna(subset=["Ticker", "Date", "Open", "High", "Low", "Close"])
    out = out[(out[["Open", "High", "Low", "Close"]] > 0).all(axis=1)]
    out["Volume"] = out["Volume"].fillna(0).clip(lower=0)

    # Đảm bảo tính hợp lệ cơ bản của OHLC
    valid_ohlc = (
        (out["High"] >= out[["Open", "Close", "Low"]].max(axis=1) * 0.95) &
        (out["Low"] <= out[["Open", "Close", "High"]].min(axis=1) * 1.05)
    )
    out = out.loc[valid_ohlc]
    out = out.drop_duplicates(["Ticker", "Date"], keep="last")
    out = out.sort_values(["Ticker", "Date"]).reset_index(drop=True)
    return out, raw

def data_quality_table(data):
    rows = []
    for t, g in data.groupby("Ticker"):
        rows.append({
            "Ticker": t,
            "Start": g["Date"].min().strftime("%Y-%m-%d"),
            "End": g["Date"].max().strftime("%Y-%m-%d"),
            "Sessions": int(g["Date"].nunique()),
            "MissingVolumePct": float(100 * (g["Volume"] == 0).mean()),
            "MedianTradedValue": float((g["Close"] * g["Volume"]).median()),
        })
    return pd.DataFrame(rows).sort_values("Sessions", ascending=False).reset_index(drop=True)

def build_price_dict(data):
    return {
        t: g[["Date", "Open", "High", "Low", "Close", "Volume"]]
             .drop_duplicates("Date").sort_values("Date").set_index("Date")
        for t, g in data.groupby("Ticker")
    }

def build_synthetic_market_proxy(price_dict, exclude=None):
    exclude = set(exclude or [])
    closes = [px["Close"].rename(t) for t, px in price_dict.items() if t not in exclude]
    m = pd.concat(closes, axis=1).sort_index()
    ret = m.pct_change(fill_method=None)
    ewret = ret.mean(axis=1, skipna=True).fillna(0)
    close = 1000 * (1 + ewret).cumprod()
    return pd.DataFrame({"Open": close, "High": close, "Low": close, "Close": close, "Volume": 0.0}, index=close.index)


# ==========================================================
# 3. TIME-SERIES TRAIN / VALIDATION / TEST SPLIT
# ==========================================================
def resolve_auto_splits(data, split_mode="auto", train_prop=0.60, val_prop=0.20, test_prop=0.20,
                        prefer_last_year=True, min_total=420, min_train=252, min_val=63, min_test=100):
    dates = pd.DatetimeIndex(sorted(pd.to_datetime(data["Date"].unique())))
    n = len(dates)
    if n < min_total:
        raise ValueError(f"Bộ dữ liệu chỉ có {n} phiên. Cần tối thiểu {min_total} phiên để chia tách.")

    if split_mode == "proportions":
        i_train = max(min_train, int(n * train_prop))
        i_val_end = max(i_train + min_val, int(n * (train_prop + val_prop)))
        i_val_end = min(i_val_end, n - min_test)
        if i_val_end <= i_train:
            raise ValueError("Không đủ dữ liệu để chia Validation/Test")
        core_end = dates[i_train - 1]
        val_start, val_end = dates[i_train], dates[i_val_end - 1]
        test_start, test_end = dates[i_val_end], dates[-1]
        return dates, core_end, val_start, val_end, test_start, test_end

    # Chế độ Auto: ưu tiên dùng năm trọn vẹn cuối cùng làm Final Test
    last_year = dates[-1].year
    last_year_dates = dates[dates.year == last_year]
    pre_last_year_count = int((dates < last_year_dates[0]).sum()) if len(last_year_dates) else 0

    if (prefer_last_year and len(last_year_dates) >= 180 and pre_last_year_count >= min_train + min_val):
        test_start, test_end = last_year_dates[0], last_year_dates[-1]
        pretest = dates[dates < test_start]
    else:
        test_n = max(min_test, min(252, int(round(test_prop * n))))
        test_start, test_end = dates[-test_n], dates[-1]
        pretest = dates[dates < test_start]

    val_n = max(min_val, min(252, int(round(0.25 * len(pretest)))))
    val_n = min(val_n, len(pretest) - min_train)
    if val_n < min_val:
        raise ValueError("Không đủ lịch sử trước Test để tách Validation riêng")

    val_start, val_end = pretest[-val_n], pretest[-1]
    core_end = pretest[-val_n - 1]
    return dates, core_end, val_start, val_end, test_start, test_end


# ==========================================================
# 4. TECHNICAL INDICATORS & TA SCORE (0 - 10)
# ==========================================================
def ema(s, span):
    return s.ewm(span=span, adjust=False, min_periods=span).mean()

def rsi_wilder(close, window=14):
    d = close.diff()
    gain, loss = d.clip(lower=0), -d.clip(upper=0)
    ag = gain.ewm(alpha=1/window, adjust=False, min_periods=window).mean()
    al = loss.ewm(alpha=1/window, adjust=False, min_periods=window).mean()
    rs = ag / al.replace(0, np.nan)
    rsi = 100 - 100/(1+rs)
    rsi.loc[(al == 0) & (ag > 0)] = 100
    rsi.loc[(ag == 0) & (al > 0)] = 0
    return rsi

def atr_wilder(df, window=14):
    prev = df["Close"].shift(1)
    tr = pd.concat([
        df["High"] - df["Low"],
        (df["High"] - prev).abs(),
        (df["Low"] - prev).abs()
    ], axis=1).max(axis=1)
    return tr.ewm(alpha=1/window, adjust=False, min_periods=window).mean()

def adx_wilder(df, window=14):
    high, low = df["High"], df["Low"]
    up, down = high.diff(), -low.diff()
    plus_dm = pd.Series(np.where((up > down) & (up > 0), up, 0.0), index=df.index)
    minus_dm = pd.Series(np.where((down > up) & (down > 0), down, 0.0), index=df.index)
    atr = atr_wilder(df, window)
    plus_di = 100 * plus_dm.ewm(alpha=1/window, adjust=False, min_periods=window).mean() / atr.replace(0, np.nan)
    minus_di = 100 * minus_dm.ewm(alpha=1/window, adjust=False, min_periods=window).mean() / atr.replace(0, np.nan)
    dx = 100 * (plus_di-minus_di).abs() / (plus_di+minus_di).replace(0, np.nan)
    adx = dx.ewm(alpha=1/window, adjust=False, min_periods=window).mean()
    return adx, plus_di, minus_di

def calculate_ta_features(df):
    out = df.copy().sort_index()
    out["EMA20"] = ema(out["Close"], 20)
    out["EMA50"] = ema(out["Close"], 50)
    out["EMA200"] = ema(out["Close"], 200)

    ema12, ema26 = ema(out["Close"], 12), ema(out["Close"], 26)
    out["MACD"] = ema12 - ema26
    out["MACD_Signal"] = ema(out["MACD"], 9)

    out["ADX"], out["Plus_DI"], out["Minus_DI"] = adx_wilder(out, 14)
    out["ATR"] = atr_wilder(out, 14)
    out["ATR_pct"] = 100 * out["ATR"] / out["Close"].replace(0, np.nan)
    out["ATR_pct_Q90"] = out["ATR_pct"].rolling(126, min_periods=40).quantile(0.90)

    out["Volume_MA20"] = out["Volume"].rolling(20, min_periods=20).mean()
    out["Volume_Ratio"] = out["Volume"] / out["Volume_MA20"].replace(0, np.nan)
    out["OBV"] = (np.sign(out["Close"].diff()).fillna(0) * out["Volume"]).cumsum()
    out["OBV_EMA20"] = ema(out["OBV"], 20)

    out["RSI"] = rsi_wilder(out["Close"], 14)
    out["ROC10"] = out["Close"].pct_change(10) * 100
    out["High_52W"] = out["High"].rolling(252, min_periods=150).max()
    out["Close_to_52W_High"] = out["Close"] / out["High_52W"].replace(0, np.nan)
    out["Avg_Traded_Value_60"] = (out["Close"] * out["Volume"]).rolling(60, min_periods=20).mean()

    out["Momentum63"] = out["Close"].pct_change(63)
    out["Momentum126"] = out["Close"].pct_change(126)
    out["Momentum252"] = out["Close"].pct_change(252)
    out["Vol63"] = out["Close"].pct_change().rolling(63, min_periods=40).std() * np.sqrt(TRADING_DAYS)
    rollmax = out["Close"].rolling(126, min_periods=60).max()
    out["DD126"] = out["Close"] / rollmax - 1

    # TA Score 0-10
    out["Trend_Score"] = (
        (out["Close"] > out["EMA50"]).astype(int)*2 +
        (out["EMA50"] > out["EMA200"]).astype(int)*2 +
        (out["MACD"] > out["MACD_Signal"]).astype(int)
    ).clip(0, 5)

    momentum_core = ((out["RSI"] >= 45) & (out["RSI"] <= 75)) | (out["ROC10"] > 0)
    out["Momentum_Score"] = (
        momentum_core.astype(int)*2 +
        (out["Close_to_52W_High"] >= 0.85).astype(int)
    ).clip(0, 3)

    out["Volume_Score"] = ((out["Volume_Ratio"] >= 1.0) | (out["OBV"] > out["OBV_EMA20"])).astype(int)
    out["Volatility_Score"] = (
        out["ATR_pct"].notna() & out["ATR_pct_Q90"].notna() &
        (out["ATR_pct"] <= out["ATR_pct_Q90"])
    ).astype(int)

    out["TA_Score"] = (
        out["Trend_Score"] + out["Momentum_Score"] +
        out["Volume_Score"] + out["Volatility_Score"]
    ).clip(0, 10).astype(int)
    return out

def stateful_active_from_score(score, buy_threshold, exit_threshold):
    state = 0
    vals = []
    for v in score.fillna(-np.inf):
        if v >= buy_threshold:
            state = 1
        elif v <= exit_threshold:
            state = 0
        vals.append(state)
    return pd.Series(vals, index=score.index, dtype=int)


# ==========================================================
# 5. METRICS & PERFORMANCE EVALUATION
# ==========================================================
def annualized_sharpe(r):
    r = pd.Series(r).replace([np.inf, -np.inf], np.nan).dropna()
    return np.sqrt(TRADING_DAYS)*r.mean()/r.std(ddof=1) if len(r) > 1 and r.std(ddof=1) > 1e-12 else np.nan

def annualized_sortino(r):
    r = pd.Series(r).replace([np.inf, -np.inf], np.nan).dropna()
    neg = r[r < 0]
    return np.sqrt(TRADING_DAYS)*r.mean()/neg.std(ddof=1) if len(neg) > 1 and neg.std(ddof=1) > 1e-12 else np.nan

def metrics_from_equity(eq):
    eq = pd.Series(eq).dropna()
    if len(eq) < 2:
        return {}
    r = eq.pct_change().dropna()
    total = eq.iloc[-1]/eq.iloc[0]-1
    years = max((eq.index[-1]-eq.index[0]).days/365.25, 1/365.25)
    cagr = (eq.iloc[-1]/eq.iloc[0])**(1/years)-1
    dd = eq/eq.cummax()-1
    mdd = dd.min()
    vol = r.std(ddof=1)*np.sqrt(TRADING_DAYS) if len(r) > 1 else np.nan
    return {
        "Return [%]": total*100,
        "CAGR [%]": cagr*100,
        "Sharpe": annualized_sharpe(r),
        "Sortino": annualized_sortino(r),
        "Volatility [%]": vol*100 if pd.notna(vol) else np.nan,
        "MaxDD [%]": mdd*100 if pd.notna(mdd) else np.nan,
        "Calmar": cagr/abs(mdd) if pd.notna(mdd) and mdd < 0 else np.nan,
    }

def simulate_single_stock(features, start, end, buy_threshold, exit_threshold,
                          initial_cash=1_000_000.0, commission=0.0015, slippage=0.0005):
    """Backtest 1 mã: tín hiệu Close(t) giao dịch tại Open(t+1)."""
    active_close = stateful_active_from_score(features["TA_Score"], buy_threshold, exit_threshold)
    active_open = active_close.shift(1).fillna(0).astype(int)
    x = features.loc[(features.index >= pd.Timestamp(start)) & (features.index <= pd.Timestamp(end))].copy()
    active_open = active_open.reindex(x.index).fillna(0).astype(int)
    if x.empty:
        return {}

    cash, shares, state, trades = float(initial_cash), 0.0, 0, 0
    eq = []
    for dt, row in x.iterrows():
        d = int(active_open.loc[dt])
        op, cl = float(row["Open"]), float(row["Close"])
        if d != state:
            if d == 1:
                px = op*(1+slippage)
                qty = math.floor(cash/(px*(1+commission)))
                if qty > 0:
                    cash -= qty*px*(1+commission)
                    shares = qty
                    state = 1
                    trades += 1
            else:
                px = op*(1-slippage)
                cash += shares*px*(1-commission)
                shares = 0
                state = 0
                trades += 1
        eq.append((dt, cash + shares*cl))

    m = metrics_from_equity(pd.Series(dict(eq)).sort_index())
    m["Trades"] = trades
    return m


# ==========================================================
# 6. UNIVERSE SELECTION & THRESHOLD OPTIMIZATION (CORE TRAIN)
# ==========================================================
def build_universe(price_dict, asof, benchmark_ticker=None,
                   min_history_floor=180, min_history_cap=400,
                   liquidity_lookback=60, liquidity_min_percentile=0.20,
                   use_liquidity_filter=True):
    all_dates = set()
    for px in price_dict.values():
        all_dates.update(px.loc[:asof].index)
    pre_sessions = len(all_dates)
    min_hist = min(min_history_cap, max(min_history_floor, int(0.55*pre_sessions)))

    rows = []
    for t, px in price_dict.items():
        if t == benchmark_ticker:
            continue
        h = px.loc[:asof]
        if h.empty:
            continue
        avg_value = (h["Close"]*h["Volume"]).tail(liquidity_lookback).mean()
        rows.append({
            "Ticker": t,
            "HistorySessions": len(h),
            "AvgTradedValue": avg_value,
            "EligibleHistory": len(h) >= min_hist,
        })

    u = pd.DataFrame(rows)
    if u.empty:
        return u
    ok = u["EligibleHistory"] & u["AvgTradedValue"].notna()
    u.loc[ok, "LiquidityPercentile"] = u.loc[ok, "AvgTradedValue"].rank(pct=True)
    u["PassLiquidity"] = (
        u["LiquidityPercentile"].fillna(0) >= liquidity_min_percentile
        if use_liquidity_filter else True
    )
    u["Candidate"] = u["EligibleHistory"] & u["PassLiquidity"]
    return u.sort_values("AvgTradedValue", ascending=False).reset_index(drop=True)

def make_equal_folds(dates, start, end, n_folds=3):
    d = pd.DatetimeIndex(dates)
    d = d[(d >= pd.Timestamp(start)) & (d <= pd.Timestamp(end))]
    if len(d) < n_folds*20:
        n_folds = max(1, len(d)//20)
    chunks = np.array_split(d, n_folds)
    folds = []
    for i, c in enumerate(chunks, 1):
        if len(c) >= 20:
            folds.append({"Fold": f"F{i}", "start": pd.Timestamp(c[0]), "end": pd.Timestamp(c[-1])})
    return folds

def choose_global_ta_thresholds(feature_store, candidates, folds,
                                buy_choices=[5, 6, 7], exit_choices=[2, 3, 4],
                                stability_penalty=0.20):
    trials = []
    for b, e in product(buy_choices, exit_choices):
        if e >= b:
            continue
        stock_scores = []
        for t in candidates:
            ms = []
            for f in folds:
                m = simulate_single_stock(feature_store[t], f["start"], f["end"], b, e)
                if m:
                    sh = m.get("Sharpe", np.nan)
                    sh = -3.0 if not np.isfinite(sh) else float(sh)
                    ms.append((sh, float(m.get("Return [%]", 0)), float(m.get("MaxDD [%]", -100))))
            if ms:
                a = np.array(ms, float)
                stock_scores.append([np.nanmean(a[:,0]), np.nanmean(a[:,1]), np.nanmin(a[:,2])])
        if not stock_scores:
            continue

        z = np.array(stock_scores, float)
        med_sh = float(np.nanmedian(z[:,0]))
        sd_sh = float(np.nanstd(z[:,0]))
        med_ret = float(np.nanmedian(z[:,1]))
        med_dd = float(np.nanmedian(z[:,2]))
        score = med_sh - stability_penalty*sd_sh + 0.005*med_ret + 0.003*med_dd
        trials.append({
            "BuyThreshold": b,
            "ExitThreshold": e,
            "GlobalCVScore": score,
            "MedianSharpe": med_sh,
            "SharpeStd": sd_sh,
            "MedianReturnPct": med_ret,
            "MedianMaxDDPct": med_dd,
            "NStocks": len(stock_scores),
        })

    tr = pd.DataFrame(trials).sort_values("GlobalCVScore", ascending=False).reset_index(drop=True)
    if tr.empty:
        return 6, 3, tr
    best = tr.iloc[0]
    return int(best["BuyThreshold"]), int(best["ExitThreshold"]), tr


# ==========================================================
# 7. CROSS-SECTIONAL RANKING & TOP N SELECTION
# ==========================================================
def trailing_stock_quality(features, asof, b, e, lookback=126):
    h = features.loc[:asof]
    if len(h) < 40:
        return {"TrailingSharpe": -3.0, "TrailingReturnPct": -100.0, "TrailingMaxDDPct": -100.0}
    sub_idx = h.index[-min(lookback, len(h)):]
    m = simulate_single_stock(features, sub_idx[0], sub_idx[-1], b, e)
    sh = m.get("Sharpe", np.nan) if m else np.nan
    return {
        "TrailingSharpe": -3.0 if not np.isfinite(sh) else float(sh),
        "TrailingReturnPct": float(m.get("Return [%]", -100.0)) if m else -100.0,
        "TrailingMaxDDPct": float(m.get("MaxDD [%]", -100.0)) if m else -100.0,
    }

def rank_and_select_asof(feature_store, price_dict, benchmark_features, benchmark_ticker,
                         asof, b, e, top_n=5,
                         rank_w={"trailing":0.25, "ta":0.20, "rs":0.25, "low_vol":0.10, "liq":0.05, "mom":0.15}):
    universe = build_universe(price_dict, asof, benchmark_ticker)
    if universe.empty:
        raise ValueError(f"Universe rỗng tại {asof}")
    candidates = universe.loc[universe["Candidate"], "Ticker"].tolist()
    bench = benchmark_features.loc[:asof]
    bench_m126 = bench["Close"].pct_change(126).iloc[-1] if len(bench) > 126 else np.nan
    uidx = universe.set_index("Ticker")

    rows = []
    for t in candidates:
        feat = feature_store[t].loc[:asof]
        if feat.empty:
            continue
        snap = feat.iloc[-1]
        tq = trailing_stock_quality(feature_store[t], asof, b, e)
        rows.append({
            "Ticker": t,
            **tq,
            "TA_Score": float(snap.get("TA_Score", np.nan)),
            "Momentum63": float(snap.get("Momentum63", np.nan)),
            "Momentum126": float(snap.get("Momentum126", np.nan)),
            "Momentum252": float(snap.get("Momentum252", np.nan)),
            "RelativeStrength126": (
                float(snap.get("Momentum126", np.nan)-bench_m126)
                if pd.notna(bench_m126) else float(snap.get("Momentum126", np.nan))
            ),
            "Vol63": float(snap.get("Vol63", np.nan)),
            "RSI": float(snap.get("RSI", np.nan)),
            "LiquidityPercentile": float(uidx.loc[t, "LiquidityPercentile"]),
        })

    r = pd.DataFrame(rows)
    if r.empty:
        raise ValueError(f"Không có cổ phiếu đủ điều kiện tại {asof}")

    # Gate 1: TA_Score >= 4
    r["PassTA"] = r["TA_Score"].fillna(0) >= 4
    if int(r["PassTA"].sum()) >= top_n:
        r = r[r["PassTA"]].copy()

    # Gate 2: Rủi ro (Vol63 <= 75%, MaxDD >= 15%, RSI <= 80)
    vol_cap = r["Vol63"].quantile(0.75)
    dd_floor = r["TrailingMaxDDPct"].quantile(0.15)
    r["PassRiskGate"] = (
        (r["Vol63"].isna() | (r["Vol63"] <= vol_cap)) &
        (r["TrailingMaxDDPct"].isna() | (r["TrailingMaxDDPct"] >= dd_floor)) &
        (r["RSI"].isna() | (r["RSI"] <= 80.0))
    )
    if int(r["PassRiskGate"].sum()) >= top_n:
        r = r[r["PassRiskGate"]].copy()

    # Xếp hạng Percentile
    r["P_Trailing"] = r["TrailingSharpe"].rank(pct=True)
    r["P_TA"] = r["TA_Score"].rank(pct=True)
    r["P_RS"] = r["RelativeStrength126"].rank(pct=True)
    r["P_LowVol"] = (-r["Vol63"]).rank(pct=True)
    r["P_Liq"] = r["LiquidityPercentile"].rank(pct=True)
    mommix = 0.6*r["Momentum126"].fillna(r["Momentum63"]) + 0.4*r["Momentum63"].fillna(0)
    r["P_Momentum"] = mommix.rank(pct=True)

    r["TA_Composite"] = (
        rank_w["trailing"]*r["P_Trailing"] +
        rank_w["ta"]*r["P_TA"] +
        rank_w["rs"]*r["P_RS"] +
        rank_w["low_vol"]*r["P_LowVol"] +
        rank_w["liq"]*r["P_Liq"] +
        rank_w["mom"]*r["P_Momentum"]
    ) * 100
    r["FinalScore"] = r["TA_Composite"]

    r = r.sort_values(["FinalScore", "TrailingSharpe"], ascending=False).reset_index(drop=True)

    # Đa dạng hóa: correlation cap <= 0.80
    corr_tickers = r["Ticker"].tolist()
    close_list = [feature_store[t].loc[:asof, "Close"].rename(t) for t in corr_tickers]
    corr_px = pd.concat(close_list, axis=1, join="inner").tail(127) if close_list else pd.DataFrame()
    corr = corr_px.pct_change(fill_method=None).dropna().corr() if len(corr_px) > 20 else pd.DataFrame()

    selected = []
    for _, row in r.iterrows():
        t = row["Ticker"]
        if selected and not corr.empty:
            too_close = any(
                t in corr.index and x in corr.columns and
                pd.notna(corr.loc[t, x]) and abs(corr.loc[t, x]) > 0.80
                for x in selected
            )
            if too_close:
                continue
        selected.append(t)
        if len(selected) >= top_n:
            break

    if len(selected) < top_n:
        for t in r["Ticker"]:
            if t not in selected:
                selected.append(t)
            if len(selected) >= top_n:
                break

    locked = r[r["Ticker"].isin(selected)].copy()
    locked["SelectedOrder"] = locked["Ticker"].map({t:i+1 for i,t in enumerate(selected)})
    locked["AsOfDate"] = pd.Timestamp(asof)
    locked = locked.sort_values("SelectedOrder")
    return r, locked, selected


# ==========================================================
# 8. SCHEDULE GENERATION & PORTFOLIO ALLOCATION MODELS
# ==========================================================
def first_session_each_period(calendar, start, end, freq="Q"):
    idx = pd.DatetimeIndex(calendar)
    idx = idx[(idx >= pd.Timestamp(start)) & (idx <= pd.Timestamp(end))]
    if len(idx) == 0:
        return []
    s = pd.Series(idx, index=idx)
    key = idx.to_period(freq)
    firsts = s.groupby(key).first().tolist()
    if idx[0] not in firsts:
        firsts = [idx[0]] + firsts
    return sorted(set(pd.Timestamp(x) for x in firsts))

def previous_session(calendar, dt):
    idx = pd.DatetimeIndex(calendar)
    prev = idx[idx < pd.Timestamp(dt)]
    return pd.Timestamp(prev[-1]) if len(prev) else None

def build_selection_schedule(feature_store, price_dict, benchmark_ticker, benchmark_features,
                             calendar, start, end, b, e, top_n=5, selection_freq="Q"):
    effective_dates = first_session_each_period(calendar, start, end, selection_freq)
    rows = []
    details = {}

    for eff in effective_dates:
        asof = previous_session(calendar, eff)
        if asof is None:
            continue
        ranking, locked, selected = rank_and_select_asof(
            feature_store, price_dict, benchmark_features, benchmark_ticker,
            asof, b, e, top_n=top_n
        )
        details[pd.Timestamp(eff)] = {"ranking": ranking, "locked": locked}
        row = {
            "EffectiveDate": pd.Timestamp(eff),
            "AsOfDate": pd.Timestamp(asof),
        }
        for i in range(top_n):
            row[f"Top{i+1}"] = selected[i] if i < len(selected) else None
        rows.append(row)

    schedule = pd.DataFrame(rows).sort_values("EffectiveDate").reset_index(drop=True)
    if schedule.empty:
        raise ValueError("Không tạo được selection schedule")
    return schedule, details

def selected_tickers_asof(selection_schedule, effective_date, top_n=5):
    x = selection_schedule[selection_schedule["EffectiveDate"] <= pd.Timestamp(effective_date)]
    if x.empty:
        return []
    row = x.iloc[-1]
    return [row[f"Top{i+1}"] for i in range(top_n) if pd.notna(row.get(f"Top{i+1}"))]

def close_matrix(price_dict, tickers, end=None, start=None):
    series = [price_dict[t]["Close"].rename(t) for t in tickers if t in price_dict]
    if not series:
        return pd.DataFrame()
    m = pd.concat(series, axis=1, join="inner").sort_index().dropna(how="any")
    if start is not None:
        m = m.loc[m.index >= pd.Timestamp(start)]
    if end is not None:
        m = m.loc[m.index <= pd.Timestamp(end)]
    return m

def equal_weights(tickers):
    tickers = list(tickers)
    return pd.Series(1/len(tickers), index=tickers, dtype=float) if tickers else pd.Series(dtype=float)

def robust_mu(prices, lookback=252, mu_shrinkage=0.50):
    r = prices.pct_change(fill_method=None).dropna().tail(lookback)
    if r.empty:
        return pd.Series(0.0, index=prices.columns)
    span = max(20, min(lookback, len(r)))
    raw = r.ewm(span=span, adjust=False).mean().iloc[-1] * TRADING_DAYS
    grand = float(raw.mean())
    return mu_shrinkage*raw + (1-mu_shrinkage)*grand

def lw_cov(prices, lookback=252):
    r = prices.pct_change(fill_method=None).dropna().tail(lookback)
    if len(r) < 30:
        return r.cov()*TRADING_DAYS
    est = LedoitWolf().fit(r.values)
    return pd.DataFrame(est.covariance_*TRADING_DAYS, index=r.columns, columns=r.columns)

def max_sharpe_weights(mu, cov, max_weight=0.50, l2_gamma=0.03):
    tickers = list(mu.index)
    n = len(tickers)
    if n == 0:
        return pd.Series(dtype=float)
    max_weight = max(max_weight, 1/n)
    muv = mu.values.astype(float)
    cv = cov.reindex(index=tickers, columns=tickers).fillna(0).values.astype(float) + np.eye(n)*1e-8

    def obj(w):
        ret = float(w@muv)
        vol = math.sqrt(max(float(w@cv@w), 1e-12))
        return -(ret/vol) + l2_gamma*float(np.sum(w*w))

    x0 = np.repeat(1/n, n)
    bounds = [(0, max_weight)]*n
    cons = ({"type":"eq", "fun":lambda w:np.sum(w)-1},)
    res = minimize(obj, x0, method="SLSQP", bounds=bounds, constraints=cons,
                   options={"maxiter":1000, "ftol":1e-10})
    if not res.success or np.any(~np.isfinite(res.x)):
        return equal_weights(tickers)
    w = np.clip(res.x, 0, None)
    return pd.Series(w/w.sum(), index=tickers) if w.sum() > 0 else equal_weights(tickers)

def min_variance_weights(cov, max_weight=0.50):
    tickers = list(cov.index)
    n = len(tickers)
    if n == 0:
        return pd.Series(dtype=float)
    max_weight = max(max_weight, 1/n)
    cv = cov.values.astype(float) + np.eye(n)*1e-8

    def obj(w):
        return float(w@cv@w)

    x0 = np.repeat(1/n, n)
    bounds = [(0, max_weight)]*n
    cons = ({"type":"eq", "fun":lambda w:np.sum(w)-1},)
    res = minimize(obj, x0, method="SLSQP", bounds=bounds, constraints=cons,
                   options={"maxiter":1000, "ftol":1e-12})
    if not res.success:
        return equal_weights(tickers)
    w = np.clip(res.x, 0, None)
    return pd.Series(w/w.sum(), index=tickers)

def _cluster_var(cov, items):
    sub = cov.loc[items, items].values
    ivp = 1/np.diag(sub)
    ivp = ivp/ivp.sum()
    return float(ivp@sub@ivp)

def _quasi_diag(link):
    link = link.astype(int)
    sort_ix = pd.Series([link[-1,0], link[-1,1]])
    num_items = int(link[-1,3])
    while sort_ix.max() >= num_items:
        sort_ix.index = range(0, sort_ix.shape[0]*2, 2)
        df0 = sort_ix[sort_ix >= num_items]
        i = df0.index
        j = df0.values-num_items
        sort_ix.loc[i] = link[j,0]
        df1 = pd.Series(link[j,1], index=i+1)
        sort_ix = pd.concat([sort_ix, df1]).sort_index()
        sort_ix.index = range(sort_ix.shape[0])
    return sort_ix.tolist()

def hrp_weights(prices, method="single"):
    r = prices.pct_change(fill_method=None).dropna()
    if len(r) < 30 or prices.shape[1] < 2:
        return equal_weights(prices.columns)
    cov = r.cov()*TRADING_DAYS
    corr = r.corr().clip(-1,1)
    dist = np.sqrt((1-corr)/2)
    link = linkage(squareform(dist.values, checks=False), method=method)
    ordered = corr.index[_quasi_diag(link)].tolist()

    w = pd.Series(1.0, index=ordered)
    clusters = [ordered]
    while clusters:
        nxt = []
        for c in clusters:
            if len(c) <= 1:
                continue
            s = len(c)//2
            c0, c1 = c[:s], c[s:]
            v0, v1 = _cluster_var(cov, c0), _cluster_var(cov, c1)
            a = 1-v0/(v0+v1) if v0+v1 > 0 else .5
            w[c0] *= a
            w[c1] *= (1-a)
            nxt += [c0, c1]
        clusters = nxt
    return w.reindex(prices.columns).fillna(0)/w.sum()

def compute_weights(model_family, prices, lookback=252):
    tickers = list(prices.columns)
    ew = equal_weights(tickers)
    if len(tickers) == 0:
        return ew
    if model_family == "EW":
        return ew

    cov = lw_cov(prices, lookback)
    if model_family == "MINVAR_LW":
        return min_variance_weights(cov)
    if model_family == "HRP":
        return hrp_weights(prices.tail(lookback))

    mu = robust_mu(prices, lookback)
    mpt = max_sharpe_weights(mu, cov)
    if model_family == "MPT_LW":
        return mpt
    if model_family == "BLEND50":
        w = 0.50*mpt + 0.50*ew
        return w/w.sum()
    raise ValueError("Model family không hỗ trợ: " + str(model_family))

def build_dynamic_weight_schedule(price_dict, selection_schedule, calendar, start, end,
                                  model_family, allocation_freq="M", top_n=5, lookback=252):
    rebalance_dates = first_session_each_period(calendar, start, end, allocation_freq)
    rows = []
    all_tickers = set()

    for eff in rebalance_dates:
        asof = previous_session(calendar, eff)
        if asof is None:
            continue
        top = selected_tickers_asof(selection_schedule, eff, top_n)
        if not top:
            continue
        all_tickers.update(top)
        prices = close_matrix(price_dict, top, end=asof)
        if prices.empty or len(prices) < 30:
            w = equal_weights(top)
        else:
            w = compute_weights(model_family, prices.tail(lookback+1), lookback)
        row = {"EffectiveDate": pd.Timestamp(eff), "AsOfDate": pd.Timestamp(asof)}
        row.update({t: float(w.get(t, 0.0)) for t in top})
        rows.append(row)

    if not rows:
        raise ValueError("Không tạo được weight schedule")

    meta = pd.DataFrame(rows).sort_values("EffectiveDate").reset_index(drop=True)
    tickers = sorted(all_tickers)
    weights = meta.set_index("EffectiveDate").reindex(columns=tickers).fillna(0.0)
    return weights, meta[["EffectiveDate", "AsOfDate"]]

def build_static_buyhold_schedule(selection_schedule, start, top_n=5):
    start = pd.Timestamp(start)
    top = selected_tickers_asof(selection_schedule, start, top_n)
    if not top:
        raise ValueError("Không có Top N ở đầu kỳ cho BuyHold")
    return pd.DataFrame([equal_weights(top)], index=[start])


# ==========================================================
# 9. MARKET REGIME & EVENT-DRIVEN PORTFOLIO SIMULATION ENGINE
# ==========================================================
def build_market_regime(benchmark_features, ema_span=200, slope_lookback=20, use_slope_confirm=True):
    x = benchmark_features.copy()
    ema_m = x["Close"].ewm(span=ema_span, adjust=False, min_periods=ema_span).mean()
    slope = ema_m - ema_m.shift(slope_lookback)
    risk_close = x["Close"] > ema_m
    if use_slope_confirm:
        risk_close = risk_close & (slope > 0)
    return pd.DataFrame({
        "EMA_Market": ema_m,
        "EMA_Slope": slope,
        "RiskOn_Close": risk_close.fillna(False),
        "RiskOn_Open": risk_close.shift(1).fillna(False),
    }, index=x.index)

def build_stock_open_masks(feature_store, tickers, b, e):
    masks = {}
    for t in tickers:
        active_close = stateful_active_from_score(feature_store[t]["TA_Score"], b, e)
        masks[t] = active_close.shift(1).fillna(0).astype(bool)
    return pd.DataFrame(masks).sort_index()

def run_portfolio_engine_dynamic(price_dict, weight_schedule, calendar, start, end,
                                 market_regime_open=None, stock_active_open=None,
                                 use_market_regime=True, use_stock_mask=True,
                                 individual_trailing_stop=False, individual_stop_pct=0.12,
                                 portfolio_trailing_stop=False, portfolio_stop_pct=0.15,
                                 initial_cash=1_000_000_000.0, commission=0.0015,
                                 slippage=0.0005, fractional=False,
                                 cash_policy="cash40"):
    tickers = list(weight_schedule.columns)
    idx = pd.DatetimeIndex(calendar)
    idx = idx[(idx >= pd.Timestamp(start)) & (idx <= pd.Timestamp(end))]
    if len(idx) == 0:
        raise ValueError("Không có ngày giao dịch trong kỳ")

    opens = pd.DataFrame(index=idx, columns=tickers, dtype=float)
    closes = pd.DataFrame(index=idx, columns=tickers, dtype=float)
    for t in tickers:
        if t in price_dict:
            opens[t] = price_dict[t]["Open"].reindex(idx)
            closes[t] = price_dict[t]["Close"].reindex(idx)

    ws = weight_schedule.sort_index().reindex(columns=tickers)
    ws = ws.reindex(ws.index.union(idx)).sort_index().ffill().reindex(idx).fillna(0.0)

    regime = (
        pd.Series(True, index=idx)
        if market_regime_open is None
        else pd.Series(market_regime_open).reindex(idx).ffill().fillna(False).astype(bool)
    )
    masks = (
        pd.DataFrame(True, index=idx, columns=tickers)
        if stock_active_open is None
        else stock_active_open.reindex(index=idx, columns=tickers).fillna(False).astype(bool)
    )

    cash = float(initial_cash)
    shares = pd.Series(0.0, index=tickers)
    last_close = pd.Series(np.nan, index=tickers)
    last_target = pd.Series(np.nan, index=tickers)
    entry_peak = pd.Series(np.nan, index=tickers)
    stop_blocked = pd.Series(False, index=tickers)
    portfolio_peak = initial_cash
    portfolio_blocked = False

    orders, eq_rows, exposure_rows = [], [], []
    turnover, rebals = 0.0, 0

    for dt in idx:
        op = opens.loc[dt].astype(float)
        cl = closes.loc[dt].astype(float)

        valid_close = cl.notna()
        last_close.loc[valid_close] = cl.loc[valid_close]
        mark_open = op.where(op.notna(), last_close)
        mark_close = cl.where(cl.notna(), last_close)

        base = ws.loc[dt].clip(lower=0)
        if base.sum() > 1 + 1e-8:
            base = base/base.sum()

        risk_on = bool(regime.loc[dt]) if use_market_regime else True
        active = masks.loc[dt].astype(bool) if use_stock_mask else pd.Series(True, index=tickers)
        for t in tickers:
            if not active[t]:
                stop_blocked[t] = False

        target = base.copy()
        if use_stock_mask:
            target *= active.astype(float)
        if individual_trailing_stop:
            target[stop_blocked] = 0.0
        if use_market_regime and not risk_on:
            target[:] = 0.0
        if portfolio_trailing_stop and portfolio_blocked:
            target[:] = 0.0

        # Thí nghiệm chính sách tiền mặt: original | cash40 | full
        target = target.clip(lower=0)
        if target.sum() > 1 + 1e-8:
            target /= target.sum()

        min_exposure = {"original": 0.0, "cash40": 0.60, "full": 1.0}.get(cash_policy, 0.60)
        base_total = float(base.sum())
        current_total = float(target.sum())
        if min_exposure > current_total + 1e-10 and base_total > current_total + 1e-10:
            feasible_exposure = min(min_exposure, base_total)
            alpha = (feasible_exposure - current_total) / (base_total - current_total)
            target = (1-alpha)*target + alpha*base
            target = target.clip(lower=0)

        changed = last_target.isna().any() or (target-last_target.fillna(0)).abs().max() > 1e-10
        if changed:
            equity_open = cash + float((shares * mark_open.fillna(0)).sum())
            desired = shares.copy()

            for t in tickers:
                if pd.isna(op[t]) or op[t] <= 0:
                    continue
                if target[t] <= 0:
                    desired[t] = 0.0
                else:
                    px = op[t]*(1+slippage)
                    raw_qty = equity_open*target[t]/(px*(1+commission))
                    desired[t] = raw_qty if fractional else math.floor(raw_qty)

            # Bán trước
            for t in tickers:
                d = desired[t] - shares[t]
                if d < -1e-12 and pd.notna(op[t]) and op[t] > 0:
                    qty = -d
                    px = op[t]*(1-slippage)
                    notional = qty*px
                    fee = notional*commission
                    cash += notional-fee
                    shares[t] -= qty
                    turnover += notional
                    orders.append({"Date":dt, "Ticker":t, "Side":"SELL", "Shares":qty, "ExecPrice":px, "Fee":fee, "Reason":"TARGET"})
                    if shares[t] <= 1e-12:
                        entry_peak[t] = np.nan

            # Mua sau
            for t in tickers:
                d = desired[t] - shares[t]
                if d > 1e-12 and pd.notna(op[t]) and op[t] > 0:
                    px = op[t]*(1+slippage)
                    unit = px*(1+commission)
                    qty = min(d, cash/unit) if fractional else min(d, math.floor(cash/unit))
                    if qty > 0:
                        notional = qty*px
                        fee = notional*commission
                        was_zero = shares[t] <= 1e-12
                        cash -= notional+fee
                        shares[t] += qty
                        turnover += notional
                        orders.append({"Date":dt, "Ticker":t, "Side":"BUY", "Shares":qty, "ExecPrice":px, "Fee":fee, "Reason":"TARGET"})
                        if was_zero and pd.notna(mark_close[t]):
                            entry_peak[t] = mark_close[t]

            last_target = target.copy()
            rebals += 1

        if individual_trailing_stop:
            for t in tickers:
                if shares[t] > 0 and pd.notna(mark_close[t]):
                    entry_peak[t] = max(entry_peak[t] if pd.notna(entry_peak[t]) else mark_close[t], mark_close[t])
                    if mark_close[t] <= entry_peak[t]*(1-individual_stop_pct):
                        stop_blocked[t] = True

        equity = float(cash + (shares*mark_close.fillna(0)).sum())
        portfolio_peak = max(portfolio_peak, equity)
        if portfolio_trailing_stop and equity <= portfolio_peak*(1-portfolio_stop_pct):
            portfolio_blocked = True

        invested = float((shares*mark_close.fillna(0)).sum())
        exposure = invested/equity if equity > 0 else 0.0
        exposure_rows.append(exposure)
        eq_rows.append((dt, equity))

    eq = pd.Series(dict(eq_rows)).sort_index()
    st_met = metrics_from_equity(eq)
    st_met.update({
        "#Orders": len(orders),
        "#Rebalances": rebals,
        "Turnover [%]": 100*turnover/initial_cash,
        "AvgExposure [%]": 100*np.mean(exposure_rows) if exposure_rows else 0.0,
        "AvgCash [%]": 100*(1-np.mean(exposure_rows)) if exposure_rows else 100.0,
    })
    return st_met, eq, pd.DataFrame(orders)

def benchmark_nav(benchmark_px, start, end, initial_cash=1_000_000_000.0):
    x = benchmark_px.loc[(benchmark_px.index >= pd.Timestamp(start)) & (benchmark_px.index <= pd.Timestamp(end)), "Close"].dropna()
    if x.empty:
        return pd.Series(dtype=float)
    return initial_cash*x/x.iloc[0]


# ==========================================================
# 10. MODEL EVALUATION & VALIDATION / TEST RUNNERS
# ==========================================================
ACTIVE_MODEL_SPECS = {
    "DYN_EW_REGIME":      {"family":"EW",        "regime":True, "mask":True},
    "DYN_MPT_LW_REGIME":  {"family":"MPT_LW",    "regime":True, "mask":True},
    "DYN_MINVAR_REGIME":  {"family":"MINVAR_LW", "regime":True, "mask":True},
    "DYN_HRP_REGIME":     {"family":"HRP",       "regime":True, "mask":True},
    "DYN_BLEND50_REGIME": {"family":"BLEND50",   "regime":True, "mask":True},
}

def model_validation_score(stats, benchmark_stats):
    sh = stats.get("Sharpe", np.nan)
    sh = -3 if not np.isfinite(sh) else sh
    so = stats.get("Sortino", np.nan)
    so = -3 if not np.isfinite(so) else so
    ret = stats.get("Return [%]", -100)
    dd = stats.get("MaxDD [%]", -100)
    turnover = stats.get("Turnover [%]", 0)
    bret = benchmark_stats.get("Return [%]", 0)
    return 0.45*sh + 0.15*so + 0.012*(ret-bret) + 0.004*dd - 0.0005*turnover

def evaluate_models_on_validation(price_dict, feature_store, benchmark_ticker, benchmark_features,
                                  val_selection, calendar, val_start, val_end, best_buy, best_exit,
                                  top_n=5, allocation_freq="M", cash_policy="cash40",
                                  use_regime=True, use_mask=True):
    regime = build_market_regime(benchmark_features)
    union_tickers = sorted({
        val_selection.loc[i, f"Top{j+1}"]
        for i in val_selection.index for j in range(top_n)
        if pd.notna(val_selection.loc[i, f"Top{j+1}"])
    })
    masks = build_stock_open_masks(feature_store, union_tickers, best_buy, best_exit)

    bnav = benchmark_nav(price_dict[benchmark_ticker], val_start, val_end)
    bstats = metrics_from_equity(bnav)

    rows, curves, orders, schedules = [], {"BENCHMARK": bnav}, {}, {}
    for name, spec in ACTIVE_MODEL_SPECS.items():
        sched, _ = build_dynamic_weight_schedule(
            price_dict, val_selection, calendar, val_start, val_end,
            spec["family"], allocation_freq=allocation_freq, top_n=top_n
        )
        st_res, eq, od = run_portfolio_engine_dynamic(
            price_dict, sched, calendar, val_start, val_end,
            market_regime_open=regime["RiskOn_Open"],
            stock_active_open=masks,
            use_market_regime=(use_regime and spec["regime"]),
            use_stock_mask=(use_mask and spec["mask"]),
            cash_policy=cash_policy,
        )
        st_res["Model"] = name
        st_res["ValidationScore"] = model_validation_score(st_res, bstats)
        rows.append(st_res)
        curves[name] = eq
        orders[name] = od
        schedules[name] = sched

    bstats.update({"Model":"BENCHMARK", "ValidationScore":np.nan})
    rows.append(bstats)
    res = pd.DataFrame(rows).set_index("Model")
    champion = res.drop(index="BENCHMARK")["ValidationScore"].idxmax()
    return res, champion, curves, orders, schedules

def evaluate_final_test_pipeline(price_dict, feature_store, benchmark_ticker, benchmark_features,
                                 test_selection, calendar, test_start, test_end, best_buy, best_exit,
                                 champion, top_n=5, allocation_freq="M", cash_policy="cash40",
                                 use_regime=True, use_mask=True):
    regime = build_market_regime(benchmark_features)
    union_tickers = sorted({
        test_selection.loc[i, f"Top{j+1}"]
        for i in test_selection.index for j in range(top_n)
        if pd.notna(test_selection.loc[i, f"Top{j+1}"])
    })
    masks = build_stock_open_masks(feature_store, union_tickers, best_buy, best_exit)

    bnav = benchmark_nav(price_dict[benchmark_ticker], test_start, test_end)
    rows = []
    curves = {"VNINDEX_or_MarketBenchmark": bnav}
    order_map = {}

    # 1. VN-Index / Benchmark
    bstats = metrics_from_equity(bnav)
    bstats["Model"] = "VNINDEX_or_MarketBenchmark"
    rows.append(bstats)

    # 2. Static Buy & Hold EW
    bh_sched = build_static_buyhold_schedule(test_selection, test_start, top_n=top_n)
    bh_stats, bh_eq, bh_od = run_portfolio_engine_dynamic(
        price_dict, bh_sched, calendar, test_start, test_end,
        market_regime_open=None, stock_active_open=None,
        use_market_regime=False, use_stock_mask=False,
        cash_policy="original"
    )
    bh_stats["Model"] = "STATIC_EW_BH"
    rows.append(bh_stats)
    curves["STATIC_EW_BH"] = bh_eq
    order_map["STATIC_EW_BH"] = bh_od

    # 3. Dynamic EW Không Regime
    ew_sched, _ = build_dynamic_weight_schedule(
        price_dict, test_selection, calendar, test_start, test_end,
        "EW", allocation_freq=allocation_freq, top_n=top_n
    )
    ew_stats, ew_eq, ew_od = run_portfolio_engine_dynamic(
        price_dict, ew_sched, calendar, test_start, test_end,
        market_regime_open=None, stock_active_open=None,
        use_market_regime=False, use_stock_mask=False,
        cash_policy="original"
    )
    ew_stats["Model"] = "DYN_EW_NO_REGIME"
    rows.append(ew_stats)
    curves["DYN_EW_NO_REGIME"] = ew_eq
    order_map["DYN_EW_NO_REGIME"] = ew_od

    # 4. Locked Champion & Diagnostics
    for name, spec in ACTIVE_MODEL_SPECS.items():
        sched, _ = build_dynamic_weight_schedule(
            price_dict, test_selection, calendar, test_start, test_end,
            spec["family"], allocation_freq=allocation_freq, top_n=top_n
        )
        st_res, eq, od = run_portfolio_engine_dynamic(
            price_dict, sched, calendar, test_start, test_end,
            market_regime_open=regime["RiskOn_Open"],
            stock_active_open=masks,
            use_market_regime=(use_regime and spec["regime"]),
            use_stock_mask=(use_mask and spec["mask"]),
            cash_policy=cash_policy
        )
        label = "LOCKED_CHAMPION" if name == champion else "DIAG_" + name
        st_res["Model"] = label
        st_res["OriginalModel"] = name
        rows.append(st_res)
        curves[label] = eq
        order_map[label] = od

    res = pd.DataFrame(rows).set_index("Model")
    if "STATIC_EW_BH" in res.index:
        res["ExcessCAGR_vs_StaticEW [pp]"] = res["CAGR [%]"] - res.loc["STATIC_EW_BH", "CAGR [%]"]
        res["ExcessSharpe_vs_StaticEW"] = res["Sharpe"] - res.loc["STATIC_EW_BH", "Sharpe"]
    bname = "VNINDEX_or_MarketBenchmark"
    res["ExcessCAGR_vs_Benchmark [pp]"] = res["CAGR [%]"] - res.loc[bname, "CAGR [%]"]
    res["DD_Improvement_vs_Benchmark [pp]"] = res["MaxDD [%]"] - res.loc[bname, "MaxDD [%]"]
    return res, curves, order_map


# ==========================================================
# 11. SIDEBAR CONTROLS & WORKFLOW TRIGGER
# ==========================================================
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/bullish.png", width=64)
    st.title("HOSE Quant Platform")
    st.caption("Chiến lược Định lượng Cổ phiếu & Phân bổ Danh mục V4")
    st.markdown("---")

    # 1. Nguồn dữ liệu
    st.subheader("📂 1. Nguồn Dữ Liệu")
    uploaded_file = st.file_uploader("Tải lên file Excel/CSV", type=["xlsx", "xls", "csv"])
    default_file = find_default_data_file()

    if uploaded_file is not None:
        data_source = uploaded_file
        st.success(f"Đã chọn file: {uploaded_file.name}")
    elif default_file is not None:
        data_source = default_file
        st.info(f"Đang dùng dữ liệu mặc định:\n`{Path(default_file).name}`")
    else:
        data_source = None
        st.warning("Vui lòng tải lên file dữ liệu giá để bắt đầu.")

    use_adjusted = st.checkbox("Sử dụng Giá điều chỉnh (Adjusted Close)", value=True)

    # 2. Cấu hình chiến lược
    st.subheader("⚙️ 2. Cấu Hình Chiến Lược")
    top_n = st.slider("Số lượng cổ phiếu Top N", min_value=3, max_value=10, value=5, step=1)
    selection_freq = st.selectbox("Tần suất chọn lại Top N", options=["Q", "M"], format_func=lambda x: "Theo Quý (Quarterly)" if x=="Q" else "Theo Tháng (Monthly)", index=0)
    allocation_freq = st.selectbox("Tần suất phân bổ tỷ trọng", options=["M", "Q"], format_func=lambda x: "Theo Tháng (Monthly)" if x=="M" else "Theo Quý (Quarterly)", index=0)

    cash_policy = st.selectbox(
        "Chính sách tiền mặt (Cash Policy)",
        options=["cash40", "original", "full"],
        format_func=lambda x: {
            "cash40": "Cash 40% (Giải ngân tối thiểu 60%)",
            "original": "Original (Tuân thủ triệt để Risk Overlay)",
            "full": "Full Equity (100% cổ phiếu - Không giữ tiền mặt)"
        }[x],
        index=0,
        help="Thí nghiệm quản trị rủi ro tiền mặt sau bộ lọc Market Regime."
    )

    # 3. Quản trị rủi ro
    st.subheader("🛡️ 3. Bộ Lọc Rủi Ro (Risk Overlay)")
    use_regime = st.checkbox("Bật Market Regime (EMA200 VNINDEX + Độ dốc)", value=True)
    use_mask = st.checkbox("Bật Stock Active Mask (Ngưỡng TA Score)", value=True)

    # 4. Tham số giao dịch
    st.subheader("💰 4. Chi Phí & Vốn")
    initial_cash = st.number_input("Vốn ban đầu (VNĐ)", value=1_000_000_000, step=100_000_000, format="%d")
    commission = st.number_input("Phí giao dịch (%)", value=0.15, step=0.01) / 100.0
    slippage = st.number_input("Trượt giá (%)", value=0.05, step=0.01) / 100.0

    st.markdown("---")
    run_btn = st.button("🚀 CHẠY KIỂM ĐỊNH CHIẾN LƯỢC", type="primary", use_container_width=True)


# ==========================================================
# 12. MAIN APP PIPELINE EXECUTION
# ==========================================================
if data_source is None:
    st.title("📈 Nền Tảng Kiểm Định Chiến Lược Định Lượng HOSE")
    st.info("👋 Chào mừng bạn! Hãy tải lên file dữ liệu OHLCV ở thanh bên trái hoặc đặt file dữ liệu vào thư mục để bắt đầu kiểm định chiến lược.")
    st.stop()

# Đọc và chuẩn hóa dữ liệu
try:
    with st.spinner("Đang chuẩn hóa và xử lý dữ liệu OHLCV..."):
        data, raw_df = load_and_standardize_data(data_source, use_adjusted=use_adjusted)
        price_dict = build_price_dict(data)

        # Benchmark
        benchmark_ticker = next((t for t in price_dict if t in BENCHMARK_ALIASES), None)
        if benchmark_ticker is None:
            benchmark_ticker = "__MARKET_PROXY__"
            price_dict[benchmark_ticker] = build_synthetic_market_proxy(price_dict)

        # Tính toán Technical Indicators cho toàn bộ mã
        feature_store = {t: calculate_ta_features(px) for t, px in price_dict.items()}

        # Chia tập thời gian
        dates, core_end, val_start, val_end, test_start, test_end = resolve_auto_splits(data)
except Exception as e:
    st.error(f"Lỗi khi đọc và tiền xử lý dữ liệu: {e}")
    st.stop()

# Tự động chạy lần đầu nếu chưa có kết quả trong session_state
should_run = run_btn or ("backtest_results" not in st.session_state)

if should_run:
    with st.status("🚀 Đang thực thi quy trình kiểm định định lượng...", expanded=True) as status:
        st.write("1️⃣ Đang tối ưu ngưỡng TA BUY/EXIT chung trên tập Core Train...")
        core_universe = build_universe(price_dict, core_end, benchmark_ticker)
        core_candidates = core_universe.loc[core_universe["Candidate"], "Ticker"].tolist()
        core_dates = dates[dates <= core_end]
        inner_n = max(63, min(126, int(0.30*len(core_dates))))
        inner_start = core_dates[-inner_n]
        threshold_folds = make_equal_folds(core_dates, inner_start, core_end, 3)

        best_buy, best_exit, threshold_trials = choose_global_ta_thresholds(
            feature_store, core_candidates, threshold_folds
        )

        st.write(f"🔒 Đã khóa ngưỡng tối ưu: BUY >= {best_buy} | EXIT <= {best_exit}")

        st.write("2️⃣ Đang tạo lịch Walk-Forward Top N cho tập Validation & Test...")
        val_selection, val_details = build_selection_schedule(
            feature_store, price_dict, benchmark_ticker, feature_store[benchmark_ticker],
            dates, val_start, val_end, best_buy, best_exit, top_n=top_n, selection_freq=selection_freq
        )

        test_selection, test_details = build_selection_schedule(
            feature_store, price_dict, benchmark_ticker, feature_store[benchmark_ticker],
            dates, test_start, test_end, best_buy, best_exit, top_n=top_n, selection_freq=selection_freq
        )

        st.write("3️⃣ Đang đánh giá các mô hình phân bổ trên Validation để chọn Champion...")
        val_res, champion, val_curves, val_orders, val_scheds = evaluate_models_on_validation(
            price_dict, feature_store, benchmark_ticker, feature_store[benchmark_ticker],
            val_selection, dates, val_start, val_end, best_buy, best_exit,
            top_n=top_n, allocation_freq=allocation_freq, cash_policy=cash_policy,
            use_regime=use_regime, use_mask=use_mask
        )

        st.write(f"🏆 Champion chiến thắng trên Validation: **{champion}**")

        st.write("4️⃣ Đang chạy Out-of-sample Final Test độc lập...")
        test_res, test_curves, test_orders = evaluate_final_test_pipeline(
            price_dict, feature_store, benchmark_ticker, feature_store[benchmark_ticker],
            test_selection, dates, test_start, test_end, best_buy, best_exit,
            champion=champion, top_n=top_n, allocation_freq=allocation_freq,
            cash_policy=cash_policy, use_regime=use_regime, use_mask=use_mask
        )

        # Lưu vào session state
        st.session_state["backtest_results"] = {
            "val_res": val_res,
            "champion": champion,
            "val_curves": val_curves,
            "val_orders": val_orders,
            "test_res": test_res,
            "test_curves": test_curves,
            "test_orders": test_orders,
            "test_selection": test_selection,
            "val_selection": val_selection,
            "threshold_trials": threshold_trials,
            "best_buy": best_buy,
            "best_exit": best_exit,
            "split_info": {
                "core_start": dates[0],
                "core_end": core_end,
                "val_start": val_start,
                "val_end": val_end,
                "test_start": test_start,
                "test_end": test_end,
                "benchmark": benchmark_ticker
            }
        }
        status.update(label="✅ Hoàn thành kiểm định chiến lược thành công!", state="complete", expanded=False)

# Lấy kết quả từ session_state
res_pack = st.session_state["backtest_results"]
test_res = res_pack["test_res"]
test_curves = res_pack["test_curves"]
test_orders = res_pack["test_orders"]
val_res = res_pack["val_res"]
champion = res_pack["champion"]
test_selection = res_pack["test_selection"]
best_buy = res_pack["best_buy"]
best_exit = res_pack["best_exit"]
split_info = res_pack["split_info"]


# ==========================================================
# 13. UI DISPLAY - TABS FOR METRICS, CHARTS, & DIAGNOSTICS
# ==========================================================
st.title("📊 Kết Quả Kiểm Định Chiến Lược Định Lượng HOSE (V4)")
st.caption(f"Universe: {data['Ticker'].nunique()} mã | Benchmark: {split_info['benchmark']} | Thời gian Final Test: {split_info['test_start'].strftime('%d/%m/%Y')} -> {split_info['test_end'].strftime('%d/%m/%Y')}")

# KPI Scorecards hàng đầu
col1, col2, col3, col4, col5 = st.columns(5)

champ_metrics = test_res.loc["LOCKED_CHAMPION"] if "LOCKED_CHAMPION" in test_res.index else test_res.iloc[0]
bench_metrics = test_res.loc["VNINDEX_or_MarketBenchmark"] if "VNINDEX_or_MarketBenchmark" in test_res.index else None
bh_metrics = test_res.loc["STATIC_EW_BH"] if "STATIC_EW_BH" in test_res.index else None

with col1:
    cagr = champ_metrics.get("CAGR [%]", 0)
    bcagr = bench_metrics.get("CAGR [%]", 0) if bench_metrics is not None else 0
    st.metric("CAGR (Lợi nhuận năm)", f"{cagr:.2f}%", f"{cagr - bcagr:+.2f}% vs VN-Index")

with col2:
    sh = champ_metrics.get("Sharpe", 0)
    bsh = bench_metrics.get("Sharpe", 0) if bench_metrics is not None else 0
    st.metric("Sharpe Ratio", f"{sh:.2f}", f"{sh - bsh:+.2f} vs VN-Index")

with col3:
    mdd = champ_metrics.get("MaxDD [%]", 0)
    bmdd = bench_metrics.get("MaxDD [%]", 0) if bench_metrics is not None else 0
    st.metric("Max Drawdown", f"{mdd:.2f}%", f"{mdd - bmdd:+.2f}% vs VN-Index", delta_color="inverse")

with col4:
    calmar = champ_metrics.get("Calmar", 0)
    st.metric("Calmar Ratio", f"{calmar:.2f}" if pd.notna(calmar) else "N/A")

with col5:
    cash_mean = champ_metrics.get("AvgCash [%]", 0)
    st.metric("Tiền mặt TB", f"{cash_mean:.1f}%", f"Cổ phiếu: {100-cash_mean:.1f}%")

st.markdown("---")

# Các Tab chuyên sâu
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📈 Final Test (OOS)",
    "🏆 Validation & Champion",
    "🔄 Top N & Phân Bổ",
    "🛡️ Market Regime & Tiền Mặt",
    "📑 Nhật Ký Lệnh & Audit",
    "🔍 Khám Phá Dữ Liệu & TA"
])

# ----------------------------------------------------------
# TAB 1: FINAL TEST (OUT-OF-SAMPLE)
# ----------------------------------------------------------
with tab1:
    st.subheader("Hiệu Suất Tăng Trưởng Tài Sản (Out-of-Sample Final Test)")

    # Lựa chọn hiển thị mô hình trên biểu đồ
    all_models = list(test_curves.keys())
    default_shown = [m for m in ["LOCKED_CHAMPION", "VNINDEX_or_MarketBenchmark", "STATIC_EW_BH", "DYN_EW_NO_REGIME"] if m in all_models]
    selected_plot_models = st.multiselect("Chọn các mô hình muốn so sánh:", options=all_models, default=default_shown)

    # Plotly Normalized NAV Chart
    fig_nav = go.Figure()

    color_palette = {
        "LOCKED_CHAMPION": "#10B981",          # Emerald Green
        "VNINDEX_or_MarketBenchmark": "#1E293B", # Dark Slate
        "STATIC_EW_BH": "#3B82F6",             # Blue
        "DYN_EW_NO_REGIME": "#F59E0B",         # Amber
    }

    for m in selected_plot_models:
        s = pd.Series(test_curves[m]).dropna()
        if len(s) == 0:
            continue
        norm_nav = s / s.iloc[0]
        line_color = color_palette.get(m, None)
        line_width = 3.0 if m == "LOCKED_CHAMPION" else (2.0 if m in ["VNINDEX_or_MarketBenchmark", "STATIC_EW_BH"] else 1.5)
        dash = "dash" if m == "STATIC_EW_BH" else ("dot" if "DIAG_" in m else "solid")

        label = f"🏆 {m} (Champion)" if m == "LOCKED_CHAMPION" else (
            "📊 VN-Index (Benchmark)" if m == "VNINDEX_or_MarketBenchmark" else (
            "📌 Static EW Buy & Hold" if m == "STATIC_EW_BH" else m
        ))

        fig_nav.add_trace(go.Scatter(
            x=norm_nav.index,
            y=norm_nav.values,
            mode='lines',
            name=label,
            line=dict(color=line_color, width=line_width, dash=dash),
            hovertemplate="<b>%{x|%d/%m/%Y}</b><br>NAV Tương đối: %{y:.3f}<extra></extra>"
        ))

    fig_nav.add_hline(y=1.0, line_dash="dot", line_color="gray", line_width=1)
    fig_nav.update_layout(
        title="Đường Tăng Trưởng Tài Sản Chuẩn Hóa (Normalized Equity Curve - Final Test)",
        xaxis_title="Thời gian",
        yaxis_title="NAV Chuẩn Hóa (Gốc = 1.0)",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=40, r=40, t=60, b=40),
        height=520
    )
    st.plotly_chart(fig_nav, use_container_width=True)

    # Underwater Drawdown Chart
    st.subheader("Mức Sụt Giảm Tối Đa (Underwater Drawdown)")
    fig_dd = go.Figure()
    for m in selected_plot_models:
        s = pd.Series(test_curves[m]).dropna()
        if len(s) == 0:
            continue
        dd = (s / s.cummax() - 1) * 100
        line_color = color_palette.get(m, None)
        fig_dd.add_trace(go.Scatter(
            x=dd.index,
            y=dd.values,
            mode='lines',
            name=m,
            fill='tozeroy' if m == "LOCKED_CHAMPION" else None,
            line=dict(color=line_color, width=1.5),
            hovertemplate="<b>%{x|%d/%m/%Y}</b><br>Drawdown: %{y:.2f}%<extra></extra>"
        ))
    fig_dd.update_layout(
        yaxis_title="Sụt giảm từ đỉnh (%)",
        xaxis_title="Thời gian",
        hovermode="x unified",
        height=320,
        margin=dict(l=40, r=40, t=20, b=40)
    )
    st.plotly_chart(fig_dd, use_container_width=True)

    # Bảng số liệu chi tiết
    st.subheader("Bảng Tổng Hợp Chỉ Số Hiệu Suất Toàn Diện (Final Test)")
    cols_display = [
        "Return [%]", "CAGR [%]", "Sharpe", "Sortino", "MaxDD [%]", "Calmar",
        "Volatility [%]", "Turnover [%]", "AvgExposure [%]", "AvgCash [%]",
        "ExcessCAGR_vs_Benchmark [pp]", "DD_Improvement_vs_Benchmark [pp]"
    ]
    avail_cols = [c for c in cols_display if c in test_res.columns]
    styled_df = test_res[avail_cols].style.format("{:.2f}")
    st.dataframe(styled_df, use_container_width=True)


# ----------------------------------------------------------
# TAB 2: VALIDATION & CHAMPION SELECTION
# ----------------------------------------------------------
with tab2:
    st.subheader("Giai Đoạn Validation: Đấu Loại Mô Hình & Khóa Champion")
    st.markdown(f"""
    > **Mô hình vô địch (Locked Champion):** <span class="badge-champion">{champion}</span>  
    > Trong giai đoạn Validation ({split_info['val_start'].strftime('%d/%m/%Y')} -> {split_info['val_end'].strftime('%d/%m/%Y')}), tất cả các trường phái phân bổ tỷ trọng (Equal Weight, Markowitz MPT-LW, Min Variance, HRP, Blend50) đều nhận **cùng một danh sách Top N cổ phiếu** để đảm bảo tính khách quan tuyệt đối.
    """, unsafe_allow_html=True)

    col_val1, col_val2 = st.columns([3, 2])
    with col_val1:
        st.write("##### Bảng xếp hạng Validation Score:")
        cols_val = ["Return [%]", "Sharpe", "Sortino", "MaxDD [%]", "Turnover [%]", "ValidationScore"]
        v_show = val_res[[c for c in cols_val if c in val_res.columns]].sort_values("ValidationScore", ascending=False)
        st.dataframe(v_show.style.format("{:.3f}").highlight_max(subset=["ValidationScore"], color="#d1fae5"), use_container_width=True)

    with col_val2:
        st.write("##### Công thức tính điểm Validation:")
        st.latex(r"""
        \text{Score} = 0.45 \cdot \text{Sharpe} + 0.15 \cdot \text{Sortino} + 0.012 \cdot (\text{Ret} - \text{Ret}_{bm}) + 0.004 \cdot \text{MaxDD} - 0.0005 \cdot \text{Turnover}
        """)
        st.caption("Công thức cân đối giữa tỷ suất sinh lời vượt trội, độ biến động âm, mức sụt giảm tối đa và chi phí đảo danh mục.")

    # Tối ưu ngưỡng BUY/EXIT tại Core Train
    st.subheader("Ngưỡng BUY / EXIT Tối Ưu Đã Khóa Trên Core Train")
    col_t1, col_t2 = st.columns([1, 2])
    with col_t1:
        st.info(f"""
        **Ngưỡng BUY Đã Khóa:** `{best_buy}`  
        **Ngưỡng EXIT Đã Khóa:** `{best_exit}`  
        *Khóa tĩnh từ Core Train ({split_info['core_start'].strftime('%d/%m/%Y')} -> {split_info['core_end'].strftime('%d/%m/%Y')}), không can thiệp vào Test.*
        """)
    with col_t2:
        trials_df = res_pack.get("threshold_trials", pd.DataFrame())
        if not trials_df.empty:
            st.dataframe(trials_df.head(6).style.format({
                "GlobalCVScore": "{:.3f}", "MedianSharpe": "{:.2f}",
                "MedianReturnPct": "{:.2f}%", "MedianMaxDDPct": "{:.2f}%"
            }), use_container_width=True)


# ----------------------------------------------------------
# TAB 3: TOP N SELECTION & ALLOCATION SCHEDULES
# ----------------------------------------------------------
with tab3:
    st.subheader("Lịch Trình Tái Cơ Cấu Cổ Phiếu (Walk-Forward Reselection)")
    st.caption("Top N được tái chọn lọc định kỳ theo từng quý. Dữ liệu chỉ lấy đến phiên giao dịch liền trước (As-Of Date t-1), triệt tiêu hoàn toàn Lookahead Bias.")

    st.write("##### Lịch chọn Top N trong giai đoạn Final Test:")
    st.dataframe(test_selection, use_container_width=True)

    # Thống kê tần suất xuất hiện các mã trong Top N
    st.write("##### Tần suất các mã lọt vào Top N danh mục:")
    all_top_stocks = []
    for c in [f"Top{i+1}" for i in range(top_n)]:
        if c in test_selection.columns:
            all_top_stocks.extend(test_selection[c].dropna().tolist())

    stock_counts = pd.Series(all_top_stocks).value_counts().reset_index()
    stock_counts.columns = ["Mã Cổ Phiếu", "Số Kỳ Xuất Hiện"]

    fig_bar = px.bar(
        stock_counts,
        x="Mã Cổ Phiếu",
        y="Số Kỳ Xuất Hiện",
        color="Số Kỳ Xuất Hiện",
        color_continuous_scale="Viridis",
        title=f"Số lần các mã được lựa chọn vào Top {top_n} trong Final Test"
    )
    fig_bar.update_layout(height=350, margin=dict(l=40, r=40, t=40, b=40))
    st.plotly_chart(fig_bar, use_container_width=True)


# ----------------------------------------------------------
# TAB 4: MARKET REGIME & CASH POLICY EXPERIMENT
# ----------------------------------------------------------
with tab4:
    st.subheader("Bộ Lọc Xu Hướng Thị Trường & Quản Trị Tiền Mặt")
    st.markdown("""
    - **Market Regime (EMA 200 + Slope Confirmation):** Khi VN-Index nằm dưới đường EMA 200 hoặc đường EMA 200 có xu hướng dốc xuống, hệ thống tự động nhận diện pha Downtrend (Risk-Off) và phòng vệ chuyển tỷ trọng sang tiền mặt.
    - **Thí nghiệm Cash Policy:**
        - `cash40`: Giải ngân tối thiểu 60% vào cổ phiếu mục tiêu, tỷ trọng tiền mặt tối đa khoảng 40%.
        - `original`: Phòng thủ tối đa (có thể cầm 100% tiền mặt khi thị trường xác nhận Risk-Off).
        - `full`: Luôn giải ngân 100% vào cổ phiếu.
    """)

    # Biểu đồ VNINDEX cùng EMA200 và vùng Risk-On / Risk-Off
    bench_feat = feature_store[split_info['benchmark']].loc[split_info['test_start']:split_info['test_end']]
    regime_df = build_market_regime(bench_feat)

    fig_regime = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.08,
                               row_heights=[0.7, 0.3], subplot_titles=("VN-Index & Đường Trung Bình EMA200", "Trạng Thái Phòng Vệ (Risk-On = 1 / Risk-Off = 0)"))

    fig_regime.add_trace(go.Scatter(x=bench_feat.index, y=bench_feat["Close"], name="VNINDEX Close", line=dict(color="#1e293b", width=2)), row=1, col=1)
    fig_regime.add_trace(go.Scatter(x=regime_df.index, y=regime_df["EMA_Market"], name="EMA 200", line=dict(color="#dc2626", width=1.5, dash="dash")), row=1, col=1)

    fig_regime.add_trace(go.Scatter(x=regime_df.index, y=regime_df["RiskOn_Open"].astype(int), name="Risk-On (Được mở vị thế)", line=dict(color="#10b981", width=1.5), fill="tozeroy"), row=2, col=1)

    fig_regime.update_layout(height=480, margin=dict(l=40, r=40, t=40, b=40), hovermode="x unified")
    st.plotly_chart(fig_regime, use_container_width=True)


# ----------------------------------------------------------
# TAB 5: TRADE ORDERS LOG & LEAKAGE AUDIT
# ----------------------------------------------------------
with tab5:
    st.subheader("Nhật Ký Khớp Lệnh Chi Tiết (Event-Driven Execution)")
    st.caption("Lệnh giao dịch chỉ thực thi tại giá Open của phiên tiếp theo sau khi tín hiệu Close được xác nhận.")

    active_order_model = st.selectbox("Chọn mô hình để xem nhật ký giao dịch:", options=list(test_orders.keys()), index=0)
    orders_df = test_orders.get(active_order_model, pd.DataFrame())

    if not orders_df.empty:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            side_filter = st.multiselect("Lọc hành động:", options=["BUY", "SELL"], default=["BUY", "SELL"])
        with col_f2:
            ticker_list = ["Tất cả"] + sorted(orders_df["Ticker"].unique().tolist())
            selected_t = st.selectbox("Lọc theo mã cổ phiếu:", options=ticker_list, index=0)

        filtered_orders = orders_df[orders_df["Side"].isin(side_filter)]
        if selected_t != "Tất cả":
            filtered_orders = filtered_orders[filtered_orders["Ticker"] == selected_t]

        st.dataframe(filtered_orders.sort_values("Date", ascending=False).style.format({
            "ExecPrice": "{:,.1f}", "Shares": "{:,.0f}", "Fee": "{:,.0f}"
        }), use_container_width=True)
    else:
        st.info("Không có lệnh giao dịch nào được ghi nhận.")

    # Leakage Audit Table
    st.subheader("Bảng Kiểm Định Phòng Ngừa Rò Rỉ Dữ Liệu (Leakage Audit)")
    leakage_table = pd.DataFrame([
        ["1. Ngưỡng TA BUY/EXIT", split_info['core_end'].strftime('%Y-%m-%d'), "Không", "Chỉ tối ưu trên Core Train"],
        ["2. Lựa chọn Mô hình Champion", split_info['val_end'].strftime('%Y-%m-%d'), "Không", "Chỉ chọn trên Validation"],
        ["3. Top N cổ phiếu từng quý", "Phiên t-1", "Không", "Mỗi kỳ chỉ dùng dữ liệu <= phiên t-1"],
        ["4. Trọng số MPT/HRP từng tháng", "Phiên t-1", "Không", "Ma trận hiệp phương sai tính trên prices <= t-1"],
        ["5. Tín hiệu Market Regime & Mask", "Phiên t-1 Close -> Open t", "Không", "Tín hiệu đóng cửa hôm trước thực thi mở cửa hôm sau"],
        ["6. Đánh giá Final Test", split_info['test_end'].strftime('%Y-%m-%d'), "Chỉ Đánh Giá", "Tuyệt đối không retune tham số"],
    ], columns=["Hạng Mục", "Mốc Dữ Liệu Tối Đa", "Dùng Dữ Liệu Final Test?", "Quy Tắc Đảm Bảo"])

    st.table(leakage_table)

    # Xuất báo cáo Excel
    st.subheader("📥 Xuất Toàn Bộ Báo Cáo Ra Excel")
    def export_excel_report():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            test_res.to_excel(writer, sheet_name="FinalTest_Metrics")
            val_res.to_excel(writer, sheet_name="Validation_Metrics")
            test_selection.to_excel(writer, sheet_name="Test_TopN_Schedule", index=False)
            leakage_table.to_excel(writer, sheet_name="Leakage_Audit", index=False)
            if not orders_df.empty:
                orders_df.to_excel(writer, sheet_name="Champion_Orders", index=False)
        return output.getvalue()

    excel_data = export_excel_report()
    st.download_button(
        label="📥 Tải xuống báo cáo kết quả Excel (.xlsx)",
        data=excel_data,
        file_name=f"HOSE_Quant_Backtest_Results_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )


# ----------------------------------------------------------
# TAB 6: DATA EXPLORER & TECHNICAL ANALYSIS SCREENER
# ----------------------------------------------------------
with tab6:
    st.subheader("Khám Phá Dữ Liệu & Chỉ Báo Kỹ Thuật Chi Tiết Từng Cổ Phiếu")

    stock_select = st.selectbox(
        "Chọn mã cổ phiếu cần soi tín hiệu kỹ thuật:",
        options=sorted([t for t in price_dict.keys() if t != split_info['benchmark']]),
        index=0
    )

    stock_feat = feature_store[stock_select].tail(252)

    # Biểu đồ giá + EMAs + Volume + TA Score
    fig_stock = make_subplots(
        rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.06,
        row_heights=[0.55, 0.25, 0.20],
        subplot_titles=(f"Diễn biến giá & Đường EMA ({stock_select})", "Chỉ số RSI & MACD", "Điểm Đánh Giá Kỹ Thuật (TA Score 0-10)")
    )

    fig_stock.add_trace(go.Scatter(x=stock_feat.index, y=stock_feat["Close"], name="Giá Đóng Cửa", line=dict(color="#1e293b", width=2)), row=1, col=1)
    fig_stock.add_trace(go.Scatter(x=stock_feat.index, y=stock_feat["EMA20"], name="EMA 20", line=dict(color="#3b82f6", width=1)), row=1, col=1)
    fig_stock.add_trace(go.Scatter(x=stock_feat.index, y=stock_feat["EMA50"], name="EMA 50", line=dict(color="#f59e0b", width=1)), row=1, col=1)
    fig_stock.add_trace(go.Scatter(x=stock_feat.index, y=stock_feat["EMA200"], name="EMA 200", line=dict(color="#ef4444", width=1.5)), row=1, col=1)

    fig_stock.add_trace(go.Scatter(x=stock_feat.index, y=stock_feat["RSI"], name="RSI (14)", line=dict(color="#8b5cf6", width=1.5)), row=2, col=1)
    fig_stock.add_hline(y=70, line_dash="dash", line_color="red", line_width=1, row=2, col=1)
    fig_stock.add_hline(y=30, line_dash="dash", line_color="green", line_width=1, row=2, col=1)

    fig_stock.add_trace(go.Bar(x=stock_feat.index, y=stock_feat["TA_Score"], name="TA Score (0-10)", marker_color="#10b981"), row=3, col=1)
    fig_stock.add_hline(y=best_buy, line_dash="dot", line_color="blue", annotation_text=f"BUY Threshold ({best_buy})", row=3, col=1)
    fig_stock.add_hline(y=best_exit, line_dash="dot", line_color="orange", annotation_text=f"EXIT Threshold ({best_exit})", row=3, col=1)

    fig_stock.update_layout(height=650, margin=dict(l=40, r=40, t=40, b=40), hovermode="x unified")
    st.plotly_chart(fig_stock, use_container_width=True)

    # Hiển thị bảng chất lượng dữ liệu
    st.write("##### Bảng thống kê chất lượng dữ liệu HOSE:")
    q_table = data_quality_table(data)
    st.dataframe(q_table, use_container_width=True)

