# CodingForFinance

# ======================================================
# Yahoo Finance -> Excel (TTE.PA)
# Annual FS (thousands) + Daily History since 2015-09-01
# (fix: remove timezone from history index before Excel)
# ======================================================

# 0) Install
!pip -q install "yfinance>=0.2.40" openpyxl

import pandas as pd
import yfinance as yf
from datetime import datetime, timezone

# 1) Colab download helper
IN_COLAB = False
try:
    from google.colab import files  # type: ignore
    IN_COLAB = True
except Exception:
    pass

# 2) Parameters
TICKER = "TTE.PA"
HIST_START = "2015-09-01"
OUTPUT_FILENAME = f"{TICKER}_YahooFinance_Annual_Thousands_DailyHistory.xlsx"

# 3) Fetch raw data (no value fabrication)
tkr = yf.Ticker(TICKER)

def fetch_financials(tkr, which, freq):
    """Try yfinance new API first, fallback to legacy attributes."""
    try:
        if which == "income":
            return tkr.get_income_stmt(freq=freq)
        if which == "balance":
            return tkr.get_balance_sheet(freq=freq)
        if which == "cashflow":
            return tkr.get_cash_flow(freq=freq)
    except Exception:
        pass
    try:
        if which == "income" and freq == "annual":
            return getattr(tkr, "income_stmt", pd.DataFrame())
        if which == "balance" and freq == "annual":
            return getattr(tkr, "balance_sheet", pd.DataFrame())
        if which == "cashflow" and freq == "annual":
            return getattr(tkr, "cashflow", pd.DataFrame())
    except Exception:
        pass
    return pd.DataFrame()

# 4) Reorder + scale to thousands
def to_yahoo_display_thousands(df: pd.DataFrame) -> pd.DataFrame:
    """
    - Colonnes: plus récent -> plus ancien (gauche -> droite)
    - Lignes: inversées (bas -> haut)
    - Valeurs numériques divisées par 1000 ("All numbers in thousands")
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df.copy()

    cols = list(df.columns)
    try:
        dt = pd.to_datetime(cols)
        cols = [c for _, c in sorted(zip(dt, cols), reverse=True)]
    except Exception:
        pass

    out = df.iloc[::-1].copy()
    out = out[cols]

    num_cols = out.select_dtypes(include=["number"]).columns
    out[num_cols] = out[num_cols] / 1000.0
    return out

# 5) Fetch annual financial statements ONLY
print("⏳ Récupération états financiers annuels (Income, Balance, CashFlow)...")
income_annual_raw   = fetch_financials(tkr, "income",   "annual")
balance_annual_raw  = fetch_financials(tkr, "balance",  "annual")
cashflow_annual_raw = fetch_financials(tkr, "cashflow", "annual")

def log_df(df, name):
    print(f"{name}: {'OK '+str(df.shape) if isinstance(df, pd.DataFrame) and not df.empty else 'N/A'}")

log_df(income_annual_raw,  "Income Annual (raw)")
log_df(balance_annual_raw, "Balance Annual (raw)")
log_df(cashflow_annual_raw,"CashFlow Annual (raw)")

if all(isinstance(df, pd.DataFrame) and df.empty for df in [
    income_annual_raw, balance_annual_raw, cashflow_annual_raw
]):
    raise ValueError("Aucun état financier annuel disponible actuellement pour TTE.PA sur Yahoo.")

# Apply Yahoo layout + thousands
income_annual_y   = to_yahoo_display_thousands(income_annual_raw)
balance_annual_y  = to_yahoo_display_thousands(balance_annual_raw)
cashflow_annual_y = to_yahoo_display_thousands(cashflow_annual_raw)

# 6) Fetch daily historical data since 2015-09-01 (and DROP TIMEZONE)
print("⏳ Récupération historique journalier depuis 2015-09-01...")
hist = tkr.history(start=HIST_START, interval="1d", actions=True, auto_adjust=False)

if isinstance(hist, pd.DataFrame) and not hist.empty:
    hist_out = hist.copy()
    # --- FIX: Excel n'accepte pas les datetimes avec timezone ---
    if getattr(hist_out.index, "tz", None) is not None:
        try:
            hist_out.index = hist_out.index.tz_convert(None)
        except Exception:
            hist_out.index = hist_out.index.tz_localize(None)
    # Si des colonnes datetime tz-aware existaient (rare), on les rend naïves aussi
    for c in hist_out.select_dtypes(include=["datetimetz"]).columns:
        try:
            hist_out[c] = hist_out[c].dt.tz_convert(None)
        except Exception:
            hist_out[c] = hist_out[c].dt.tz_localize(None)
    hist_out.index.name = "Date"
else:
    hist_out = pd.DataFrame()
log_df(hist_out, "Daily History (tz-naive)")

# 7) Write Excel (annual only + history)
print("💾 Écriture du fichier Excel ...")
with pd.ExcelWriter(OUTPUT_FILENAME, engine="openpyxl") as writer:
    if not income_annual_y.empty:
        income_annual_y.to_excel(writer, sheet_name="Income_Annual")
    if not balance_annual_y.empty:
        balance_annual_y.to_excel(writer, sheet_name="Balance_Annual")
    if not cashflow_annual_y.empty:
        cashflow_annual_y.to_excel(writer, sheet_name="CashFlow_Annual")

    if not hist_out.empty:
        hist_out.to_excel(writer, sheet_name="Daily_History")

    meta = pd.DataFrame({
        "Field": ["Ticker", "Source", "Library", "Fetched_UTC",
                  "FS_Row_Order", "FS_Column_Order", "FS_Value_Scale",
                  "History_Start", "History_Interval", "History_Adjusted",
                  "History_Timezone_Stripped"],
        "Value": [
            TICKER,
            "Yahoo Finance via yfinance",
            yf.__version__,
            datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S %Z"),
            "Bas → Haut (index inversé)",
            "Plus récent → Plus ancien (gauche → droite)",
            "All numbers in thousands (÷1000) — états financiers uniquement",
            HIST_START,
            "1d",
            "auto_adjust=False (Close & Adj Close disponibles)",
            "Yes (index datetime sans timezone)"
        ]
    })
    meta.to_excel(writer, sheet_name="Info_Source", index=False)

print(f"✅ Fichier Excel généré : {OUTPUT_FILENAME}")

# 8) Auto-download in Colab
if IN_COLAB:
    files.download(OUTPUT_FILENAME)
else:
    print("ℹ️ Exécution locale : le fichier est dans le répertoire courant.")
