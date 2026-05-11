#!/usr/bin/env python
import requests
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Alignment
from openpyxl.utils import get_column_letter
import warnings
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
warnings.simplefilter("ignore", FutureWarning)

BASE = "https://financialmodelingprep.com/api/v3"
API_KEY = "Your_API_KEY"   ##enter your purchased API key
OUTPUT_DIR = Path.home() / "Desktop" / "FullHistory"
CREATE_EXCEL = True

ENDPOINTS = [
    ("Annual_IncomeSheet", "income-statement", {}),
    ("Annual_BalanceSheet", "balance-sheet-statement", {}),
    ("Annual_CashFlow", "cash-flow-statement", {}),
    ("Quarterly_IncomeSheet", "income-statement", {"period": "quarter"}),
    ("Quarterly_BalanceSheet", "balance-sheet-statement", {"period": "quarter"}),
    ("Quarterly_Cashflow", "cash-flow-statement", {"period": "quarter"}),
]


#Company Information
def fetch_company_profile(session, ticker, api_key):
    url = f"{BASE}/profile/{ticker}"
    params = {"apikey": api_key}
    r = session.get(url, params=params, timeout=60)
    r.raise_for_status()
    data = r.json()
    if isinstance(data, dict) and data.get("Error Message"):
        raise ValueError(data["Error Message"])
    if not isinstance(data, list) or not data:
        raise ValueError(f"No company profile found for {ticker}")
    return data[0]

#UI front-end progress bar
def show_progress_window(total_tickers):
    win = tk.Tk()
    win.title("Download Progress")
    win.geometry("400x120")
    win.resizable(False, False)
    label = tk.Label(win, text="Downloading tickers...", font=("Arial", 11))
    label.pack(pady=10)
    bar = ttk.Progressbar(win, orient="horizontal", length=300, mode="determinate")
    bar.pack(pady=5)
    bar["maximum"] = total_tickers
    bar["value"] = 0
    count_label = tk.Label(win, text="0 / {}".format(total_tickers), font=("Arial", 10))
    count_label.pack(pady=5)
    win.update()
    return win, bar, count_label

#UI front-end Front-view textarea to download
def get_tickers_from_dialog():
    result = {"value": None}
    root = tk.Tk()
    root.title("TICKER DOWNLOAD")
    root.geometry("800x500")
    root.resizable(False, False)
    label = tk.Label(
        root,
        text="Enter tickers separated by comma or newline",
        font=("Arial", 10)
    )
    label.pack(pady=10)
    text_box = tk.Text(root, font=("Arial", 10), height=20, width=80)
    text_box.pack(padx=20, pady=10)
    def submit():
        result["value"] = text_box.get("1.0", "end").strip()
        root.destroy()
    def on_close():
        # User clicked X: treat as cancel
        result["value"] = None
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_close)    
    submit_btn = tk.Button(root, text="DOWNLOAD", command=submit, font=("Arial", 10), width=12)
    submit_btn.pack(pady=10)
    root.bind("<Control-Return>", lambda event: submit())
    root.mainloop()
    raw = result["value"]
    if raw is None:
        return None
    raw = raw.replace("\n", ",")
    tickers = [t.strip() for t in raw.split(",") if t.strip()]
    if not tickers:
        # Empty or just commas/spaces: treat as cancel
        return None
    return tickers


# Ticker - Japan and other geographies
def normalize_ticker(raw: str):
    """
    raw: what the user typed (e.g. '1833' or '1833.T').
    returns: (api_ticker, file_stem)
    """
    t = raw.strip().upper()
    # If user already provided suffix, respect it
    if "." in t:
        base = t.split(".")[0]
        return t, base
    # If it's all digits, treat as Japanese .T ticker
    if t.isdigit():
        return t + ".T", t  # API 1833.T, file 1833
    # Default: use as-is
    return t, t

#Download Annual and Quarterly Statements and Format excel sheet
def fetch_statement(session, endpoint, ticker, api_key, extra_params=None, limit=400):
    params = {"apikey": api_key, "limit": limit}
    if extra_params:
        params.update(extra_params)
    url = f"{BASE}/{endpoint}/{ticker}"
    r = session.get(url, params=params, timeout=60)
    r.raise_for_status()
    data = r.json()
    if isinstance(data, dict) and data.get("Error Message"):
        raise ValueError(data["Error Message"])
    if not isinstance(data, list):
        raise ValueError(f"Unexpected response for {endpoint}: {data}")
    return pd.DataFrame(data)

def sanitize_sheet_name(name):
    return name[:31]

def format_sheet(ws):
    normal_font = Font(bold=False)
    no_border = Border()
    bottom_left = Alignment(horizontal="left", vertical="bottom")
    max_row = ws.max_row
    max_col = ws.max_column
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = normal_font
            cell.border = no_border
            cell.alignment = bottom_left

    for c in range(1, max_col + 1):
        col_letter = get_column_letter(c)
        ws.column_dimensions[col_letter].width = 20

#Save the in .xlsx Desktop under FullHistory Folder
def build_file_for_ticker(raw_ticker, outdir):
    api_ticker, file_stem = normalize_ticker(raw_ticker)
    session = requests.Session()
    frames = {}
    for sheet_name, endpoint, extra in ENDPOINTS:
        limit = 250 if extra.get("period") == "quarter" else 120
        df = fetch_statement(session, endpoint, api_ticker, API_KEY, extra_params=extra, limit=limit)
        frames[sheet_name] = df
    profile = fetch_company_profile(session, api_ticker, API_KEY)
    profile_df = pd.DataFrame(list(profile.items()), columns=["Field", "Value"])
    xlsx_path = outdir / f"{file_stem}.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        # financial statements (transposed)
        for sheet_name, df in frames.items():
            df.T.to_excel(
                writer,
                sheet_name=sanitize_sheet_name(sheet_name),
                header=False,
                index=True,
                startrow=1
            )
        # company info sheet
        profile_df.to_excel(
            writer,
            sheet_name="Company_Info",
            header=False,
            index=False
        )
    wb = load_workbook(xlsx_path)
    for ws in wb.worksheets:
        format_sheet(ws)
    wb.save(xlsx_path)

#Main code
def main():
    if API_KEY == "PUT_YOUR_NEW_FMP_KEY_HERE":
        raise ValueError("Please paste your new FMP API key into API_KEY first.")
    outdir = OUTPUT_DIR
    outdir.mkdir(parents=True, exist_ok=True)
    tickers = get_tickers_from_dialog()
    if not tickers:
        return
    total = len(tickers)
    prog_win, prog_bar, prog_label = show_progress_window(total)
    errors = []
    done_count = 0
    for ticker in tickers:
        try:
            build_file_for_ticker(ticker, outdir)
        except Exception as e:
            print(f"Error for {ticker}: {e}")
            errors.append((ticker, str(e)))
        done_count += 1
        prog_bar["value"] = done_count
        prog_label.config(text=f"{done_count} / {total}")
        prog_win.update_idletasks()
    prog_win.destroy()
    root = tk.Tk()
    root.withdraw()
    if errors:
        msg = "Completed with errors for some tickers:\n"
        msg += "\n".join(f"{t}: {err}" for t, err in errors)
        messagebox.showwarning("Done with errors", msg)
    else:
        messagebox.showinfo("Done", f"All files saved in:\n{outdir.resolve()}")
    root.destroy()

if __name__ == "__main__":
    main()
