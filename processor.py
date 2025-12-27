import os
import re
from typing import Optional
from openpyxl import load_workbook
import pandas as pd

DATE_RE = re.compile(r"(\d{4})(\d{2})(\d{2})_NSE\.csv$", re.IGNORECASE)

PORTFOLIO_FILE = "portfolio_tracker.xlsx"

def parse_date_from_filename(filename: str) -> Optional[pd.Timestamp]:
    match = DATE_RE.search(filename)
    if not match:
        return None
    year, month, day = match.groups()
    try:
        return pd.to_datetime(f"{year}-{month}-{day}", format="%Y-%m-%d")
    except Exception:
        return None
    
def read_daily_close(path: str) -> pd.Series:
    df = pd.read_csv(
        path, 
        usecols=["SYMBOL", "SERIES", "CLOSE"], 
        dtype={"SYMBOL": str, "SERIES": str, "CLOSE": float}
    )
    
    df = df[df["SERIES"].isin(["EQ", "BE"])]
    df = df.dropna(subset=["SYMBOL", "CLOSE"])
    df = df.drop_duplicates(subset=["SYMBOL"], keep="last")

    return df.set_index("SYMBOL")["CLOSE"]

def process_csv_files(
    folder_path: str,
    start_date: pd.Timestamp,
    lookback_months: int = 12,
    top_k: int = 10,
    buffer: int = 10,
    benchmark: float = 6.0,
    portfolio_capital: float = 1_00_000,
    progress_signal=None
) -> str:
    
    #Collect relevant files
    files = []
    for f in os.listdir(folder_path):
        date = parse_date_from_filename(f)
        if date is not None and date <= start_date:
            files.append((date, f))
    
    if not files:
        raise ValueError("No CSV files found up to the selected Date.")
    
    files.sort(key=lambda x: x[0])
    
    #Lookback Window
    lookback_start = start_date - pd.DateOffset(months=lookback_months)
    
    files = [(date, fname) for date, fname in files if date >= lookback_start]
    
    if len(files) < 2:
        raise ValueError("Not enough data in the lookback window.")
    
    #Load Prices

    price_rows = []
    for i, (date, fname) in enumerate(files):
        path = os.path.join(folder_path, fname)
        closes = read_daily_close(path)
        row = closes.to_frame().T
        row.index = [date]
        price_rows.append(row)

        if progress_signal:
            progress_signal.emit(int((i+1) / len(files) * 40))
        
    prices = pd.concat(price_rows).sort_index()
        
    first_prices = prices.iloc[0]
    last_prices = prices.iloc[-1]
    
    momentum = (last_prices / first_prices) - 1.0
    momentum = momentum.dropna()

    weight = 1.0 / top_k
    max_price = weight * portfolio_capital

    affordable  = last_prices[last_prices <= max_price].index
    momentum = momentum.loc[momentum.index.intersection(affordable)]

    ranked = momentum.sort_values(ascending=False)
    ranked = ranked.head(top_k + buffer)
    #Load Portfolio
    benchmark_return = benchmark / 100.0
    benchmark_return = ((1 + benchmark_return) ** (lookback_months / 12)) - 1

    to_sell = []
    to_buy = []


    portfolio_path = os.path.join(folder_path, PORTFOLIO_FILE)
    if os.path.exists(portfolio_path):
        prev_portfolio = pd.read_excel(portfolio_path, sheet_name="Holdings")

        if "Symbol" not in prev_portfolio.columns:
            raise ValueError("Previous portfolio file missing 'Symbol' column")

        prev_portfolio = prev_portfolio.set_index("Symbol")

    else:
        prev_portfolio = None
    
    if prev_portfolio is None:
        selected = ranked.head(top_k)

        weight = 1.0 / top_k
        alloc = weight * portfolio_capital

        holdings = pd.DataFrame({
            "Symbol": selected.index,
            "Quantity": (alloc / last_prices[selected.index]).astype(int),
            "Avg_price": last_prices[selected.index].values
        }).set_index("Symbol")
    else:
        prev_symbols = set(prev_portfolio.index)
        current_symbols = list(ranked.index)

        common_symbols = prev_symbols.intersection(current_symbols)

        to_sell = prev_symbols - common_symbols

        if len(to_sell) == 0:
            holdings = prev_portfolio.copy()

        else:
            updated_holdings = prev_portfolio.loc[list(common_symbols)].copy()

            to_buy_candidates = [s for s in current_symbols if s not in prev_symbols]

            cash = 0.0

            for sym in to_sell:
                if sym not in last_prices.index:
                    continue

                price = last_prices.loc[sym]
                qty = prev_portfolio.loc[sym, "Quantity"]
                cash += price * qty
                
            slots_to_fill = len(to_sell)

            for sym in to_buy_candidates:
                if slots_to_fill == 0:
                    break

                price = last_prices.loc[sym]                
                if pd.isna(price) or price <= 0:
                    continue

                qty = int(cash / slots_to_fill / price)
                if qty <= 0:
                    continue

                spend = qty * price
                if spend <= cash:
                    to_buy.append(sym)
                    updated_holdings.loc[sym] = {
                        "Quantity": qty,
                        "Avg_price": price
                    }
                    cash -= spend
                    slots_to_fill -= 1

            if slots_to_fill > 0:
                print(f"Warning: {slots_to_fill} positions retained as cash")

            holdings = updated_holdings

    actual_sells = sorted(to_sell)
    actual_buys = sorted(to_buy)

    holdings["Momentum"] = momentum.reindex(holdings.index)
    holdings = holdings.sort_values(by="Momentum", ascending=False)
    
    output = pd.DataFrame({
        "Momentum": holdings["Momentum"],
        "Latest_Close": last_prices.reindex(holdings.index),
        "Quantity": holdings["Quantity"]
    })

    output["Momentum"] = pd.to_numeric(output["Momentum"], errors="coerce")
    output["Beats Benchmark"] = output["Momentum"] >= benchmark_return
    output["Market Value"] = output["Quantity"] * output["Latest_Close"]

    output_path = os.path.join(folder_path, "momentum_selected_stocks.xlsx")

    files_read_df = pd.DataFrame(
        {
            "Date": [d for d, _ in files],
            "Filename": [f for _, f in files],
        }
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        output.to_excel(writer, sheet_name="Selected_Stocks")
        
        pd.DataFrame({
            "As_of_Date": [start_date],
            "Lookback_Months": [lookback_months],
            "Top_K": [top_k],
            "Portfolio_Capital": [portfolio_capital],
            "Benchmark": [benchmark_return],
        }).to_excel(writer, sheet_name="Parameters", index=False)
    
        files_read_df.to_excel(writer, sheet_name="Files_Read", index=False)

    if progress_signal:
        progress_signal.emit(100)

    portfolio_path = os.path.join(folder_path, PORTFOLIO_FILE)

    portfolio_value = (holdings["Quantity"] * last_prices[holdings.index]).sum()

    new_summary_row = pd.DataFrame({
        "Run_Date": [start_date],
        "Portfolio_Value": [portfolio_value],
        "Absolute_Return": [0.0],
        "Relative_Return": [0.0],
    })

    rebalance_row = pd.DataFrame(
    [{
        "Run_Date": start_date,
        "Sell": ", ".join(actual_sells) if actual_sells else None,
        "Buy": ", ".join(actual_buys) if actual_buys else None
    }]
)

    if os.path.exists(portfolio_path):
        try:
            existing_summary = pd.read_excel(
                portfolio_path, sheet_name="Summary"
            )

            existing_rebalance = pd.read_excel(
                portfolio_path, sheet_name = "Rebalance Log"
            )

            existing_summary["Run_Date"] = pd.to_datetime(
                existing_summary["Run_Date"]
            )

            # Relative Return = current - last run
            last_portfolio_value = existing_summary.iloc[-1]["Portfolio_Value"]
            new_summary_row["Relative_Return"] = (
                ((portfolio_value - last_portfolio_value) / last_portfolio_value) * 100
            )

            # Absolute Return = current - first run of same year
            current_year = start_date.year
            same_year = existing_summary[
                existing_summary["Run_Date"].dt.year == current_year
            ]

            if not same_year.empty:
                first_year_value = same_year.iloc[0]["Portfolio_Value"]
                new_summary_row["Absolute_Return"] = (
                    ((portfolio_value - first_year_value) / first_year_value) * 100
                )

            summary_df = pd.concat(
                [existing_summary, new_summary_row],
                ignore_index=True
            )

            rebalance_df = pd.concat(
                [existing_rebalance, rebalance_row],
                ignore_index=True
            )

        except Exception:
            summary_df = new_summary_row
            rebalance_df = rebalance_row
    else:
        summary_df = new_summary_row
        rebalance_df = rebalance_row

    with pd.ExcelWriter(portfolio_path, engine="openpyxl") as writer:
        holdings.drop(columns=["Momentum"]).reset_index().rename(
            columns={"index": "Symbol"}
        ).to_excel(writer, sheet_name="Holdings", index=False)

        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        rebalance_df.to_excel(writer, sheet_name="Rebalance Log", index=False)

    return output_path