# NSE Momentum Portfolio Selector

A Python-based desktop application (PySide6) for selecting and tracking
NSE stocks using a momentum-driven strategy with portfolio persistence
across runs.

This project is designed for **research, strategy validation, and
controlled portfolio simulation**, not live trading.

------------------------------------------------------------------------

## Core Strategy Overview

1.  Read daily NSE bhavcopy-style CSV files named `YYYYMMDD_NSE.csv`
2.  Restrict universe to `SERIES = EQ` or `BE`
3.  Use files ≤ selected start date and within lookback window
4.  Compute momentum: `(Last_Close / First_Close) - 1`
5.  Rank stocks by momentum (descending)
6.  Apply budget constraints
7.  Select Top-K + buffer
8.  Persist portfolio across runs
9.  Track holdings, portfolio value, and rebalance decisions

------------------------------------------------------------------------

## Application Architecture

### main.py

-   PySide6 GUI
-   Parameter input
-   Background processing thread

### processor.py

-   File parsing
-   Momentum computation
-   Budget-aware selection
-   Excel output generation

------------------------------------------------------------------------

## Input Data Format

CSV columns:

    SYMBOL,SERIES,CLOSE

Filename format:

    YYYYMMDD_NSE.csv

------------------------------------------------------------------------

## Output Files

### momentum_selected_stocks.xlsx

-   Selected_Stocks
-   Parameters
-   Files_Read

### portfolio_tracker.xlsx

-   Holdings
-   Summary
-   Holdings_History
-   Rebalance_Log

------------------------------------------------------------------------

## How to Run

``` bash
python main.py
```

------------------------------------------------------------------------

## Dependencies

-   pandas
-   openpyxl
-   PySide6

Install with:

``` bash
pip install -r requirements.txt
```

------------------------------------------------------------------------

## ⚠️ Disclaimer

This project is for educational and research purposes only.
