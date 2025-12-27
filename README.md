# NSE Momentum Stock Selector

This project implements a momentum-based stock selection framework
for NSE-listed equities using daily bhavcopy data.

## Current Features
- PySide6-based desktop GUI
- Date-based NSE CSV ingestion
- Lookback-window momentum calculation
- Budget-aware stock filtering
- Portfolio state persistence across runs
- Rebalance comparison logging (Buy/Sell lists)
- Portfolio performance tracking

## Status
This repository reflects the **core system architecture and data pipeline**.
Final execution and allocation refinements are intentionally excluded.

## Data
Daily NSE CSV files are expected in the format:
YYYYMMDD_NSE.csv

(Data files are not included in this repository.)

## Disclaimer
This project is for research and educational purposes only.
It does not constitute investment advice.