# Trader Sentiment Analysis - Code

## Files
- `analysis.py` - loads fear_greed_index.csv + historical_data.csv, merges them,
  computes summary stats, generates 5 charts (PNG) and the Excel workbook.
- `generate_report.js` - builds the Word report (.docx), embedding the charts
  and summary tables produced by analysis.py.

## Setup
```bash
pip install pandas numpy matplotlib openpyxl --break-system-packages
npm install -g docx
```

## Run
1. Place `fear_greed_index.csv` and `historical_data.csv` in the same folder.
2. Run the analysis:
   ```bash
   python3 analysis.py
   ```
   This produces: merged.csv, chart1-5_*.png, trader_sentiment_analysis.xlsx
3. Generate the Word report:
   ```bash
   node generate_report.js
   ```
   This produces: Trader_Sentiment_Analysis_Report.docx
   (expects the chart PNGs to be in the same folder - adjust paths in the
   script if needed)

## Notes
- `historical_data.csv` columns expected: Timestamp IST, Closed PnL, Size USD,
  Direction, Side, Account, Coin, etc. (standard Hyperliquid export).
- `fear_greed_index.csv` columns expected: date, classification.
- The report text (executive summary, recommendations) is hardcoded with
  numbers from the original dataset run - re-run analysis.py first and update
  the table values in generate_report.js if your data differs.
