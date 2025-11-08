# Stock Evaluation – Google Sheets + Earnings Alerts

Automates a daily pipeline that:
1. Reads a list of tickers from a Google Sheet  
2. Pulls fresh metrics with **yfinance** (price, EPS, FCF, ROE, Debt/Equity, P/E)  
3. Writes the results back to the Sheet  
4. If any ticker had earnings **yesterday**, emails you a short **NewsAPI** roundup  

All secrets are loaded from environment variables (via `.env` + `python-dotenv`).  
No keys are committed to Git.

---

## 📊 Features

- ✅ Pulls **Close price**, **EPS (diluted)**, **Free Cash Flow**, **ROE**, **Debt/Equity**, **P/E**
- ✅ Batch-writes results to your Google Sheet (`Sheet1!B:G`)
- ✅ Detects **yesterday’s** earnings and sends a **Gmail SMTP** email with top 3 news links
- ✅ Runs locally or on PythonAnywhere (scheduled daily)
- ✅ Secrets via env vars (safe for GitHub)

---

## ⚙️ Requirements

Install all dependencies with:

```bash
pip install -r requirements.txt
