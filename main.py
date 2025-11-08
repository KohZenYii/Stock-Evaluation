import requests
import yfinance as yf
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
import smtplib
import os
from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv())

SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")  # absolute path on PA
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
TICKER_INPUT_RANGE = 'Sheet1!A2:A'
SHEET_NAME = "Sheet1"
MY_EMAIL = os.getenv("MY_EMAIL")

PASSWORD = os.getenv("PASSWORD")  # use an App Password
NEWS_API_KEY = os.getenv("NEWS_API_KEY")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")


# time
today_date = datetime.today().date()
yesterday_date = today_date - timedelta(1)  # Using yesterday closing price as I don't have real time data access
googlesheets_data = []


########## News Mechanism #################
def news(ticker):
    news_params = {
        "q": f"{ticker}",
        "apiKey": NEWS_API_KEY,
        "sortBy": "popularity",
        "language": 'en'
    }

    news_response = requests.get("https://newsapi.org/v2/everything", params=news_params)
    news_response.raise_for_status()
    news_data = news_response.json()
    total_news = news_data['totalResults']
    if total_news < 3:
        loop = total_news
    else:
        loop = 3
    news_summary = f"Following news link for {ticker}\n\n"  # Initialize an empty string to accumulate news information
    i = 0
    index = 0
    while i <= loop:
        try:
            news = news_response.json()['articles'][i]
        except IndexError:
            break
        title = news['title']
        if title == '[Removed]':
            i += 1
            continue
        link = news['url']
        i += 1
        # Append the formatted news information to the email body
        news_summary += f"{index + 1}. Title: {title}\n   Link: {link}\n\n"
        index += 1

    return news_summary


def connect_googlesheets():
    try:
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        print("Successfully connected to Google Sheets API.")
    except Exception as e:
        print(f"Error connecting to Google Sheets API. Check credentials.json and SCOPES: {e}")
        exit()

    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=TICKER_INPUT_RANGE
    ).execute()

    # ---- Get headers to map column positions ----
    header_result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!1:1"
    ).execute()
    return service, result
    # headers = header_result.get('values', [[]])[0]
    # header_map = {h.strip().lower(): i for i, h in enumerate(headers)}


######## get _financials here ##########

# after get_financials
def update_googlesheets(service, result):
    # Write all rows in one go to B2:G (aligned with A2:A tickers)
    start_row = 2
    end_row = start_row + len(googlesheets_data) - 1
    target_range = f"Sheet1!B{start_row}:G{end_row}"

    body = {
        "valueInputOption": "USER_ENTERED",
        "data": [{
            "range": target_range,
            "majorDimension": "ROWS",
            "values": googlesheets_data
        }]
    }
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=body
    ).execute()


# ---- Get tickers ----
def get_financials(result):
    values = result.get('values', [])
    stocks = [row[0].strip() for row in values if row and row[0].strip()]
    today_earnings_stocks = []

    for ticker in stocks:
        print(ticker)
        stock = yf.Ticker(ticker)
        # Stock Price
        stock_price = round(stock.history(period="1d")['Close'].iloc[-1], 2)
        print(stock_price)

        #### Financial Statements #####
        income_statement = stock.income_stmt
        balance_sheet = stock.balancesheet
        cashflow_statement = stock.cashflow

        #### ROE CALCULATION #####
        # Note: loc is for label (name of the row/column), iloc is for index (position of the row/column)
        net_income = income_statement.loc['Net Income'].iloc[0]
        stockholders_equity = balance_sheet.loc['Stockholders Equity'].iloc[0]
        roe = net_income / stockholders_equity

        ####  EPS CALCULATION #####
        eps = income_statement.loc['Diluted EPS'].iloc[0]

        ##### FCF CALCULATION #####
        fcf = cashflow_statement.loc['Free Cash Flow'].iloc[0]

        #### Debt/Equtiy ####
        total_debt = balance_sheet.loc['Total Debt'].iloc[0]
        debt_equity = total_debt / stockholders_equity
        print(f"Debt/Equity: {debt_equity:.2f}")

        #### PE RATIO ####
        pe = stock.info.get('trailingPE')

        #### earnings date ####
        earnings_date = stock.earnings_dates.index[0]
        if earnings_date.date() == yesterday_date:
            today_earnings_stocks.append(ticker)

        googlesheets_data.append([
            round(stock_price, 2),  # Price (B)
            round(eps, 2),  # EPS (C)
            f"{fcf:,.0f}" if fcf is not None else "",  # Free Cash Flow with commas (D)
            round(roe * 100, 2) if roe is not None else "",  # ROE as % (E)
            round(debt_equity, 2) if debt_equity is not None else "",  # Debt/Equity (F)
            pe  # P/E (G)
        ])

    return today_earnings_stocks


### EMAIL MECHANISM ###
def send_email(today_earnings_stocks):
    email_title_stock = ",".join(today_earnings_stocks)
    email_title = email_title_stock + " earnings report is out!!"
    email_body = ""
    for today_stock in today_earnings_stocks:
        email_body += news(today_stock)

    # deal with non-ascii characters to avoid Unicode errors
    email_title_ascii = email_title.encode("ascii", "ignore").decode()
    email_body_ascii = email_body.encode("ascii", "ignore").decode()

    with smtplib.SMTP("smtp.gmail.com", port=587) as connection:
        connection.starttls()  ## encrypted the email you are sending
        connection.login(user=MY_EMAIL, password=PASSWORD)
        connection.sendmail(from_addr=MY_EMAIL, to_addrs="amos.koh02@gmail.com",
                            msg=f'Subject:{email_title_ascii}'
                                f'\n\n{email_body_ascii}')
        connection.close()


service, result = connect_googlesheets()
today_earnings_stocks = get_financials(result)
update_googlesheets(service, result)
if today_earnings_stocks:
    send_email(today_earnings_stocks)


