import pandas as pd
import yfinance as yf
from datetime import datetime

def calculate_shares(file_path, start_date, end_date, investment_amount):
    # Read the stocks.csv file
    stocks_df = pd.read_csv(file_path)

    # Prepare results list
    results = []

    # Convert dates to datetime objects
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    today = pd.Timestamp.today().normalize()  # Normalize to avoid time mismatches

    # Iterate over each ticker in the CSV file
    for _, row in stocks_df.iterrows():
        original_ticker = row['Ticker']  # Store the original ticker format
        ticker = original_ticker + ".NS"  # Adjust for NSE format
        weightage = row['Weightage']  # Weightage for the ticker (exact value as in CSV)
        print(f"Fetching data for {ticker}...")

        ticker_result = {
            "Ticker": original_ticker,  # Use original ticker without .NS
            "Weightage": weightage
        }

        # Iterate over the two dates
        for date in [start_date, end_date]:
            try:
                # Check if the date is today
                if date == today:
                    stock_data = yf.Ticker(ticker)
                    live_price = stock_data.history(period="1d")['Close']
                    if not live_price.empty:
                        price = live_price.iloc[-1]  # Fetch the latest available price
                    else:
                        raise ValueError("Live price not available")
                else:
                    # Fetch historical stock data for the specified date
                    stock_data = yf.download(ticker, start=date, end=date + pd.Timedelta(days=1))

                    if stock_data.empty:
                        raise ValueError(f"No historical data for {date}")

                    price = stock_data.loc[date.strftime('%Y-%m-%d'), 'Close']

                # Ensure price is a valid scalar value
                if isinstance(price, pd.Series):
                    price = price.iloc[0]

                # Calculate the number of shares
                shares_purchasable = round(investment_amount / price, 111)

                # Store result for this ticker and date
                ticker_result[date.strftime('%Y-%m-%d')] = shares_purchasable

            except (KeyError, ValueError):
                print(f"No data available for {ticker} on {date}. Skipping...")
                ticker_result[date.strftime('%Y-%m-%d')] = "No Data"
            except Exception as e:
                print(f"An error occurred for {ticker} on {date}: {e}")
                ticker_result[date.strftime('%Y-%m-%d')] = "Error"

        # Append the ticker result to the list of results
        results.append(ticker_result)

    # Create results DataFrame
    results_df = pd.DataFrame(results)

    # Display the results in the required format
    print("\nResult:\n")
    print(results_df)

    # Save results to an Excel file
    output_file = "shares_purchasable.xlsx"
    results_df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

# Input details
file_path = "stocks.csv"  # Ensure the file contains 'Ticker' and 'Weightage' columns
start_date = input("Enter the first date (YYYY-MM-DD): ")
end_date = input("Enter the second date (YYYY-MM-DD): ")
investment_amount = float(input("Enter the investment amount: "))

calculate_shares(file_path, start_date, end_date, investment_amount)
