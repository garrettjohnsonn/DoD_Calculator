import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import streamlit as st
from io import BytesIO
import pandas_market_calendars as mcal


def is_mutual_fund(ticker_type):
    """Check if a security is a mutual fund based on its type."""
    return ticker_type.lower() == 'mutual fund'


def get_next_business_day(date):
    """Get the next business day after a given date."""
    nyse = mcal.get_calendar('NYSE')
    next_day = date + timedelta(days=1)
    while not nyse.valid_days(start_date=next_day, end_date=next_day).size:
        next_day += timedelta(days=1)
    return next_day


def get_previous_business_day(date):
    """Get the previous business day before a given date."""
    nyse = mcal.get_calendar('NYSE')
    prev_day = date - timedelta(days=1)
    while not nyse.valid_days(start_date=prev_day, end_date=prev_day).size:
        prev_day -= timedelta(days=1)
    return prev_day


def empty_result(note="No data available"):
    """Return an empty result dictionary with a note."""
    return {
        'Price': None, 'Note': note,
        'Close': None, 'High': None, 'Low': None,
        'Friday_High': None, 'Friday_Low': None, 'Friday_Close': None,
        'Monday_High': None, 'Monday_Low': None, 'Monday_Close': None
    }


def get_ticker_data(ticker, batch_data, single_ticker):
    """Extract data for a specific ticker from batch download results."""
    if single_ticker:
        return batch_data
    else:
        try:
            return batch_data[ticker]
        except KeyError:
            return pd.DataFrame()


def get_day_data(ticker_data, target_date):
    """Get data for a specific date from ticker data."""
    if isinstance(target_date, pd.Timestamp):
        target_date = target_date.date()
    return ticker_data[ticker_data.index.date == target_date]


def calculate_security_price(ticker, ticker_type, date_of_death, decimal_places,
                             batch_data, is_weekend_or_holiday, single_ticker):
    """Calculate the security price using pre-fetched batch data."""
    try:
        ticker_data = get_ticker_data(ticker, batch_data, single_ticker)
        if ticker_data.empty:
            return empty_result(f"No data available for {ticker}")

        if is_mutual_fund(ticker_type):
            if is_weekend_or_holiday:
                pricing_date = get_previous_business_day(date_of_death)
            else:
                pricing_date = date_of_death

            hist = get_day_data(ticker_data, pricing_date)
            if hist.empty:
                return empty_result("No data available for this date")

            close_price = round(float(hist['Close'].iloc[0]), decimal_places)

            return {
                'Price': close_price,
                'Note': f"Mutual Fund - Using {'Friday' if is_weekend_or_holiday else 'date of death'} closing price",
                'Close': close_price,
                'High': None, 'Low': None,
                'Friday_High': None, 'Friday_Low': None,
                'Friday_Close': close_price if is_weekend_or_holiday else None,
                'Monday_High': None, 'Monday_Low': None, 'Monday_Close': None
            }

        else:  # Stock or ETF
            if is_weekend_or_holiday:
                friday = get_previous_business_day(date_of_death)
                monday = get_next_business_day(date_of_death)

                friday_hist = get_day_data(ticker_data, friday)
                monday_hist = get_day_data(ticker_data, monday)

                if friday_hist.empty or monday_hist.empty:
                    return empty_result("No data available for this date range")

                friday_high = round(float(friday_hist['High'].iloc[0]), decimal_places)
                friday_low = round(float(friday_hist['Low'].iloc[0]), decimal_places)
                friday_close = round(float(friday_hist['Close'].iloc[0]), decimal_places)
                monday_high = round(float(monday_hist['High'].iloc[0]), decimal_places)
                monday_low = round(float(monday_hist['Low'].iloc[0]), decimal_places)
                monday_close = round(float(monday_hist['Close'].iloc[0]), decimal_places)

                friday_avg = (friday_high + friday_low) / 2
                monday_avg = (monday_high + monday_low) / 2
                final_price = round((friday_avg + monday_avg) / 2, decimal_places)

                return {
                    'Price': final_price,
                    'Note': "Weekend/Holiday price - Average of Previous/Next Business Day",
                    'Close': None, 'High': None, 'Low': None,
                    'Friday_High': friday_high, 'Friday_Low': friday_low, 'Friday_Close': friday_close,
                    'Monday_High': monday_high, 'Monday_Low': monday_low, 'Monday_Close': monday_close
                }
            else:
                hist = get_day_data(ticker_data, date_of_death)
                if hist.empty:
                    return empty_result("No data available for this date")

                high_price = round(float(hist['High'].iloc[0]), decimal_places)
                low_price = round(float(hist['Low'].iloc[0]), decimal_places)
                close_price = round(float(hist['Close'].iloc[0]), decimal_places)
                final_price = round((high_price + low_price) / 2, decimal_places)

                return {
                    'Price': final_price,
                    'Note': "Regular Trading Day High/Low Average",
                    'Close': close_price, 'High': high_price, 'Low': low_price,
                    'Friday_High': None, 'Friday_Low': None, 'Friday_Close': None,
                    'Monday_High': None, 'Monday_Low': None, 'Monday_Close': None
                }

    except Exception as e:
        return empty_result(f"Error processing {ticker}: {str(e)}")


def main():
    st.set_page_config(page_title="Step Up Calculator")
    st.title("DoD Step-Up Cost Basis Calculator")

    # File upload
    uploaded_file = st.file_uploader(
        "Upload Excel file with columns: Ticker, Shares, Type (List Previously Mentioned Titles in First Cell of Each Column)",
        type=['xlsx'])

    # Date input
    date_of_death = st.date_input("Date of Death")

    # Add decimal places input
    decimal_places = st.number_input("Number of decimal places for rounding", min_value=0, max_value=10, value=2)

    if uploaded_file and date_of_death:
        try:
            # Read input file
            df = pd.read_excel(uploaded_file)

            # Validate required columns
            required_columns = ['Ticker', 'Shares', 'Type']
            if not all(col in df.columns for col in required_columns):
                st.error("Excel file must contain 'Ticker', 'Shares', and 'Type' columns")
                return

            # Determine if date of death falls on a weekend or holiday
            nyse = mcal.get_calendar('NYSE')
            is_weekend_or_holiday = date_of_death.weekday() >= 5 or not nyse.valid_days(
                start_date=date_of_death, end_date=date_of_death).size

            # Determine date range for batch download
            if is_weekend_or_holiday:
                start_date = get_previous_business_day(date_of_death)
                end_date = get_next_business_day(date_of_death) + timedelta(days=1)
            else:
                start_date = date_of_death
                end_date = date_of_death + timedelta(days=1)

            # Batch download all tickers in a single API call
            tickers = df['Ticker'].unique().tolist()
            single_ticker = len(tickers) == 1

            if single_ticker:
                batch_data = yf.download(
                    tickers[0], start=start_date, end=end_date, auto_adjust=False
                )
            else:
                batch_data = yf.download(
                    tickers, start=start_date, end=end_date, auto_adjust=False, group_by='ticker'
                )

            # Calculate prices for each security
            results = []
            for _, row in df.iterrows():
                result_dict = calculate_security_price(
                    row['Ticker'], row['Type'], date_of_death, decimal_places,
                    batch_data, is_weekend_or_holiday, single_ticker
                )
                price = result_dict['Price']

                result = {
                    'Date': date_of_death,
                    'Ticker': row['Ticker'],
                    'Shares': row['Shares'],
                    'Price': price,
                    'Total Value': round(price * row['Shares'], decimal_places) if price else None,
                }

                # Add price details based on security type and whether it's a weekend or holiday
                if is_mutual_fund(row['Type']):
                    result['Closing Price'] = result_dict['Close']
                else:
                    if result_dict['Friday_High'] is not None:  # Weekend/Holiday case
                        result['Friday High'] = result_dict['Friday_High']
                        result['Friday Low'] = result_dict['Friday_Low']
                        result['Monday High'] = result_dict['Monday_High']
                        result['Monday Low'] = result_dict['Monday_Low']
                    else:  # Regular trading day
                        result['High'] = result_dict['High']
                        result['Low'] = result_dict['Low']

                result['Note'] = result_dict['Note']
                results.append(result)

            # Create results DataFrame
            results_df = pd.DataFrame(results)

            # Reorder columns to move 'Note' to the last column
            cols = results_df.columns.tolist()
            cols.append(cols.pop(cols.index('Note')))
            results_df = results_df[cols]

            # Display results in the app
            st.subheader("Results")
            st.dataframe(results_df)

            # Create download button
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                results_df.to_excel(writer, index=False)

            st.download_button(
                label="Download Results",
                data=output.getvalue(),
                file_name=f"security_prices_{date_of_death}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.balloons()

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")


if __name__ == "__main__":
    main()
