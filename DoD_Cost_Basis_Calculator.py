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


def calculate_security_price(ticker, ticker_type, date_of_death, decimal_places):
    """Calculate the security price based on the type of security and date of death."""
    try:
        security = yf.Ticker(ticker)
        nyse = mcal.get_calendar('NYSE')
        is_weekend_or_holiday = date_of_death.weekday() >= 5 or not nyse.valid_days(start_date=date_of_death,
                                                                                    end_date=date_of_death).size

        if is_mutual_fund(ticker_type):
            if is_weekend_or_holiday:
                pricing_date = get_previous_business_day(date_of_death)
            else:
                pricing_date = date_of_death

            # Get mutual fund closing price
            hist = security.history(start=pricing_date, end=pricing_date + timedelta(days=1))
            if hist.empty:
                return {
                    'Price': None,
                    'Note': "No data available for this date",
                    'Close': None,
                    'High': None,
                    'Low': None,
                    'Friday_High': None,
                    'Friday_Low': None,
                    'Friday_Close': None,
                    'Monday_High': None,
                    'Monday_Low': None,
                    'Monday_Close': None
                }

            # Round the closing price immediately when pulling from yfinance
            close_price = round(hist['Close'][0], decimal_places)

            return {
                'Price': close_price,
                'Note': f"Mutual Fund - Using {'Friday' if is_weekend_or_holiday else 'date of death'} closing price",
                'Close': close_price,
                'High': None,  # Not used for mutual funds
                'Low': None,  # Not used for mutual funds
                'Friday_High': None,
                'Friday_Low': None,
                'Friday_Close': close_price if is_weekend_or_holiday else None,
                'Monday_High': None,
                'Monday_Low': None,
                'Monday_Close': None
            }

        else:  # Stock or ETF
            if is_weekend_or_holiday:
                # Get Friday and Monday prices
                friday = get_previous_business_day(date_of_death)
                monday = get_next_business_day(date_of_death)

                friday_hist = security.history(start=friday, end=friday + timedelta(days=1))
                monday_hist = security.history(start=monday, end=monday + timedelta(days=1))

                if friday_hist.empty or monday_hist.empty:
                    return {
                        'Price': None,
                        'Note': "No data available for this date range",
                        'Close': None,
                        'High': None,
                        'Low': None,
                        'Friday_High': None,
                        'Friday_Low': None,
                        'Friday_Close': None,
                        'Monday_High': None,
                        'Monday_Low': None,
                        'Monday_Close': None
                    }

                # Round individual prices before calculations
                friday_high = round(friday_hist['High'][0], decimal_places)
                friday_low = round(friday_hist['Low'][0], decimal_places)
                friday_close = round(friday_hist['Close'][0], decimal_places)
                monday_high = round(monday_hist['High'][0], decimal_places)
                monday_low = round(monday_hist['Low'][0], decimal_places)
                monday_close = round(monday_hist['Close'][0], decimal_places)

                # Calculate averages using rounded prices
                friday_avg = (friday_high + friday_low) / 2
                monday_avg = (monday_high + monday_low) / 2
                final_price = round((friday_avg + monday_avg) / 2, decimal_places)

                return {
                    'Price': final_price,
                    'Note': "Weekend/Holiday price - Average of Previous/Next Business Day",
                    'Close': None,
                    'High': None,
                    'Low': None,
                    'Friday_High': friday_high,
                    'Friday_Low': friday_low,
                    'Friday_Close': friday_close,
                    'Monday_High': monday_high,
                    'Monday_Low': monday_low,
                    'Monday_Close': monday_close
                }
            else:
                # Get single day high/low average
                hist = security.history(start=date_of_death, end=date_of_death + timedelta(days=1))
                if hist.empty:
                    return {
                        'Price': None,
                        'Note': "No data available for this date",
                        'Close': None,
                        'High': None,
                        'Low': None,
                        'Friday_High': None,
                        'Friday_Low': None,
                        'Friday_Close': None,
                        'Monday_High': None,
                        'Monday_Low': None,
                        'Monday_Close': None
                    }

                # Round individual prices before calculations
                high_price = round(hist['High'][0], decimal_places)
                low_price = round(hist['Low'][0], decimal_places)
                close_price = round(hist['Close'][0], decimal_places)

                # Calculate average using rounded prices
                final_price = round((high_price + low_price) / 2, decimal_places)

                return {
                    'Price': final_price,
                    'Note': "Regular Trading Day High/Low Average",
                    'Close': close_price,
                    'High': high_price,
                    'Low': low_price,
                    'Friday_High': None,
                    'Friday_Low': None,
                    'Friday_Close': None,
                    'Monday_High': None,
                    'Monday_Low': None,
                    'Monday_Close': None
                }

    except Exception as e:
        return {
            'Price': None,
            'Note': f"Error processing {ticker}: {str(e)}",
            'Close': None,
            'High': None,
            'Low': None,
            'Friday_High': None,
            'Friday_Low': None,
            'Friday_Close': None,
            'Monday_High': None,
            'Monday_Low': None,
            'Monday_Close': None
        }


def main():
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

            # Calculate prices for each security
            results = []
            for _, row in df.iterrows():
                result_dict = calculate_security_price(row['Ticker'], row['Type'], date_of_death, decimal_places)
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

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")


if __name__ == "__main__":
    main()