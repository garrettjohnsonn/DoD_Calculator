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


def calculate_security_price(ticker, ticker_type, date_of_death):
    """Calculate the security price based on the type of security and date of death."""
    try:
        security = yf.Ticker(ticker)
        nyse = mcal.get_calendar('NYSE')
        is_weekend_or_holiday = date_of_death.weekday() >= 5 or not nyse.valid_days(start_date=date_of_death, end_date=date_of_death).size

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

            return {
                'Price': round(hist['Close'][0], 2),
                'Note': f"Mutual Fund - Using {'Friday' if is_weekend_or_holiday else 'date of death'} closing price",
                'Close': round(hist['Close'][0], 2),
                'High': None,  # Not used for mutual funds
                'Low': None,  # Not used for mutual funds
                'Friday_High': None,
                'Friday_Low': None,
                'Friday_Close': round(hist['Close'][0], 2) if is_weekend_or_holiday else None,
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

                friday_avg = (friday_hist['High'][0] + friday_hist['Low'][0]) / 2
                monday_avg = (monday_hist['High'][0] + monday_hist['Low'][0]) / 2

                return {
                    'Price': round((friday_avg + monday_avg) / 2, 2),
                    'Note': "Weekend/Holiday price - Average of Previous/Next Business Day",
                    'Close': None,
                    'High': None,
                    'Low': None,
                    'Friday_High': round(friday_hist['High'][0], 2),
                    'Friday_Low': round(friday_hist['Low'][0], 2),
                    'Friday_Close': round(friday_hist['Close'][0], 2),
                    'Monday_High': round(monday_hist['High'][0], 2),
                    'Monday_Low': round(monday_hist['Low'][0], 2),
                    'Monday_Close': round(monday_hist['Close'][0], 2)
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

                return {
                    'Price': round((hist['High'][0] + hist['Low'][0]) / 2, 2),
                    'Note': "Regular Trading Day High/Low Average",
                    'Close': round(hist['Close'][0], 2),
                    'High': round(hist['High'][0], 2),
                    'Low': round(hist['Low'][0], 2),
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
    st.title("Date of Death New Cost Basis Calculator")

    # File upload
    uploaded_file = st.file_uploader("Upload Excel file with columns: Ticker, Shares, Type", type=['xlsx'])

    # Date input
    date_of_death = st.date_input("Date of Death")

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
                result_dict = calculate_security_price(row['Ticker'], row['Type'], date_of_death)
                price = result_dict['Price']

                result = {
                    'Date': date_of_death,
                    'Ticker': row['Ticker'],
                    'Shares': row['Shares'],
                    'Price': price,
                    'Total Value': round(price * row['Shares'], 2) if price else None,
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