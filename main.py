# Required Libraries
import pandas as pd
from datetime import datetime, timedelta


def main():
    # File paths
    path_stops_report = "Stops_Report.csv"
    path_phases = "phases.xlsx"
    path_region_lookup = "region lookup.xlsx"
    path_invoices_report = "Invoices_Report.csv"
    path_survey_report = "No_Sale_Survey.csv"

    # Load Data
    stops_report = pd.read_csv(path_stops_report)
    phases = pd.read_excel(path_phases)
    region_lookup = pd.read_excel(path_region_lookup)
    invoices_report = pd.read_csv(path_invoices_report)
    survey_report = pd.read_csv(path_survey_report)

    # Date for Yesterday
    yesterday = datetime.now() - timedelta(days=1)
    yesterday = pd.to_datetime(yesterday.strftime('%Y-%m-%d'))  # Convert yesterday to a datetime object

    # Phase and Day for Yesterday
    phase_info_yesterday = phases[phases['Date'] == pd.Timestamp(yesterday)]
    day_mapping = {'Sunday': 1, 'Monday': 2, 'Tuesday': 3, 'Wednesday': 4, 'Thursday': 5, 'Friday': 6, 'Saturday': 7}
    phase_mapping = {'One': 1, 'Two': 2, 'Three': 3, 'Four': 4}
    day_yesterday_num = day_mapping[phase_info_yesterday['Day Of Week'].values[0]]
    phase_yesterday_num = phase_mapping[phase_info_yesterday['Phase'].values[0]]

    # Filter Stops Report
    filtered_report = stops_report[(stops_report['Phase'] == phase_yesterday_num) & (stops_report['Day Of Week'] == day_yesterday_num)]

    # Add Region column
    merged_report = pd.merge(filtered_report, region_lookup, on='Territory', how='left')

    # Add Expected Day of Sale column
    merged_report['Expected Day of Sale'] = yesterday

    # Add Invoice on Expected Day column
    invoices_report['Date'] = pd.to_datetime(invoices_report['Date'])
    merged_report['Expected Day of Sale'] = pd.to_datetime(merged_report['Expected Day of Sale'])
    merged_report = pd.merge(merged_report, invoices_report[['Customer ID', 'Date']],
                         left_on=['Customer ID', 'Expected Day of Sale'],
                         right_on=['Customer ID', 'Date'],
                         how='left')
    merged_report.rename(columns={'Date': 'Invoice on Expected Day'}, inplace=True)

    # Add Sale Made on Expected Day? column
    merged_report['Sale Made on Expected Day?'] = merged_report['Invoice on Expected Day'].apply(lambda x: 'No' if pd.isnull(x) else 'Yes')

    # Add Net Sales for Expected Day column
    net_sales = invoices_report[invoices_report['Date'] == yesterday].groupby('Customer ID')['Total Cases'].sum().reset_index()
    net_sales.rename(columns={'Total Cases': 'Net Sales for Expected Day'}, inplace=True)
    merged_report = pd.merge(merged_report, net_sales, on='Customer ID', how='left')

    # Add Positive Sale column
    merged_report['Positive Sale'] = merged_report['Net Sales for Expected Day'].apply(lambda x: 1 if x > 0 else (0 if x <= 0 else float('nan')))

    # Add Last Sale Date column
    last_sale_date = invoices_report[invoices_report['Total Cases'] > 0].groupby('Customer ID')['Date'].max().reset_index()
    last_sale_date.rename(columns={'Date': 'Last Sale Date'}, inplace=True)
    merged_report = pd.merge(merged_report, last_sale_date, on='Customer ID', how='left')
    merged_report['Last Sale Date'] = merged_report.apply(lambda row: 'No Sale Last 6 Days' if pd.isnull(row['Invoice on Expected Day']) else row['Last Sale Date'], axis=1)

    # Add Last Sale Date Amount column
    last_sale_date_amount = invoices_report[invoices_report['Total Cases'] > 0].sort_values('Date').groupby('Customer ID').tail(1)[['Customer ID', 'Total Cases']]
    last_sale_date_amount.rename(columns={'Total Cases': 'Last Sale Date Amount'}, inplace=True)
    merged_report = pd.merge(merged_report, last_sale_date_amount, on='Customer ID', how='left')

    # Add Sale Complete On Expected Day? column
    merged_report['Sale Complete On Expected Day?'] = merged_report.apply(lambda row: 'Completed' if row['Sale Made on Expected Day?'] == 'Yes' and row['Last Sale Date Amount'] > 0
                                                                      else ('Service Completed In Last 6 Days' if row['Sale Made on Expected Day?'] == 'No' and row['Last Sale Date Amount'] > 0
                                                                      else '6 Day Non Buy'), axis=1)

    # Add Survey Results column
    survey_report['Date'] = pd.to_datetime(survey_report['Date'])
    merged_report = pd.merge(merged_report, survey_report[['Customer ID', 'Date', 'Please select a reason why no sale took place:']],
                         left_on=['Customer ID', 'Expected Day of Sale'],
                         right_on=['Customer ID', 'Date'],
                         how='left')
    merged_report.rename(columns={'Please select a reason why no sale took place:': 'Survey Results'}, inplace=True)
    merged_report.drop(columns=['Date'], inplace=True)

    # Save the final dataframe to a CSV file
    merged_report.to_csv('Stops_Worksheet.csv', index=False)

if __name__ == "__main__":
    main()
