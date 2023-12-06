import pandas as pd

def load_and_clean_data():
    
    orders = pd.read_excel('mnt/data/test-orders.xlsx', )
    currency_rates = pd.read_excel('mnt/data/test-currency-rates.xlsx',)
    affiliate_rates = pd.read_excel('mnt/data/test-affiliate-rates.xlsx', )

    orders.drop_duplicates(inplace=True)
    currency_rates.drop_duplicates(inplace=True)
    affiliate_rates.drop_duplicates(inplace=True)


    orders['Order Date'] = pd.to_datetime(orders['Order Date'])
    currency_rates['date'] = pd.to_datetime(currency_rates['date'])
    affiliate_rates['Start Date'] = pd.to_datetime(affiliate_rates['Start Date'])

    return orders, currency_rates, affiliate_rates

def convert_to_eur(row, currency_rates):
    if row['Currency'] == 'EUR':
        return row['Order Amount']

    order_date = pd.to_datetime(row['Order Date'])
    relevant_rates = currency_rates[currency_rates['date'] <= order_date]

    if relevant_rates.empty:
        return None  

    if row['Currency'] == 'USD':
        latest_rate = relevant_rates['USD'].iloc[-1] if not relevant_rates['USD'].empty else None
    elif row['Currency'] == 'GBP':
        latest_rate = relevant_rates['GBP'].iloc[-1] if not relevant_rates['GBP'].empty else None
    else:
        return None 

    return row['Order Amount'] * latest_rate if latest_rate is not None else None



def calculate_fees(row, affiliate_rates):
    rates = affiliate_rates[(affiliate_rates['Affiliate ID'] == row['Affiliate ID']) & 
                            (affiliate_rates['Start Date'] <= row['Order Date'])]

    if rates.empty:
        return row  

    rate_info = rates.iloc[-1]

    row['Processing Fee'] = row['Order Amount EUR'] * rate_info['Processing Rate'] if 'Processing Rate' in rate_info else 0
    row['Refund Fee'] = rate_info['Refund Fee'] if row['Order Status'] == 'Refunded' and 'Refund Fee' in rate_info else 0
    row['Chargeback Fee'] = rate_info['Chargeback Fee'] if row['Order Status'] == 'Chargeback' and 'Chargeback Fee' in rate_info else 0

    return row


def generate_weekly_reports(orders, affiliate_rates):
    reports = {}
    for _, affiliate in affiliate_rates[['Affiliate ID', 'Affiliate Name']].drop_duplicates().iterrows():
        affiliate_orders = orders[orders['Affiliate ID'] == affiliate['Affiliate ID']]
        affiliate_orders.set_index('Order Date', inplace=True)

        weekly_report = affiliate_orders.resample('W-SUN').agg({
            'Order Number': 'count',
            'Order Amount EUR': 'sum',
            'Processing Fee': 'sum',
            'Refund Fee': 'sum',
            'Chargeback Fee': 'sum'
        }).rename(columns={'Order Number': 'Number of Orders'})

        weekly_report['Week'] = weekly_report.index.to_series().apply(
            lambda x: f"{(x - pd.Timedelta(days=x.weekday())).strftime('%d-%m-%Y')} - {(x + pd.Timedelta(days=6-x.weekday())).strftime('%d-%m-%Y')}")

        reports[affiliate['Affiliate Name']] = weekly_report

    return reports

def save_reports(reports):
    for affiliate_name, report in reports.items():
        report.to_excel(f'mnt/data/{affiliate_name}_report.xlsx', index=False)

def main():
    orders, currency_rates, affiliate_rates = load_and_clean_data()
    orders['Order Amount EUR'] = orders.apply(convert_to_eur, axis=1, args=(currency_rates,))
    orders = orders.apply(calculate_fees, axis=1, args=(affiliate_rates,))
    reports = generate_weekly_reports(orders, affiliate_rates)
    save_reports(reports)
    print("Reports have been generated.")

if __name__ == "__main__":
    main()
