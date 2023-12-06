# Order Processing and Reporting

This repository contains a Python application designed for the automated processing of order data. The application cleans and transforms order data, converts currencies to EUR, calculates various fees, and generates comprehensive weekly reports for each affiliate.

The reports provide insights into the number of orders, total order amount in EUR, and detailed fee breakdowns. This tool is particularly useful for finance and sales departments looking to streamline their reporting process.

## Features

- **Data Loading**: Load data from multiple sources and formats.
- **Data Cleaning**: Remove duplicates, handle missing values, and correct inconsistencies.
- **Currency Conversion**: Convert order amounts from USD and GBP to EUR using daily exchange rates.
- **Fee Calculation**: Compute processing, refund, and chargeback fees for each order.
- **Weekly Reporting**: Aggregate and summarize weekly data for each affiliate in Excel format.

### Prerequisites

Before running the application, ensure you have the following prerequisites installed:

- Python 3.x
- Pandas library
- OpenPyXL library

## File Structure
-test-orders.xlsx: Order data including dates, amounts, and currencies.
-test-currency-rates.xlsx: Daily currency exchange rates.
-test-affiliate-rates.xlsx: Affiliate-specific fee rates.


## Reports
-Week range (e.g., 01-10-2023 - 07-10-2023)
-Number of Orders
-Total Order Amount in EUR
-Processing Fees
-Refund Fees
-Chargeback Fees.

