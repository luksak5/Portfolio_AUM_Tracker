This project is an extension of the Portfolio Tracker, designed to track a client’s Assets Under Management (AUM) across multiple asset classes, regions, and currencies on a daily basis since inception. The system ensures that portfolio values are accurately represented in the client's respective reporting currency.

The system is built using google sheets with backend coding is done in google app script. Instead of using external APIs, the system generates a temporary sheet within Google Sheets to fetch price details.

Data Sources and Processing

1. Input.csv

Stores client details and transaction history.
Includes fields such as Client Name, Client ID, Asset Type, Security Name, Ticker, Transaction Type, Units, Purchase Price, and Currency.
Data is either manually entered or sourced directly from the client’s brokerage account.


Processing Flow

Upon execution, a pop-up window prompts the user to enter the Client Name and Client ID.
The system validates these inputs against Input.csv to ensure accuracy.
Once validated, the output sheet is generated with daily AUM tracking since inception.
Output Sheet – Daily AUM Report

Displays historical AUM values, tracking portfolio performance over time.
Maintains cumulative units per ticker, ensuring correct asset allocation.
Converts and normalizes values into the client's reporting currency.
