This project is an extension of the Portfolio Tracker, designed to track a client’s Assets Under Management (AUM) across multiple asset classes, regions, and currencies on a daily basis since inception. The system ensures that portfolio values are accurately represented in the client's respective reporting currency.

The system is built using google sheets with backend coding is done in google app script. Instead of using external APIs, the system generates a temporary sheet within Google Sheets to fetch price details.

Data Sources and Processing

1. Input.csv


• Contains transaction details for each client.
   
• Data is taken from the client's brokerage account.
   
• you can refer the sheet for the data attributes and the last attributes which is live price is directly 
  incorporated from the security symbol using google sheet formula.


2. Processing Flow

• Upon execution, a pop-up window prompts the user to enter the Client Name and Client ID.The system validates these inputs against Input.csv to ensure accuracy.Once validated, the output sheet is generated with daily AUM tracking since inception.
 
3.Output Sheet – A sheet with Client Name will be generated with historical AUM values
