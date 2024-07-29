# HackUTAR
Google Workspace Hackathon

# Automated Invoice Generation Tool

## Overview

This project provides an automated solution for generating invoices from data stored in Google Sheets. The tool extracts relevant information, calculates totals, generates invoices in a structured format, and sends them via email to clients. It is designed to streamline the invoicing process, reduce manual errors, and enhance efficiency for small to medium-sized enterprises.

## Features

- **Automated Data Extraction**: Pulls data from a Google Sheets spreadsheet.
- **Invoice Calculation**: Calculates item totals and overall invoice amounts.
- **Invoice Generation**: Creates invoices with itemized lists and totals.
- **Email Delivery**: Automatically sends generated invoices to clients via email.
- **Status Tracking**: Updates the spreadsheet with the status of each invoice, marking them as "Sent" once processed.

## How It Works

1. **Data Collection**: The tool reads data from a Google Sheets spreadsheet. Each row in the spreadsheet represents a line item for an invoice, including details such as invoice number, client information, item description, quantity, unit price, and status.
   
2. **Invoice Processing**: The script filters the data to process only "Pending" invoices. It calculates the total amount for each item and generates an invoice structure in HTML format.
   
3. **Email Invoices**: The generated invoices are emailed to the respective clients. The email contains a well-formatted HTML version of the invoice, including company branding and item details.

4. **Status Update**: After successfully sending the invoice, the tool updates the status of each line item in the spreadsheet to "Sent."

## Setup

### Prerequisites

- A Google account with access to Google Sheets and Google Apps Script.
- Basic knowledge of JavaScript and Google Apps Script.

### Installation

1. **Clone or Download the Repository**: 
   Download the code files to your local machine or clone the repository using Git.

2. **Google Apps Script Setup**:
   - Open your Google Sheets document.
   - Go to `Extensions > Apps Script` to open the Apps Script editor.
   - Copy and paste the script into the Apps Script editor.
   - Save the script with a meaningful name.

3. **Google Apps Script Permissions**:
   - When running the script for the first time, you will need to authorize it to access your Google Sheets and Gmail.
   - Follow the on-screen instructions to grant the necessary permissions.

### Configuration

- **Spreadsheet Structure**: Ensure your Google Sheets document has the following columns: `Invoice Number`, `Client Name`, `Client Email`, `Item Description`, `Quantity`, `Unit Price`, `Total Amount`, `Status`.
- **Email Setup**: Customize the email template and subject line as needed. Ensure that your Gmail account has sufficient quota to send emails.

## Usage

- **Run the Script**: Trigger the script manually from the Apps Script editor or set up a time-based trigger to automate the process periodically.
- **Monitor and Adjust**: Regularly check the Google Sheets and your email to ensure that invoices are being processed and sent correctly.

## Troubleshooting

- **Permissions Issues**: Ensure that the script has the necessary permissions to access your Google Sheets and Gmail.
- **Spreadsheet Errors**: Verify that the data in the spreadsheet is correctly formatted and that all necessary columns are present.
- **Email Sending Limits**: Be aware of Gmail's sending limits to avoid temporary blocks.

## Contribution

Contributions are welcome! If you have suggestions or improvements, please create a pull request or open an issue in the repository.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.

## Contact

For further information or support, please contact Tang Jia Hui at jiahuitang25@gmail.com.
