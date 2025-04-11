# Email Extract

## Features

- Connects to an Outlook inbox using the MAPI interface.
- Filters emails within a specified date range.
- Extracts unique sender email addresses and their display names.
- Saves the extracted data to a CSV file.

## Requirements

- Python 3.11 or higher
- Microsoft Outlook installed on the windows system
- The `pywin32` library for interacting with Outlook

## Installation
- Install required dependency `pywin32`

## Usage
### Open the main.py file and update the following variables as needed:

  - FROM_DATE: The start date and time for filtering emails (format: MM/DD/YYYY HH:MM AM/PM).
  - TO_DATE: The end date and time for filtering emails (format: MM/DD/YYYY HH:MM AM/PM).
  - output_csv: The file path where the extracted data will be saved.

### Run the script:
  `python email_extract/main.py`

### The script will:
  - Connect to your Outlook inbox.
  - Filter emails within the specified date range.
  - Extract unique sender email addresses and display names.
  - Save the results to the specified CSV file.

### The output will include:
  - Total number of scanned emails.
  - Total number of unique sender addresses.
  - The path to the generated CSV file.