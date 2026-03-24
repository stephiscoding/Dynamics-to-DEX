# Excel to XML FINAL

Takes an Excel file containing IPPS sessions, and converts it into an XML file for upload to DSS DEX. It will automatically interpret the data directly from the CRM export spreadsheet – no need to transfer to an intermediate spreadsheet. It automatically selects the correct event code (stored in event_codes.json), and if it doesn’t recognise a certain event, it will prompt the user for a new code that it will use for future events.

## Usage:

1. Install the requirements listed in requirements.txt:
```
pip3 install -r requirements.txt
```
2. Run the Python script:
```
python main.py NAME_OF_EXCEL_SPREADSHEET
```
This will create `[MONTH] output.xml`, which is the file to upload to DEX.