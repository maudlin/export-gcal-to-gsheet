# Google Calendar Events Exporter

Google Calendar Events Exporter is a Google Apps Script which exports events from a user's Google Calendar to a Google Sheets document. 

## Functionality

The script fetches events from the previous month and writes them to the active Google Sheets document, starting from row 6. The data for each event includes:

1. The weekday of the event
2. The date and time of the event
3. The title of the event
4. The location of the event

Events containing the string "-project123" in their titles are excluded. The script also checks for the existence of the calendar, the validity of the date range, and the presence of events in the specified date range.

## Usage

1. Open a Google Sheets document where you want to deploy the script.
2. Click on `Extensions > Apps Script`.
3. Delete any code in the Apps Script editor and paste this script.
4. Save the project with a suitable name.
5. Run the function `export_gcal_to_gsheet` by selecting it from the select function drop-down and then click on the 'Run' button. The first time you run the script, it will ask for permissions to access your Google Calendar and the active Google Sheets document. Grant the necessary permissions.
6. Now, every time you want to update the sheet with the calendar events, you just need to run the `export_gcal_to_gsheet` function.

For automatic execution at certain intervals (like daily or weekly), you can set up a trigger. To do this, click on the clock icon ('Triggers') in the left sidebar of the Apps Script Editor. Then, click on '+ Add Trigger' at the bottom right of the page. Choose 'export_gcal_to_gsheet' for the function to run, set 'time-driven' for the deployment event type, and then specify the frequency (hourly, daily, etc.) and time of day as per your preference.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.
