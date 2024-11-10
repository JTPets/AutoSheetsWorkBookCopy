Google Sheets Automation Script for Data Sync and Notifications
This Google Apps Script automates the process of syncing data between two Google Sheets and sends email notifications upon completion or failure. It also allows for daily automated execution via a scheduled trigger.

Features
Data Sync: Copies data from a source sheet to a target sheet, maintaining formatting and sorting.
Email Notifications: Sends a success or failure email after each sync, including details like the number of rows/columns updated and links to the source and target sheets.
Dynamic Configuration: Configure source and target sheets, email notifications, and daily triggers directly from the Google Sheets menu.
Daily Automation: Set up a daily trigger to run the sync at a specified time.
Installation
Open your Google Sheet.
Go to Extensions > Apps Script.
Copy and paste the script into the editor.
Save the project and refresh your Google Sheet.
Setup
1. Set Up Menu
After installation, a new menu called Custom Tools will appear in your Google Sheet. This menu includes the following options:

Refresh Data: Manually run the data sync.
Set Notification Email: Configure the email address for notifications.
Set Source and Target Sheets: Set the source and target sheet IDs and names.
Set Daily Timer: Configure a daily trigger to automate the sync.
2. Configure Source and Target Sheets
Select Set Source and Target Sheets from the Custom Tools menu.
Follow the prompts to input:
Source spreadsheet ID
Source sheet name
Target spreadsheet ID
Target sheet name
3. Set Notification Email
Select Set Notification Email from the Custom Tools menu.
Enter the email address where notifications will be sent.
4. Set Daily Trigger
Select Set Daily Timer from the Custom Tools menu.
Enter the time in 24-hour format (e.g., 08 for 8 AM, 14 for 2 PM) to set the daily sync trigger.
Usage
Manual Sync
Use the Refresh Data option to manually sync data from the source sheet to the target sheet.
Automated Sync
Once the daily trigger is set, the script will automatically sync data at the configured time each day.
Notification Emails
Success Email:

Total rows and columns updated.
Links to the source and target sheets.
Failure Email:

Error message describing the issue.
Script Functions
Key Functions
copyData(): Copies data from the source sheet to the target sheet.
setEmailAddress(): Configures the email address for notifications.
setupSheetConfiguration(): Sets the source and target sheets.
setupDailyTrigger(): Prompts the user to set a daily trigger time.
createDailyTrigger(hour): Creates a daily trigger for the specified time.
deleteExistingTriggers(): Removes existing triggers for this script.
notifySuccess(): Sends a success email with details.
notifyFailure(): Sends a failure email with an error message.
Dynamic Configuration
All configuration values (sheet IDs, sheet names, email) are stored in PropertiesService for persistent use across script executions.
Example Emails
Success:

less
Copy code
Subject: Data Refresh Successful

The data refresh was completed successfully.

Details:
- Total rows updated: 100
- Total columns updated: 10
- Source Sheet: https://docs.google.com/spreadsheets/d/1ABCDEF1234567890XYZ
- Target Sheet: https://docs.google.com/spreadsheets/d/1GHIJKL9876543210UVW
Failure:

vbnet
Copy code
Subject: Data Refresh Failed

The data refresh failed with the following error:

<Error message>
Limitations
Ensure the user account running the script has access to both the source and target sheets.
Triggers are subject to Google Apps Script's daily execution limits.
License
This project is licensed under the MIT License. See the LICENSE file for details.
