# Attendance Tracker
This script processes Zoom attendance data by comparing a list of expected participants against a list of actual participants. It loads two CSV files: one containing the expected participants with their unofficial and official names (`expected_participants.csv`), and the other containing the actual participants (`participants_1234567890.csv`). It exports an Excel file containing the attendance information, including the names of the attendees and their attendance status (Present, Absent, or Present (Unrecognized Name)).

## Requirements
Python 3.7 or higher  
Pandas library   
openpyxl library 
- If openpyxl is not installed, the script will generate a CSV file instead of an Excel file.
- You can install openpyxl by running `conda install -c anaconda openpyxl` or `pip install openpyxl`)

## Usage
1. Export the participants list from [Zoom](https://zoom.us/account/my/report) ensuring that you check both "Export with meeting data" and "Show unique users". Please refer to this guide for more information: https://www.eduhk.hk/ocio/content/faq-how-retrieve-attendance-list-zoom-meeting.
2. Create a CSV file named `expected_participants.csv` with the following columns: "Name (Original Name)" and "Official Name" and "Coach Name". List all the participants that are expected to attend the event in this file. Make sure that the "Name (Original Name)" column matches the name column in the exported CSV file from Zoom.
3. Run the script by executing `python attendance_tracker.py`.
4. The script will generate an Excel file named `attendance.xlsx` with the attendance information, including the names of the attendees and their attendance status (Present, Absent, or Present (Unrecognized Name)).

## Example of Excel Output <a name="screenshot"></a>
<p align="center">
  <img src="extra\Excel_File.png" />
</p>


[//]: <> (## Note: The script assumes that the exported participants list from Zoom is saved as `participants_xxxxxxx.csv` in the same directory as the script file.)

