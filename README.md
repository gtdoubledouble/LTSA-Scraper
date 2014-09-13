This was an one-off script I wrote for a friend. She had a list of parcel IDs that can be queried on the LTSA (BC Land Title and Survey) to show information such as the owner's name and address. This script uses Selenium to initiate a Chrome instance as the driving web browser, login with your credentials (free account), then execute the queries one by one while storing the gathered owner names into an Excel spreadsheet.

The input needs to be an Excel file (pids.xlsx) with the first column filled with PIDs.
Some manual tweaking of the code needs to be done to match your own test case however. Contact me if you have any questions at garytse89@gmail.com
