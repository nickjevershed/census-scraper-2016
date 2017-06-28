# Census scraper
The first results for the 2016 Australian census will be available on June 27. However, the first official tabulated data waon't be available until later (TableBuilder on July 4 and DataPacks on July 12).

If you want tabulated data on June 27, this script converts the community profile time series spreadsheets for each sa2 into a single sqlite database for analysis.

Or you can download the resulting database [here](https://drive.google.com/file/d/0B1aVLtLn2O4-ZDV1RHVZWnlyY0E/view?usp=sharing).

Feel free to use this, but if you do and you find it useful please let me know.

**Note**: The items that are usually represented as lists, such as language other than english spoken at home, ancestry etc. have been converted to JSON and then written into the SQLlite db as strings. These will need to be converted back into an object, or excluded if you want a CSV. 
