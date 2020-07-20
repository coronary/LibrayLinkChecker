# Library Database Title Validator

Node script made to double check the validity of item links found in an excel spreadsheet. Uses puppeteer because the website being checked is a single page application so we have to wait until the content is loaded before pulling the html.

Must add in retrieved title column in data and fill with zeroes for script to edit the cells.

Script is highly specific to my use case. If you need to modify it change the column letters that match your spreadsheet on lines 25-27.
