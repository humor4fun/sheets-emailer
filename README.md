# sheets-emailer
Automatic email sender for Google Sheets

# Usage
1. Import these two CSV files into a new spreadsheet. Alternative, create your own sheets that match these descriptions.
2. Make sure that the sheets (tabs) are named correctly (Recipients, Templates)
3. Create a new App Script and copy in the script.gs code
4. Create an html file for the script and copy in the StatusPage.html code
5. Run the `onOpen` function, or reload the spreadsheet
6. From the menu, execute the `Dry Run` menu item to test the code. This will execute the entire function but will not send any emails


# Notes
FedEx: Use this formula to create links off of the tracknig ID alone
```
=concat("https://www.fedex.com/apps/fedextrack/?tracknumbers=",$C2)
```

# Recipients Sheet
![Screenshot](/examples/Sheet-Recipients.png)

# Template Sheet
![Screenshot](/examples/Sheet-Template.png)
