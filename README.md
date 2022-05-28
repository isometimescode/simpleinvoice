# ⚠️ This repository is no longer maintained ⚠️

## Simple Invoicing with Google Scripts

## Background
I was looking for something very simple that would copy spreadsheet rows to a document, all within Google Apps/Google Drive. There are plenty of mail merge and invoicing projects listed in the Google add-ons to accomplish some of this, but a lot of them require the use of third party sites or have more customization or features that I don't need. They are worth looking into though, as this requires some extra steps to use, like creating a developer API key. 

## Setup

You'll at least need a spreadsheet with data in it, and a document with template variables to fill in. 

### Google Picker

In order to use the [Google Picker](https://developers.google.com/apps-script/guides/dialogs#file-open_dialogs) for a file dialog, you first need to get a Developer's API Key. After following the directions in the Google example, I was getting a lot of "invalid key" errors because I didn't wait long enough for it to be activated or sync'd with Google. Its not instantaneous. 

### Spreadsheet

I name the spreadsheet after the company I'm invoicing. That is used to fill in the _Company Name_ on the template. The number of columns doesn't matter, but the code does assume the first row is a header row, and that there is at least a date column. Extra columns can exist, and if they don't match a template variable the data will just be skipped. My columns are:
```
Date, Description, Hours, Rate, Collected, Day Total
```

![Spreadsheet Screenshot](/doc/spreadsheet.png)

### Doc Template

The document template will be cloned, so you can use it as many times as needed and across different spreadsheets. All template variables look like `<<VariableName>>`. Values can be replaced more than once, and there are several Global values available:
* `<<CompanyName>>`: this is the spreadsheet's name unless defined otherwise in the code
* `<<CurrentDate>>`: today's date
* `<<TotalDue>>`: total dollar sum of all invoice rows added
* `<<InvoiceNumber>>`: today's date like '20160414'

Invoice item rows are in a table that must have a header and footer row. The header's first column must be called _date_ because that's how its found by the code. See `Main.gs::findInvoiceTable()` function if you want to change it.

The code will copy the row with all the `<<ItemXName>>` values filled in. The _XName_ in this case is the column name from your spreadsheet. Dates will be formatted and currency will be formatted. 

![Invoice Template Screenshot](/doc/invoice-template.png)

### Google Scripts

You can add these scripts like any other Google script; by going to the spreadsheet and clicking *Tools > Script Editor...*. The name only matters for the HTML file, because that is referenced in the code. The rest can be named anything or combined. 

You can edit the top of `Main.gs` if you want to change date or currency formats. You can also change the `SpreadsheetName` variable if you don't want to use it for the company name. Don't forget to fill out your API key in the `Picker.html` file. 

You'll also have to setup a trigger for the menu item. In the script editor, I have *Resources > All your triggers > On open > From spreadsheet*. This just adds the "Invoicing" item in the menu bar:

![Menu Item](/doc/spreadsheet-menu.png)

## Final Result

![Final Doc](/doc/invoice-final.png)
