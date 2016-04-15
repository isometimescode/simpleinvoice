/**
 * Given a spreadsheet with formatted data,
 * take each row and fill it out into an invoice if its not collected.
 *
 * License: MIT
 *
 * Copyright 2016 Toni Wells <isometimescode.com>
 */



/**
 * Set this to what you want your dates to look like, i.e. 01-Jan-2016
 * https://developers.google.com/apps-script/reference/utilities/utilities#formatdatedate-timezone-format
 */
var DateFormat = "dd-MMM-YYYY";
var TimeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

/**
 * Set this to how you want your currency formatted, i.e. $100.50
 * https://developers.google.com/apps-script/reference/utilities/utilities#formatstringtemplate-args
 */
var CurrencyFormat = "\$%.2f";

/**
 * This is used to fill in "CompanyName" on the invoice, assumed name of spreadsheet
 */
var SpreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();

/**
 * Today :)
 */
var CurrentDate = Utilities.formatDate(new Date(), TimeZone, DateFormat);




/**
 * Given a document ID, copy the file and call parseRowData to
 * fill in the template variables.
 *
 * @param {string} id The Google Drive document ID to clone.
 *
 * @return {object} An object with the url and title of the document.
 */
function cloneDocument(id) {
  // id = 'uncomment and add doc id for testing';

  // before cloning this template, let's make sure it is valid
  // will throw an error if not
  findInvoiceTable(DocumentApp.openById(id).getBody());

  // make a copy of the template file named with today's date
  var name = 'Invoice '+SpreadsheetName+' '+CurrentDate;
  var newfile = DriveApp.getFileById(id).makeCopy(name);

  // copy all data into new document clone
  parseTemplate(newfile.getId());

  // return with info about the new invoice for the panel display
  return { url:newfile.getUrl(), title:name };
}

/**
 * Search a document body for a table that looks like it should have invoice data.
 *
 * @throws {string} If invoice row table not found in body
 *
 * @param  {object} body DocumentApp.body object
 *
 * @return {object}      DocumentApp.body.table object
 */
function findInvoiceTable(body) {
  var tables = body.getTables();
  var invoicetable = undefined;

  for(var i = 0; i < tables.length; i++) {
    var text = tables[i].getRow(0).getCell(0).getText();

    // looking for the table that has our invoice row header
    // TODO could ask for this on the form and match words that way
    if(text.toLowerCase() == "date") {
      invoicetable = tables[i];
      break;
    }
  }

  // If the table wasn't found, we probably are using an inappropriate template
  if(invoicetable === undefined) {
    throw("No table template found in document for invoice data");
  }

  return invoicetable;
}

/**
 * Modify the provided Google Document by filling in all variables
 * with names like <<VARIABLE>> using related information from this spreadsheet.
 *
 * @param {int} fileid The Google Drive document ID to modify
 */
function parseTemplate(fileid) {
  var head = DocumentApp.openById(fileid).getHeader();
  var body = DocumentApp.openById(fileid).getBody();
  var invoiceNumber = Utilities.formatDate(new Date(), TimeZone, 'YYYYMMdd');

  // Replace global text in header and body
  body.replaceText('<<CurrentDate>>', CurrentDate);
  body.replaceText('<<CompanyName>>', SpreadsheetName);
  body.replaceText('<<InvoiceNumber>>', invoiceNumber);
  head.replaceText('<<CurrentDate>>', CurrentDate);
  head.replaceText('<<CompanyName>>', SpreadsheetName);
  head.replaceText('<<InvoiceNumber>>', invoiceNumber);

  // fills out as many rows as needed and then returns the sum total
  var totaldue = parseTableData(findInvoiceTable(body));

  // Final total due
  body.replaceText('<<TotalDue>>', Utilities.formatString(CurrencyFormat, totaldue));
}

/**
 * Using the current active spreadsheet and a document table object,
 * fill in all <<VARIABLE>> values with data from every row in the spreadsheet.
 *
 * @param {object} doctable DocumentApp.body.table object
 *
 * @return {int} summary of all total rows applied, possibly 0 if no totals found
 */
function parseTableData(doctable) {
  var as = SpreadsheetApp.getActiveSheet();

  var rawdata = as.getDataRange().getValues();
  var header = rawdata[0];
  var totalcol = -1;
  var totaldue = 0;

  // looking for the column that has the row total
  // may not exist, in which case totaldue == 0
  for(var i = header.length-1; i >= 0; i--) {
    if(header[i].match(/total/i)) {
      totalcol = i;
      break;
    }
  }

  // loop through every row except the header
  for(var i = 1; i < rawdata.length; i++) {
    if(totalcol >= 0) {
      // if the total due is not a positive number, than we can assume its not invoiceable
      // and move on to the next row
      if(rawdata[i][totalcol] <= 0) {
        continue;
      }

      totaldue += rawdata[i][totalcol];
    }

    // loop through every column in this row and fill in variables in the template
    // if a given column does not have a matching template variable, nothing will change in the doc
    var rowcopy = doctable.getRow(1).copy();
    for(var x = 0; x < rawdata[i].length; x++) {
      // only letters/numbers in name
      var name = header[x].replace(/[^A-Za-z0-9]/g,'');
      rowcopy.replaceText('<<Item'+name+'>>', tryFormat(name,rawdata[i][x]));
    }

    // the last row is a summary row, so add this copy row in before that (i.e. second to last)
    // TODO could ask the user if the last row is a summary row
    doctable.insertTableRow(doctable.getNumRows()-1,rowcopy);
  }

  // no longer need the template row, so we can delete it
  doctable.getRow(1).removeFromParent();

  return totaldue;
}

/**
 * Conditional formatting based on a guess from the provided name. Otherwise
 * returns data unchanged.
 *
 * @param {string} name Name of column or variable to check, i.e. "Date"
 * @param {string} value Value to format
 *
 * @return {string} Value formatted as date or currency, or unchanged
 */
function tryFormat(name,value) {
  if(name.match(/date/i)) {
    return Utilities.formatDate(value, TimeZone, DateFormat);
  }

  if(name.match(/total/i) || name.match(/rate/i)) {
    return Utilities.formatString(CurrencyFormat, value);
  }

  return value;
}