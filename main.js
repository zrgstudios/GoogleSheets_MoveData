function setSpreadSheetByName(Name)
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var activesheet = sheet.getSheetByName(Name);
  return activesheet;
}

function checkCreateSpreadsheet(newdates)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetnames = [];
  for (var i = 0; i < sheets.length; i++)
  {
    sheetnames.push(sheets[i].getName())
  }

  var founddates = [];
  for (var j = 0; j < newdates.length; j++)
  {
    if (sheetnames.indexOf(newdates[j]) > -1)
    {
      //Logger.log(35 + " " + newdates[j])
    }
    else 
    {
      founddates.push(newdates[j]);
      ss.insertSheet(newdates[j])
    }
  }
  //Logger.log(39 + " " + sheetnames)
  //Logger.log(40 + " " + founddates.toString())
}

function getNumRows(sheet, range)
{
  Logger.log(37 + " " + range)

  // https://stackoverflow.com/questions/6882104/faster-way-to-find-the-first-empty-row-in-a-google-sheet-column
  var column = sheet.getRange(range)
  var values = column.getValues();
  var row = 0;
  while (values[row][0] != "")
  {
    row++;
  }
  return row;
}

function getDates()
{
  var ss = setSpreadSheetByName('InsertData')
  var range = "A2:A"
  var alldates = ss.getRange(range).getValues().filter(String);

  Logger.log(53 + " " + alldates.toString());
  Logger.log(54 + " " + alldates.length.toString());
  var uniquevalues = [];
  
  var rows = getNumRows(ss, range);
  Logger.log(69 + " " + rows);
  for (i=0; i < rows; i ++)
  {
    var currdate = alldates[i];
    var dateobj = new Date(currdate)
    var yearmonth = Utilities.formatDate(dateobj, Session.getScriptTimeZone(), "yyyy-MM")

    if (uniquevalues.indexOf(yearmonth) === -1)
    {
      uniquevalues.push(yearmonth)
    }
  }
  
  if (uniquevalues.length > 0)
  {
    Logger.log(73 + " " + uniquevalues.toString())
    checkCreateSpreadsheet(uniquevalues)
  } 
  //return uniquevalues;
}

function getRangeMinusHeaders(sheet, range, rangenotation)
{
  var startrow = range[0]
  var startcol = range[1]
  var maxrows = range[2]
  var numcols = range[3]

  var height = getNumRows(sheet, rangenotation)
  Logger.log(88 + " " + height)
  if (height === 0)
  {
    return null;
  }
  else
  {
    // row, column, num rows, num columns (goes out to e)
    return sheet.getRange(startrow, startcol, height, numcols);
  }
  //var width = range.getWidth();
  
}

function moveCells() 
{
  var sheet = setSpreadSheetByName('InsertData')

  var headerlessrange = getRangeMinusHeaders(sheet, [2, 1, 1000, 5], "A2:A");
  if (headerlessrange != null)
  {
    var headerlessdata = headerlessrange.getValues();
    var filteredheaderless = headerlessdata.filter(String);
    Logger.log(108 + " " + filteredheaderless);
    // Logger.log(Object.prototype.toString.call(filteredheaderless));

    for (var i = 0; i < filteredheaderless.length; i++)
    {
      // Determine which file the line goes into
      var dateobj = new Date(filteredheaderless[i][0]);
      var yearmonth = Utilities.formatDate(dateobj, Session.getScriptTimeZone(), "yyyy-MM");
      var yearmonthday = Utilities.formatDate(dateobj, Session.getScriptTimeZone(), "yyyy-MM-dd");

      var targetsheet = setSpreadSheetByName(yearmonth);
      // add one to start on next line, rows + new line
      var numrows = getNumRows(targetsheet, "A:A") + 1;

      Logger.log(115 + " " + yearmonth)

      var store = filteredheaderless[i][1];
      var category = filteredheaderless[i][2];
      var item = filteredheaderless[i][3];
      var cost = filteredheaderless[i][4];
      
      var data = [[yearmonthday, store, category, item, cost]]
      Logger.log(127 + " " + data)
      targetsheet.getRange(numrows, 1, 1, 5).setValues(data)
    }
    headerlessrange.clearContent();
  }

  var incomeheaderlessrange = getRangeMinusHeaders(sheet, [2, 7, 1000, 3], "G2:G");
  if (incomeheaderlessrange != null)
  {
    var incomeheaderlessdata = incomeheaderlessrange.getValues();
    var filtereddata = incomeheaderlessdata.filter(String);
    Logger.log(138 + " " + filtereddata);

    for (var i = 0; i < filtereddata.length; i++)
    {
      var dateobj = new Date(filtereddata[i][0]);
      var yearmonth = Utilities.formatDate(dateobj, Session.getScriptTimeZone(), "yyyy-MM");
      var yearmonthday = Utilities.formatDate(dateobj, Session.getScriptTimeZone(), "yyyy-MM-dd");
      var source = filtereddata[i][1]
      var income = filtereddata[i][2]

      var targetsheet = setSpreadSheetByName(yearmonth);
      // add one to start on next line, rows + new line
      var numrows = getNumRows(targetsheet, "G2:G") + 1;
      
      var data = [[yearmonthday, source, income]]
      Logger.log(127 + " " + numrows)
      targetsheet.getRange(numrows, 7, 1, 3).setValues(data)
    }
    incomeheaderlessrange.clearContent();
  }
}

function setdatavalidation()
{
  var sheet = setSpreadSheetByName('InsertData')
  var listrange = sheet.getRange("K1:K18")
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(listrange, true).setAllowInvalid(false).build();
  sheet.getRange("C2:C").setDataValidation(rule)
}

function main()
{
  getDates();
  moveCells();
  setdatavalidation();
}
