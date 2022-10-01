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

function getNumRows(sheet)
{
  var range = sheet.getDataRange();
  var values = range.getValues();
  var row = 0;
  for (var row=0; row < values.length; row++)
  {
    if (!values[row].join("")) break;
  }
  return row;
}

function getDates()
{
  var ss = setSpreadSheetByName('InsertData')
  var alldates = ss.getRange(1, 1, 1000).getValues();
  var uniquevalues = [];
  
  var rows = getNumRows(ss);
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
  //Logger.log(uniquevalues.toString())
  checkCreateSpreadsheet(uniquevalues)
  return uniquevalues;
}

function getRangeMinusHeaders(range)
{
  var height = range.getHeight();
  if (height == 1)
  {
    return null;
  }
  var width = range.getWidth();
  var sheet = range.getSheet();
  // row, column, num rows, num columns (goes out to e)
  return sheet.getRange(2, 1, height-1, 5);
}

function getFirstEmptyWholeRow(name)
{
  var ss = setSpreadSheetByName(name);
  var rows = getNumRows(ss);
  return (rows+1);
}

function MoveCells() 
{
  var basedata = setSpreadSheetByName('InsertData').getDataRange();

  var headerlessrange = getRangeMinusHeaders(basedata)
  var headerlessdata = headerlessrange.getValues();
  Logger.log(headerlessdata.toString());
  // Logger.log(Object.prototype.toString.call(headerlessdata));
  for (var i = 0; i < headerlessdata.length; i++)
  {
    // Determine which file the line goes into
    var dateobj = new Date(headerlessdata[i][0]);
    var yearmonth = Utilities.formatDate(dateobj, Session.getScriptTimeZone(), "yyyy-MM");
    Logger.log(yearmonth)

    var yearmonthday = Utilities.formatDate(dateobj, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var store = headerlessdata[i][1];
    var category = headerlessdata[i][2];
    var item = headerlessdata[i][3];
    var cost = headerlessdata[i][4];

    var targetsheet = setSpreadSheetByName(yearmonth);
    var numrows = getNumRows(targetsheet) + 1;
    
    var data = [[yearmonthday, store, category, item, cost]]
    Logger.log(numrows)
    targetsheet.getRange(numrows, 1, 1, 5).setValues(data)
  }
  headerlessrange.clearContent();
}
