// TradePhantomsPro Notifier
// Version 1.0
// Copyright Mark Alway, phippster@gmail.com

// Default Values
var CHECK_FREQUENCY = 5; // Frequency in minutes for checks
var LOGIC_OPERATOR = "AND"; // Use "AND" or "OR" to combine results
var COLUMN_VALUE_PAIRS = [
]; // Array of column-value pairs to check
   // Format as { "column name" : ["operator", "value"] }
   // Supported operators: =, >, <, >=, <=, ~

var copySheetName = "Data Copy";
var filteredSheetName = "Filtered Trades";
var tempSheetName = "Temp Data";
var displayTempDebugSheet = false;

// Global Variables
var SHEET_URL = "https://docs.google.com/spreadsheets/d/1c9TjTHlt24NkvbnzHblbXjYg9EQyuJXS2sm1HRXwXGo/edit?gid=0#gid=0"; // Replace with the URL of the read-only sheet
var SHEET_NAME = "Current Month"; // Replace with the name of the sheet you want to copy
var TIMEZONE = "America/New_York"; // Adjust to your timezone

// Column name to index mapping
var COLUMN_MAP = {
  "date": 1,
  "asset class": 2,
  "ticker/ pair": 3,
  "direction": 4,
  "type of trade": 5,
  "entry": 6,
  "current price (delayed)": 7,
  "stop loss": 8,
  "target 1": 9,
  "target 2": 10,
  "target 3": 11,
  "strike": 12,
  "expiration": 13,
  "notes / trade management": 14,
  "results": 15,
  "% return option": 16,
  "% w/l account": 17,
  "rowid": 18,
  "rowchecksum": 19
};

// Function to run when the spreadsheet is opened
function onOpen() {
  initializeProperties();
  initializeSheets();
  createMenu();
  runOnce();
  createOnOpenTrigger();
  activateSheet(filteredSheetName);
}

// Script Properties Functions ------------------------------------------

// Function to initialize script properties
function initializeProperties() {
  var scriptProperties = PropertiesService.getScriptProperties();

  if (!scriptProperties.getProperty("LOGIC_OPERATOR")) {
    setLogicOperator(LOGIC_OPERATOR); // Default value
  }

  if (!scriptProperties.getProperty("CHECK_FREQUENCY")) {
    setCheckFrequency(CHECK_FREQUENCY); // Default frequency
  }

  if (!scriptProperties.getProperty("COLUMN_VALUE_PAIRS")) {
    setColmnValuePairs(COLUMN_VALUE_PAIRS); // Default frequency
  }
}

function activateSheet(sheetName){
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.activate();
}

// Function to initialize necessary sheets
function initializeSheets(){
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Check if Data Copy sheet exists.  If not create it.
  var copyDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheetName);
  if (!copyDataSheet) {
    currentSpreadsheet.insertSheet(copySheetName);
  }

  // Check if Filtered Trades sheet exists, if not create it.
  var filteredSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(filteredSheetName);
  if(!filteredSheet){
    // Check to see if the Sheet1 name exists and rename it.
    var defaultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
    if(defaultSheet){
      // Go ahead and rename it.
      defaultSheet.clear();
      defaultSheet.setName(filteredSheetName);
    }else{
      // Otherwise create it.
      currentSpreadsheet.insertSheet(filteredSheetName);
    }
  }


  // Check to see if TempTrades sheet exists, if not create it.
  if(displayTempDebugSheet){
    var tempSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tempSheetName);
    if (!tempSheet) {
      currentSpreadsheet.insertSheet(tempSheetName);
    }
  }

}

// Set functions for ScriptProperties
function setLogicOperator(val){
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("LOGIC_OPERATOR", val);
}
function setCheckFrequency(val){
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("CHECK_FREQUENCY", val); 
}
function setColmnValuePairs(val){
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("COLUMN_VALUE_PAIRS", JSON.stringify(val));
}

// Function to retrieve properties
function getLogicOperator() {
  return PropertiesService.getScriptProperties().getProperty("LOGIC_OPERATOR") || "AND";
}

function getCheckFrequency() {
  return parseInt(PropertiesService.getScriptProperties().getProperty("CHECK_FREQUENCY")) || 1;
}

function getColumnValuePairs() {
  var pairs = PropertiesService.getScriptProperties().getProperty("COLUMN_VALUE_PAIRS");
  return pairs ? JSON.parse(pairs) : [];
}

function showAllScriptProperties() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var allProperties = scriptProperties.getProperties(); // Retrieve all properties

  if (Object.keys(allProperties).length === 0) {
    SpreadsheetApp.getUi().alert("No script properties found.");
    return;
  }

  // Format properties for display
  var formattedProperties = "Current Script Properties:\n";
  for (var key in allProperties) {
    formattedProperties += `${key}: ${allProperties[key]}\n`;
  }

  // Show alert with formatted properties
  SpreadsheetApp.getUi().alert(formattedProperties);
}

// Accessory Functions not called directly ----------------------------------------

function formatDate(date) {
  if (!(date instanceof Date)) return date;
  
  // Convert to specified timezone
  var formattedDate = Utilities.formatDate(date, TIMEZONE, "MM/dd/yyyy");
  return formattedDate;
}


function createHeaderRow(sheet) {
  // Create array of headers from COLUMN_MAP
  var headers = new Array(Object.keys(COLUMN_MAP).length);
  Object.entries(COLUMN_MAP).forEach(([key, index]) => {
    headers[index - 1] = key; // -1 because COLUMN_MAP uses 1-based indices
  });
  
  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Optional: Format header row to make it stand out
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#E8E8E8');
}

function compareWithOperator(cellValue, operator, targetValue) {
  // Convert to strings for initial processing
  cellValue = String(cellValue).toLowerCase();
  targetValue = String(targetValue).toLowerCase();
  
  // For numeric comparisons, convert both values to numbers
  if (operator !== "=" && operator !== "~") {
    // Blank string "" converts to 0.  We want to ignore empty strings.
    if(cellValue === ""){
      return false;
    }
    cellValue = Number(cellValue);
    targetValue = Number(targetValue);
    
    // Return false if either value isn't a valid number
    if (isNaN(cellValue) || isNaN(targetValue)) {
      return false;
    }
  }
  
  switch(operator) {
    case "=":
      return cellValue === targetValue;
    case ">":
      return cellValue > targetValue;
    case "<":
      return cellValue < targetValue;
    case ">=":
      return cellValue >= targetValue;
    case "<=":
      return cellValue <= targetValue;
    case "~":
      return cellValue.includes(targetValue);
    default:
      throw new Error(`Unsupported operator: ${operator}`);
  }
}

function runOnce(){
  copyFromTradePhantomsSheet();
  filterResults();
}

function filterResults() {
  var values = getSheetData(copySheetName);

  var foundRowsObj = {}; // For completely new rows
  var foundRowsArray = [];
  var logicOperator = getLogicOperator();  // cache this so we don't need to read it every iteration in loop below.
  var rowsUpdated = [];  // track which rows were updated for notificadtions;
  var rowsAdded = [];  // Track which rows were added for notifications


  // Convert keys in COLUMN_VALUE_PAIRS to lowercase
  var pairs = getColumnValuePairs()
  var lowerCaseColumnValuePairs = pairs.map(pair => {
    var key = Object.keys(pair)[0];
    var [operator, value] = pair[key];
    return {[key.toLowerCase()]: [operator, value]};
  });
  

  // Process matching rows from Data Copy
  for (var i = 0; i < values.length; i++) {
    var match = logicOperator === "AND";

    // Check if row matches criteria
    for (var j = 0; j < lowerCaseColumnValuePairs.length; j++) {
      var column = Object.keys(lowerCaseColumnValuePairs[j])[0];
      var [operator, value] = lowerCaseColumnValuePairs[j][column];
      var columnIndex = COLUMN_MAP[Object.keys(COLUMN_MAP).find(col => col.toLowerCase() === column)] - 1;

      var cellValue = values[i][columnIndex];
      if (cellValue instanceof Date) {
        cellValue = formatDate(cell);
      }

      if (logicOperator === "OR") {
        if (compareWithOperator(cellValue, operator, value)) {
          match = true;
          break;
        }
      } else if (logicOperator === "AND") {
        if (!compareWithOperator(cellValue, operator, value)) {
          match = false;
          break;
        }
      }
    }
    
    if (match) {
        var currentRow = values[i];
        // For total uniqueness we also use the row number from the copy data file 
        var rowID = calculateUniqueRowID(currentRow,i);
        currentRow.push(rowID)
        var checksum = calculateRowChecksum(currentRow);
        currentRow.push(checksum);

        foundRowsObj[rowID] = {"checksum": checksum, "rowData": currentRow};
        foundRowsArray.push(currentRow);
    }
  }

  // The latest filtered date is in the foundRowsObj.
  // We now want to read the data from our prior check and compare that data to the new data.
  // 1.  Any row that is in the old data but not the new we will remove.
  // 2.  Any row that is in the new data but not in the old we will add to the old data
  // 3.  Any row in the old that has changed in the new we will update in the old with the latest values.
  // 

  //variable for old Data
  var oldDataObj = {};

  // Get the old data.
  var oldDataArray = getSheetData(filteredSheetName);
  // Convert into a nice JS Object to work with.
  // Row index 17 is the RowID, row indox 18 is the checksum
  oldDataArray.forEach((row) => {
    oldDataObj[row[17]] = {"checksum": row[18], "rowData": row};
  });

  // 1. Any removals?
  Object.keys(oldDataObj).forEach((rowID) => {
    // Check if it exists in the foundDataObj
    if(! (rowID in foundRowsObj)){
      // Row no longer in data, delete from oldDataObj
      delete oldDataObj[rowID];
    }
  });

  // 2. Any Additions or updates?
  Object.keys(foundRowsObj).forEach((rowID) => {
    // If key is not in olddataObj then add it.
    if(! (rowID in oldDataObj)){
      oldDataObj[rowID] = foundRowsObj[rowID];
      rowsAdded.push(foundRowsObj[rowID]['rowData']);
    }else{
      // 3. Update section, key exists so verify the checksum.
      if(oldDataObj[rowID]["checksum"] != foundRowsObj[rowID]["checksum"]){
        // If not the same checksum copy over the latest data.
        oldDataObj[rowID] = foundRowsObj[rowID];
        rowsUpdated.push(foundRowsObj[rowID]["rowData"]);
      }
    }
  });

  // At this point oldDataObj should have all the correct data and we can display it.
  displayDataOnSheet(filteredSheetName, oldDataObj);

  // Debug by showing on temp sheet
  if(displayTempDebugSheet){
    if (foundRowsArray.length > 0) {
      displayDataOnSheet(tempSheetName, foundRowsObj);
    }
  }

  // Create detailed email content only if there are updates or new rows
  if (rowsAdded.length > 0 || rowsUpdated.length > 0) {
    sendNotificationEmail(rowsAdded, rowsUpdated);
  }
}

function displayDataOnSheet(sheetName, dataObj){
  // Convert the object to an array to be used for setValues();
  var dataArray = [];
  Object.keys(dataObj).forEach((rowID) => {
    dataArray.push(dataObj[rowID]['rowData']);
  });

  // Sort by date
  var sortedDataArray = dataArray.sort(
    function (a,b){
      return a[0] - b[0];
    }
  );


  // Get the sheet to display on.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("The '" + sheetName + "' tab not found in the current spreadsheet");
  }

  // Clear the sheet
  sheet.clear();

  // Create header row if temp data is empty
  if (sheet.getLastRow() === 0) {
    createHeaderRow(sheet);
  }

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, sortedDataArray.length, sortedDataArray[0].length).setValues(sortedDataArray);

  // Remove the data validation on header
  sheet.getRange(1,1,1,1).setDataValidation(null);

}

// Accessory function to get the data in a sheet.
function getSheetData(sheetName){
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var copiedDataSheet = currentSpreadsheet.getSheetByName(sheetName);
  var range = copiedDataSheet.getDataRange();  // This is the Data Copy sheet.
  var values = range.getValues();

  // Remove the header row if it's a header row.
  if(values[0][0].includes('date') || values[0][0].includes('Date')){
    values[0][0].includes('date')
  }

  return values;
}

// Accessory function that creates a unique row ID based on the first 5 coluumn values of a row.
function calculateUniqueRowID(rowArray, originalIndex) {
  // Get the first 5 values from the row array
  const valuesToHash = [...rowArray.slice(0, 6),originalIndex].join('|'); // Join with a delimiter for uniqueness
  
  // Create a SHA-256 hash of the concatenated string
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, valuesToHash);
  
  // Convert the byte array to a hexadecimal string
  return hash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

// Accessory function that calculates a row checksum to see if the row has changed or not.
// Key to this is to ignore the current price column (index 6) in the calulcation if it is a number.
function calculateRowChecksum(rowArray) {
  // checksum for entire row
  var valuesToHash = rowArray.join('|'); // Join with a delimiter for uniqueness 

  // Check the current price column to see if it's a number (that's always changing) and 
  // would throw our "changed" off with just the price action.
  if(!isNaN(rowArray[6])){ // If it's a number checksum without it.
    // Make a copy of array so slice doesn't damage it.
    var copiedRowArray = [...rowArray];
    copiedRowArray.splice(6,1); // remove the current price column.
    valuesToHash = copiedRowArray.join('|');  // Change the values to hash
  }

  
  // Create a SHA-256 hash of the concatenated string
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, valuesToHash);
  
  // Convert the byte array to a hexadecimal string
  return hash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}


function sendNotificationEmail(rowsAdded, rowsUpdated){
  var emailBody = "";
    var headers = new Array(Object.keys(COLUMN_MAP).length);
    Object.entries(COLUMN_MAP).forEach(([key, index]) => {
      headers[index - 1] = key; // -1 because COLUMN_MAP uses 1-based indices
    });
  
    if (rowsAdded.length > 0) {
      if (emailBody) emailBody += "\n";
      emailBody += `${rowsAdded.length} new trade${rowsAdded.length > 1 ? 's were' : ' was'} added:\n\n`;
      rowsAdded.forEach(function(row) {
        headers.forEach((header, index) => {
          emailBody += `${header}: ${row[index]}\n`;
        });
        emailBody += "-".repeat(40) + "\n\n";
      });
    }

    if (rowsUpdated.length > 0) {
      if (emailBody) emailBody += "\n";
      emailBody += `${rowsUpdated.length} trade${rowsUpdated.length > 1 ? 's were' : ' was'} updated:\n\n`;
      rowsUpdated.forEach(function(row) {
        headers.forEach((header, index) => {
          emailBody += `${header}: ${row[index]}\n`;
        });
        emailBody += "-".repeat(40) + "\n\n";
      });
    }

   // Send email with detailed information
    var email = Session.getActiveUser().getEmail();
    MailApp.sendEmail({
      to: email,
      subject: `Trade Updates in ${filteredSheetName} (${rowsUpdated.length} updated, ${rowsAdded.length} new)`,
      body: emailBody
    }); 
    
}
  
function copyFromTradePhantomsSheet() {
  var sourceSpreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
  var sourceSheet = sourceSpreadsheet.getSheetByName(SHEET_NAME);
  
  // Get the current spreadsheet
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var copySheetName = "Data Copy";
  var destinationSheet = currentSpreadsheet.getSheetByName(copySheetName);
  
  // If the copy sheet doesn't exist, create it
  if (!destinationSheet) {
    throw new Error("The '" + copySheetName + "' tab not found in the current spreadsheet");
  }
  
  // Copy the data and format dates
  var range = sourceSheet.getRange(9,1, sourceSheet.getLastRow() - 8, sourceSheet.getLastColumn());
  var values = range.getValues();
  
  // Format dates in the copied data
  values = values.map((row, rowIndex) => {
    return row.map((cell, colIndex) => {
      // Check if this column is the date column
      if (colIndex === COLUMN_MAP["date"] - 1 && cell instanceof Date) {
        return formatDate(cell);
      }
      // Check if this column is the expiration column
      if (colIndex === COLUMN_MAP["expiration"] - 1 && cell instanceof Date) {
        return formatDate(cell);
      }
      return cell;
    });
  });
  
  destinationSheet.clear();
  destinationSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  
}


// UI Menu Handling functions ----------------------------------------------


function createMenu() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('TradePhantomsPro Notifier');
  
  // Add the "Run Notifier Once" option
  menu.addItem('Search Criteria Builder', 'showSearchBuilder')
    .addSeparator()
    .addItem('Run Notifier Once', 'runOnce')
    .addSeparator();
  
  var notifierActive = isBackgroundNotifierRunning();
  
  if (notifierActive) {
    var freq = getCheckFrequency()
    menu.addItem('Notifier running every ' + freq + ` minute${freq > 1 ? 's' : ''}`, 'doNothing')
        .addItem('Stop Notifier', 'stopBackgroundNotifier');
  } else {
    menu.addItem('Start Background Notifier', 'createTrigger')
        .addItem('Stop Notifier', 'stopBackgroundNotifier');
  }
  menu.addItem("Set Frequency",'promptFrequency')
    .addSeparator();

  var debugMenu = ui.createMenu('Debug');
  
  // Add other tools
  debugMenu.addItem('Show all Properties','showAllScriptProperties')
    .addItem('Reset all Properties','resetAllScriptProperties')
    .addItem('Show Triggers','displayTriggers')
    .addItem('Reset Data','resetSheets')
    .addItem('Delete All Triggers','deleteAllTriggers');

  menu.addItem('Configuration Help', 'showHelpDialog')
    .addItem('About This Script', 'showAboutDialog')
    .addSubMenu(debugMenu);
  
  menu.addToUi();
}

function resetSheets(){
  var filteredsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(filteredSheetName);
  filteredsheet.clear();
  var copysheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheetName);
  copysheet.clear();
}

function promptFrequency() {
  var ui = SpreadsheetApp.getUi();
  
  // Prompt the user for a number (check frequency)
  var response = ui.prompt('Enter the frequency in minutes (1,5, 10, 15, or 30):');
  
  // Get the user's input
  var userInput = response.getResponseText();
  
  // Validate if the input is a number
  var frequency = parseInt(userInput, 10);
  if (!isNaN(frequency) && (frequency == 1 || frequency == 5 || frequency == 10 || frequency == 15 || frequency == 30)) {
    // Save the valid frequency value to script properties
    setCheckFrequency(frequency.toString());
    // If trigger exists, re-create it with the new frequency.  Otherwise don't do anything.
    if(isBackgroundNotifierRunning()){
      // Cancel all Triggers
      stopBackgroundNotifier();
      // Restart Triggers
      createTrigger();
    }  
  //ui.alert('Check frequency set to: ' + frequency + ' minutes.');
  } else {
    // If input is invalid, prompt again
    ui.alert('Please enter a valid frequency of 1, 5, 10, 15, or 30.');
  }
}

function resetAllScriptProperties() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
  SpreadsheetApp.getUi().alert('All script properties have been deleted.');
}


function doNothing() {
  return true;
}

function createTrigger() {
  // Run Right now.
  runOnce();

  //  And then also set the trigger
  ScriptApp.newTrigger("runOnce")
    .timeBased()
    .everyMinutes(getCheckFrequency())
    .create();
  
  // Recreate the menu to show updated status
  createMenu();
}

// Create the onRun
function createOnOpenTrigger(){
  // Check to see if the onOpen trigger is already created.
  var triggers = ScriptApp.getProjectTriggers();
  var found = false;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() === ScriptApp.EventType.ON_OPEN && triggers[i].getHandlerFunction() == "onOpen"){
      found = true;
    }
  }

  if(!found){

    var sheet = SpreadsheetApp.getActive();

    //  And then also set the trigger
    ScriptApp.newTrigger("onOpen")
    .forSpreadsheet(sheet).onOpen()
    .create();
  }
 
}

function stopBackgroundNotifier() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if(isBackgroundNotifierRunning()){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Recreate the menu to show updated status
  createMenu();
}

function isBackgroundNotifierRunning(){
   // Check if trigger exists and create appropriate menu items
   var triggers = ScriptApp.getProjectTriggers();
   for (var i = 0; i < triggers.length; i++) {
    if(triggers[i].getEventType() === ScriptApp.EventType.CLOCK && triggers[i].getHandlerFunction() == "runOnce"){
      return true;
    }
  }
  return false;
}

function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // Recreate the menu to show updated status
  createMenu();
}

function displayTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerInfo = 'Current Triggers:\n\n';
  
  if (triggers.length === 0) {
    triggerInfo += 'No triggers are currently active.\n';
  } else {
    triggers.forEach(function(trigger) {
      triggerInfo += 'Function: ' + trigger.getHandlerFunction() + '\n';
      triggerInfo += 'Event Type: ' + trigger.getEventType() + '\n';
      if(trigger.getEventType() === ScriptApp.EventType.CLOCK){
        triggerInfo += 'Frequency: ' + getCheckFrequency() + ` minute${getCheckFrequency() > 1 ? 's' : ''}` + '\n';
      }
      triggerInfo += '\n';
    });
  }
  
  SpreadsheetApp.getUi().alert(triggerInfo);
}


function showHelpDialog() {
  var htmlTemplate = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 20px;
            line-height: 1.6;
          }
          h2 {
            color: #1155cc;
            border-bottom: 1px solid #ddd;
            padding-bottom: 5px;
          }
          .example {
            background: #f8f9fa;
            padding: 10px;
            border-left: 3px solid #1155cc;
            margin: 10px 0;
          }
        </style>
      </head>
      <body>
        <h2>Search Criteria Help</h2>
        <p>This tool helps you track specific trades by setting up rules about what to look for in your spreadsheet. 
        You can set up multiple rules using a user-friendly HTML form.</p>
        
        <h2>How to Set Up Rules</h2>
        <p>To set up your search criteria:</p>
        <ol>
          <li>Click on the "Search Criteria Builder" option in the menu.</li>
          <li>A dialog will open where you can add criteria for your search.</li>
          <li>For each criterion, select the column you want to check, choose an operator (e.g., Equals, Contains, Greater Than), and enter the value you're looking for.</li>
          <li>You can add multiple criteria by clicking the "Add Criteria" button.</li>
          <li>Once you've set your criteria, click "Save Criteria" to apply your changes.</li>
        </ol>

        <h2>Available Operators</h2>
        <ul>
          <li><strong>=</strong> Matches exactly</li>
          <li><strong>~</strong> Contains this text</li>
          <li><strong>></strong> Greater than</li>
          <li><strong><</strong> Less than</li>
          <li><strong>>=</strong> Greater than or equal to</li>
          <li><strong><=</strong> Less than or equal to</li>
        </ul>

        <h2>Managing the Background Process</h2>
        <p>You can start or stop the background process that checks for updates while you're away:</p>
        <ol>
          <li>In the menu, you will see options to "Start Background Notifier" or "Stop Notifier".</li>
          <li>Click "Start Background Notifier" to begin the process. The notifier will run at the frequency you set.</li>
          <li>To stop the notifier, simply click "Stop Notifier".</li>
        </ol>

        <h2>Important Notes</h2>
        <ul>
          <li>Column names aren't case sensitive (e.g., "Type of Trade" and "type of trade" work the same).</li>
          <li>Text searches (using = or ~) aren't case sensitive.</li>
          <li>When using >, <, >=, or <=, ensure you're using them with number values.</li>
          <li>The logic operator setting ("AND" or "OR") determines how multiple rules work together:
            <ul>
              <li>"AND" means all rules must match.</li>
              <li>"OR" means any rule can match.</li>
            </ul>
          </li>
        </ul>
      </body>
    </html>
  `)
  .setWidth(600)
  .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlTemplate, 'Search Criteria Configuration Help');
}

function showSearchBuilder() {
  var pairs = getColumnValuePairs();
  var currentLogicOperator = getLogicOperator();
  var htmlTemplate = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 20px;
            line-height: 1.6;
          }
          .criteria-row {
            margin-bottom: 15px;
            padding: 10px;
            background: #f8f9fa;
            border-radius: 4px;
          }
          select, input {
            margin: 5px;
            padding: 5px;
            width: 200px;
          }
          button {
            margin: 5px;
            padding: 8px 15px;
            background: #1155cc;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
          }
          button:hover {
            background: #0d47a1;
          }
          .remove-btn {
            background: #dc3545;
            float: right;
          }
          .remove-btn:hover {
            background: #c82333;
          }
          #criteria-container {
            margin-bottom: 20px;
          }
          #save-button-container {
            display: flex;
            justify-content: center;
            margin-top: 20px;
          }
        </style>
      </head>
      <body>
        <h3>Search Criteria Builder</h3>
        <div id="criteria-container"></div>
        <button onclick="addCriteria()">Add Criteria</button>
        <h3>Logic Operator</h3>
        <p>If you have more than one criteria above, do you want all criteria to match or any one of them?
        <div>
          <select id="logicOperator">
            <option value="AND" ${currentLogicOperator === "AND" ? "selected" : ""}>All criteria must match</option>
            <option value="OR" ${currentLogicOperator === "OR" ? "selected" : ""}>Any one criteria</option>
          </select>
        </div>
        <div id="save-button-container">
          <button onclick="saveCriteria()">Save Criteria</button>
        </div>
        

        <script>
          // Get column names from the COLUMN_MAP
          const columnNames = ${JSON.stringify(Object.keys(COLUMN_MAP))};
          
          // Get existing criteria
          const existingCriteria = ${JSON.stringify(pairs || [])};
          
          function createCriteriaRow(existingCriterion) {
            const div = document.createElement('div');
            div.className = 'criteria-row';
            
            // Column dropdown
            const colSelect = document.createElement('select');
            columnNames.forEach(col => {
              const option = document.createElement('option');
              option.value = col;
              option.text = col;
              if (existingCriterion && 
                  col.trim().toLowerCase() === Object.keys(existingCriterion)[0].replace(/"/g, '').trim().toLowerCase()) {
                option.selected = true;
              }
              colSelect.appendChild(option);
            });
            
            // Operator dropdown
            const opSelect = document.createElement('select');
            const operators = [
              {value: '=', text: 'Equals'},
              {value: '~', text: 'Contains'},
              {value: '>', text: 'Greater Than'},
              {value: '<', text: 'Less Than'},
              {value: '>=', text: 'Greater Than or Equal'},
              {value: '<=', text: 'Less Than or Equal'}
            ];
            operators.forEach(op => {
              const option = document.createElement('option');
              option.value = op.value;
              option.text = op.text;
              if (existingCriterion && op.value === existingCriterion[Object.keys(existingCriterion)[0]][0]) {
                option.selected = true;
              }
              opSelect.appendChild(option);
            });
            
            // Value input
            const valueInput = document.createElement('input');
            valueInput.type = 'text';
            valueInput.placeholder = 'Enter value';
            if (existingCriterion) {
              valueInput.value = existingCriterion[Object.keys(existingCriterion)[0]][1];
            }
            
            // Remove button
            const removeBtn = document.createElement('button');
            removeBtn.textContent = 'Remove';
            removeBtn.className = 'remove-btn';
            removeBtn.onclick = function() {
              div.remove();
            };
            
            div.appendChild(colSelect);
            div.appendChild(opSelect);
            div.appendChild(valueInput);
            div.appendChild(removeBtn);
            
            return div;
          }
          
          function addCriteria(existingCriterion) {
            const container = document.getElementById('criteria-container');
            container.appendChild(createCriteriaRow(existingCriterion));
          }
          
          function saveCriteria() {
            const rows = document.getElementsByClassName('criteria-row');
            const criteria = [];
            
            Array.from(rows).forEach(row => {
              const selects = row.querySelectorAll('select');
              const input = row.querySelector('input');
              if (input.value.trim()) {
                const key = selects[0].value;
                criteria.push({
                  [key]: [selects[1].value, input.value.trim()]
                });
              }
            });

            google.script.run
              .withSuccessHandler(() => {
                google.script.host.close();
              })
              .saveCriteriaToScript(criteria);

            // Get the selected logic operator from the dropdown
            const logicOperator = document.getElementById('logicOperator').value;

            google.script.run.setLogicOperator(logicOperator);
          }
          

          
          // Initialize with existing criteria or add empty row if none exist
          if (existingCriteria && existingCriteria.length > 0) {
            existingCriteria.forEach(criterion => addCriteria(criterion));
          } else {
            addCriteria();
          }
        </script>
        
      </body>
    </html>
  `)
  .setWidth(600)
  .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlTemplate, 'Search Criteria Builder');
}

function saveCriteriaToScript(criteriaObj) {
  var currentCriteria = getColumnValuePairs();
  // Check if the new criteria is different from the current criteria
  if (JSON.stringify(criteriaObj) !== JSON.stringify(currentCriteria)) {
    // If the criteria has changed, delete all rows in Filtered Trades

    var filteredsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(filteredSheetName);

    if(displayTempDebugSheet){
      var tempsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tempSheetName);
      if (tempsheet) {
        tempsheet.clear();  // Clears all rows and columns
      }
    }

    if (filteredsheet) {
      filteredsheet.clear();  // Clears all rows and columns
    }
  
    try {
      setColmnValuePairs(criteriaObj); 
      runOnce();
      return true;
    } catch(error) {
      throw new Error('Failed to save criteria: ' + error.toString());
    }
  }
}

function showAboutDialog() {
  // Define the version number from the comments at the top of the file
  const versionNumber = "1.0"; // Update this if the version changes

  var htmlTemplate = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 20px;
            line-height: 1.6;
          }
          h2 {
            color: #1155cc;
          }
          p {
            margin: 10px 0;
          }
        </style>
      </head>
      <body>
        <p><strong>Version:</strong> ${versionNumber}</p>
        <p><strong>Author:</strong> Mark Alway</p>
        <p><strong>Contact:</strong> <a href="mailto:phippster@gmail.com">phippster@gmail.com</a></p>
        <p>This script helps you track specific trades by setting up rules about what to look for in your spreadsheet.</p>
      </body>
    </html>
  `)
  .setWidth(400)
  .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlTemplate, 'About This Script');
}

