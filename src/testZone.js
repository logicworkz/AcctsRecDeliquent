




function showSidebar() {
  return SpreadsheetApp.getUi()
  .showSidebar(HtmlService
               .createHtmlOutput('<button onclick="google.script.run.myFunction()">Run</button>'));
}

function test (){
  const htmlOutput = HtmlService.createHtmlOutput( '<p>A change of speed, a change of style...</p>', ).setWidth(250).setHeight(300); 
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'My add-on');
  
  //var ui = SpreadsheetApp.getUi();
  //ui.showModalDialog(htmlOutput);

  //Browser.inputBox("test");
}

/*function onSelectionChange(e) {
  var shtActiveNm = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName();
  var rngClicked = e.range;

  if (shtActiveNm == 'Aging Report'){
  
    if (rngClicked.getColumn() === 1 
      && rngClicked.getNumColumns() === 1 
      && rngClicked.getRow() >= 6
      && rngClicked.getCell(1, 1).getValue() != ''
      && rngClicked.getCell(1, 1).getValue() != 'Invoice No.') {

      var strClickedClient = SpreadsheetApp.getActiveSheet().getActiveRange().getValue();
        
      rngClicked.setBackground('yellow');
      findClientID(strClickedClient, rngClicked);
      rngClicked.setBackground(null);
    }
  }  
}
*/
//----------------------------------------------
//---------- UNHIDE SELECTED CLIENT SHEET ------
//----------------------------------------------

function findClientID(strClientNm, rngCell){
//function findClientID(){

  var ssActive = SpreadsheetApp.getActiveSpreadsheet();
  var shtSheet_IDs = ssActive.getSheetByName('Sheet_IDs');
  var arRng = shtSheet_IDs.createTextFinder(strClientNm).findAll();
  var arFound = arRng[0];

  if(!arFound){
    rngCell.setBackground('red');//not working
    return;
    }else {
      ssActive.toast('Switching to ' + strClientNm,);
    }
  debugger

  var shtID = arFound.offset(0,1).getValue();

  ssActive.getRange('A1').activate();//prevents the event trigger from going off when user goes back to aging sheet 
  ssActive.getSheetById(shtID).showSheet();
  ssActive.getSheetById(shtID).activate();
}


function transpose2DArray(a)
{
  return Array(a).map(r=> r.map(cell => cell[0]))
}

function transpose(a)
{
  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
}

function pasteArrayOfFormulas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("A1:C3"); // Adjust the range as needed
  var formulas = [
    ['=SUM(B1:B10)', '=AVERAGE(C1:C10)', '=MAX(D1:D10)'],
    ['=MIN(B1:B10)', '=COUNTA(C1:C10)', '=PRODUCT(D1:D10)'],
    ['=STDEV(B1:B10)', '=VAR(C1:C10)', '=MEDIAN(D1:D10)']
  ];
  
  // Set formulas to the range
  range.setFormulas(formulas);
}

function showPrompt2() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

// Display a sidebar with custom HtmlService content.
/*const htmlOutput = HtmlService
                       .createHtmlOutput(
                           '<p>A change of speed, a change of style...</p>',
                           )
                       .setTitle('Client history summary');
SpreadsheetApp.getUi().showSidebar(htmlOutput);
debugger*/

// Add an item to the built-in Extensions menu, under a sub-menu whose name is set
// automatically.
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Showing', 'showSidebar')
      .addToUi();
}
  var result = ui.prompt(
    "Let's get to know each other!",
    "Please enter your name:",
    ui.ButtonSet.OK_CANCEL,
  );

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert("Your name is " + text + ".");
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert("I didn't get your name.");
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert("You closed the dialog.");
  }
}


function showAlert() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
    "Please confirm",
    "Are you sure you want to continue?",
    ui.ButtonSet.YES_NO,
  );

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert("Confirmation received.");
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert("Permission denied.");
  }
}
function clearBackground(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rng = sh.getDataRange();
  rng.setBackground(null);
  
}


