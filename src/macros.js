function resetfill() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A7').activate();
  spreadsheet.getActiveRangeList().setBackground(null);
};

function newsheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.insertSheet(3);
  spreadsheet.getActiveSheet().setName('new client sheet');
};

function nmRng() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A7:H13').activate();
  spreadsheet.setNamedRange('Fisher', spreadsheet.getRange('A7:H13'));
};

function clearBanding() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A7:F16').activate();
  var banding = spreadsheet.getRange('A7:F16').getBandings()[0];
  banding.remove();
};

function UpdateNmRange() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A7:H11').activate();
  
};

function copyRow() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(1, -1).activate();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow() - 1, 1, 1, sheet.getMaxColumns()).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};

function linkIt() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setCurrentCell(spreadsheet.getCurrentCell());
  
  spreadsheet.setCurrentCell(spreadsheet.getCurrentCell())
  .setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('Greent St')
  .setTextStyle(0, 9, SpreadsheetApp.newTextStyle()
  .setForegroundColor('#1155cc')
  .setUnderline(true)
  .build())
  .build());
};

function blah() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(8, -2).activate();
  spreadsheet.setCurrentCell(spreadsheet.getCurrentCell())
  .setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('Greent St')
  .setTextStyle(0, 9, SpreadsheetApp.newTextStyle()
  .setForegroundColor('#1155cc')
  .setUnderline(true)
  .build())
  .build());
};

function bkgrnd() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G4').activate();
  spreadsheet.getActiveRangeList().setBackground('#274e13')
  .setFontColor('#ffffff');
};

function act() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:F1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Aging Report'), true);
};

function updateNmRng() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(6, 0, 5, 8).activate();
  spreadsheet.getNamedRanges().forEach(function(namedRange) { if (namedRange.getName() == 'testst') { namedRange.setName('teststreet'); } });
};

function delrw() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('11:11').activate();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
};

function delsheets() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.deleteActiveSheet();
};

function delnmrng() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.removeNamedRange('WatersAve');
};

function unhidesheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getSheetByName('Criteria').showSheet()
  .activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Criteria'), true);
};

function mysort() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().sort(spreadsheet.getActiveRange().getColumn(), true);
};

function cellMerge() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:B1').activate()
  .mergeAcross();
};


function freezefool() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A5').activate();
  spreadsheet.getActiveSheet().setFrozenRows(5);
};

function CtrlDown() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:B1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
};

function UntitledMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D4').activate();
  spreadsheet.getActiveRangeList().setBackground('#783f04');
};

function delshiftup() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D7:E7').activate();
  spreadsheet.getRange('D7:E7').deleteCells(SpreadsheetApp.Dimension.ROWS);
};

function outborder() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2:C5').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
};

function res() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D11').activate();
  spreadsheet.getActiveSheet().setColumnWidth(8, 202);
  spreadsheet.getActiveSheet().setColumnWidth(8, 236);
  spreadsheet.getActiveSheet().setColumnWidth(8, 263);
};

function dt() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B8').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('M/d/yyyy');
};

function HomeIcon() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:F1').activate();
  var sheet = spreadsheet.getActiveSheet();
  var images = sheet.getImages();
  var image = images[images.length - 1];
  image.setHeight(27)
  .setWidth(24);
  images = sheet.getImages();
  image = images[images.length - 1];
  image.assignScript('linkToHome');
};

function delNmRng() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A8:H11').activate();
  spreadsheet.removeNamedRange('Template');
};

function button() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:F1').activate();
};

function clearRow() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('6:6').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, commentsOnly: true, skipFilteredRows: true});
};

function movesheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:F1').activate();
  spreadsheet.moveActiveSheet();
};

function protectSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var currentCell = spreadsheet.getCurrentCell().offset(1, 5);
  spreadsheet.getCurrentCell().offset(-6, -1, 8, 7).activate();
  spreadsheet.setCurrentCell(currentCell);
  var protection = spreadsheet.getActiveRange().protect();
  protection.setDescription('AgingReportLock')
  .setWarningOnly(true);
};

function unprotect() {
  var spreadsheet = SpreadsheetApp.getActive();
  var allProtections = spreadsheet.getActiveSheet().getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var matchingProtections = allProtections.filter(function(existingProtection) {
  return existingProtection.getRange().getA1Notation() == 'A1:G8';
  });
  var protection = matchingProtections[0];
  protection.remove();
};

function protectWholeSht() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  var protection = spreadsheet.getActiveSheet().protect();
  protection.setDescription('AgingSheetLock')
  .setWarningOnly(true);
};