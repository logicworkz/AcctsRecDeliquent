

function submitBug(){
 //learn how ot use the side panel. This is too short.
  Browser.inputBox(String.fromCharCode(128405) + " Submit Bug");
}

function onOpen() {
  const shtAging = ssActive.getSheetByName("Aging Report");
  const shtSettings = ssActive.getSheetByName('Settings');

  ui.createMenu('Aging Report')
    .addSubMenu(   
        ui.createMenu('Manage client delinquency')
        .addItem('New', 'newClient')
        .addItem('Rename', 'renameClient')
        .addItem('Delete', 'deleteClient'))
      .addSubMenu(
        ui.createMenu('Manage contact info')
        .addItem('Search', 'newClient')
        .addItem('New', 'newClient')
        .addItem('Rename', 'renameClient')
        .addItem('Delete', 'newClient'))
      .addSeparator()
        .addItem('Build Sheets','buildSheets')
        .addItem('Submit Bug','submitBug')
        .addToUi();
    
  //DSUM formulas arrive #VALUE upon open. Swiching to another sheet and switching back fixes the glitch.
  shtSettings
    .activate()
    .hideSheet();
  shtAging.activate();
}

function newClient() {

  if (!isSheet('Aging Report')){
    ui.alert("The Aging Report was not found. Please run the \'Build Sheets\' function first.") 
    return; 
  }

  const shtSettings = ssActive.getSheetByName('Settings');
  const shtAging = ssActive.getSheetByName('Aging Report');
  const shtTemplate = ssActive.getSheetByName('Template');
  
  const templateHeaderRow = '7';
  const templateBottom = '11';
  const headerRow = '5';

 //Prompt user for new client name
  const arInputBx = myPrompt("","new");
  if (arInputBx[0] == ui.Button.OK) {

    const strNewClientNm = arInputBx[1];
    if(!strNewClientNm.isBlank && strNewClientNm){ 
      
      //Duplicate template sheet for new clients
        unlockit(shtTemplate);
        
        shtTemplate.activate();
        const shtNewClient = ssActive.duplicateActiveSheet();
        shtNewClient
          .showSheet()
          .activate();
        shtTemplate.hideSheet();

        lockit(shtTemplate,'TemplateSheetLock');

      //Rename copied sheet & add report title to A1
        shtNewClient.getRange('B1').setValue(strNewClientNm);
        shtNewClient.setName(strNewClientNm);
        const strClientShtID = shtNewClient.getSheetId();

      //Move new client sheet after aging  report
        ssActive.moveActiveSheet(shtAging.getIndex());

      //Sanatize range name in preparation of Aging report DSUM formulas
        let strRangeNm = prepRangeNm(strNewClientNm,strClientShtID);
       
      //Name the range. append sheet id to range name for uniqueness       
        ssActive.setNamedRange(
          strRangeNm, shtNewClient.getRange('A' + templateHeaderRow + ':H' + templateBottom));
      
      //Load array with Criteria's named range objects for DSUM formulas
        const arCriteria = getDSUMcriteria(strRangeNm);

      //transpose array to paste it in horizontally
        const arDSUMformulas = transposeArray(arCriteria);

      //Insert new client row into aging report
        unlockit(shtAging);

        const cAging = shtAging.getRange('A' + startRow);
        if (!cAging.isBlank()){
          shtAging.insertRowsBefore(6, 1);
        }  

      //place DSUM formulas
        cAging.offset(0, 1, 1, 5).setFormulas(arDSUMformulas);

            //--format the client row - only needed if user deletes the first client created on a blank aging report 
            //Then creates another client and immediately deletes again lol 
              cAging.offset(0, 0, 1, 6)
                .setNumberFormat('"$"#,##0.00')
                .setHorizontalAlignment('center')
                .setFontFamily('Lexend')
                .setFontSize(11);

      //place grand total for newly added client on the same row
        shtAging.getRange("G" + startRow).setFormula('=SUM(B' + startRow + ':F' + startRow + ')');
       
          //--format the grand total cell for the client. only needed if user deletes the first client created on a blank aging report Then creates another client and immediately deletes again lol 
            const rngAnyHeading = shtAging.getRange('G' + startRow);
            rngAnyHeading.setBackground('#274e13')
              .setNumberFormat('"$"#,##0.00')
              .setFontColor('#ffffff')
              .setHorizontalAlignment('center')
              .setFontFamily('Lexend')
              .setFontSize(11);

      //update header formulas
        const rwAgingEnd = shtAging.getDataRange().getLastRow()+1;
        
       //-- update money totals formula
        let rngHeaderFormula = shtAging.getRange('B4');
        rngHeaderFormula.setFormula('=sum(B' + rwAgingStart + ':B' + rwAgingEnd + ')');

       //-- Autofill money totals formula
        const destRng = rngHeaderFormula.offset(0, 0, 1, 6);
        rngHeaderFormula.autoFill(destRng, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

       //-- update client count
        rngHeaderFormula = shtAging.getRange('A5');
        rngHeaderFormula.setFormula('= \"Client [\" & COUNTA(A' + rwAgingStart + ':A' + rwAgingEnd + ') & \"]\"');

      //reapply background color to grand total column
        shtAging.getRange('G4').setBackground('#274e13')
          .setFontColor('#ffffff');

      //place client name
        shtAging.getRange("A" + startRow).setValue(strNewClientNm);
        createBandings_Aging(shtAging,headerRow);

      //Create link to new client sheet & place client name
        link(shtAging.getRange("A" + startRow),strClientShtID);

        shtAging.autoResizeColumn(1);
        mySort(shtAging, startRow);
        shtAging.getRange("A8").activate();
        lockit(shtAging,'AgingSheetLock');

      //Add client sheet to Settings lookup
        unlockit(shtSettings);

        const intBtm = shtSettings.getRange("D1:E1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()+1;
        shtSettings.getRange("D" + intBtm).setValue(strNewClientNm);
        shtSettings.getRange("E" + intBtm).setValue(strClientShtID);

        lockit(shtSettings,'SettingsSheetLock');
      
      //DONE
        ssActive.toast("New client added!", "Ready...");
        shtNewClient.getRange("B8").activate(); //first row, invoice date

    } else {

        ui.alert("Please provide a client name before proceeding.");
        return;
    }  
  } 
}

function renameClient() {

  if (!isSheet('Aging Report')){
    ui.alert("The Aging Report was not found. Please run the \'Build Sheets\' function first.")
    return;
  }

  const shtSettings = ssActive.getSheetByName('Settings');
  const shtAging = ssActive.getSheetByName('Aging Report');
  const shtActive = ssActive.getActiveSheet();
  const cAging = shtAging.getActiveCell();
  let strAgingSheetCheck;
  let strPrevClientNm;

  if (shtActive.getSheetName() == 'Aging Report') {    
    if (cAging.getColumn() == 1 
        && cAging.getValue() != 'Aging Report' 
        && !cAging.isBlank()){

      strAgingSheetCheck = true;
      strPrevClientNm = cAging.getValue();

    } else {
      
      ui.alert("Please select a client name or sheet before proceeding.");
      return;
    }
  } else {
      //Must be client sheet, NOT Aging Report sheet
      strAgingSheetCheck = false;  
      strPrevClientNm = ssActive.getActiveSheet().getSheetName();
  }
 
   //Prompt user for new client name
  const arInputBx = myPrompt(strPrevClientNm,"rename");
  if (arInputBx[0] == ui.Button.OK) {

    const strNewClientNm = arInputBx[1];
    if(!strNewClientNm.isBlank && strNewClientNm){ 
  
      //update Sheet ID lookup with new client name
        const arRng = shtSettings.createTextFinder(strPrevClientNm).findAll();
        let rngFound = arRng[0];

        if(rngFound){
          unlockit(shtSettings);
          rngFound.setValue(strNewClientNm);
          lockit(shtSettings,"SettingsSheetLock");
        }else{
          ui.alert(strPrevClientNm + " was not found in lookup table. Exiting")
          return;
        }
        let strClientShtID = rngFound.offset(0,1).getValue();

      //rename sheet
        const shtClient = ssActive.getSheetById(strClientShtID);
        shtClient.setName(strNewClientNm);
          
      //update client report title
        shtClient.getRange("B1").setValue(strNewClientNm);

      //rename range
        const arNmRng = shtClient.getNamedRanges();
        let strRangeNm;

        if (arNmRng[0]){ 
          //--confirm it's a named range created by my code
          for (const nmRng of arNmRng){
            strRangeNm = nmRng.getName()                    
            if (strRangeNm.indexOf(strClientShtID) > 0){ 
              strRangeNm = prepRangeNm(strNewClientNm);
              
              //--DSUM formulas in Aging Report will auto-update! =)
                arNmRng[0].setName(strRangeNm);
            }
          }
        }
      //apply new name to aging report
        let strAging;
        if (strAgingSheetCheck){       
            cAging.setValue(strNewClientNm);

          //--re-establish link to client sheet
            strAging = shtAging.getActiveRange().getA1Notation();
            link(shtAging.getRange(strAging),strClientShtID);
            
        } else { //called from client sheet
            
            arRng = shtAging.createTextFinder(strPrevClientNm).findAll();
            rngFound = arRng[0];
            if (rngFound){
              rngFound.setValue(strNewClientNm);

          //--re-establish link to client sheet
              strAging = rngFound.getA1Notation();
              link(shtAging.getRange(strAging),strClientShtID); 
            }
        }
        mySort(shtAging,startRow);
        shtClient.activate();
        ssActive.toast(strPrevClientNm + " renamed to " + strNewClientNm, "Done" );
    }        
  }
}

function deleteClient() {

  if (!isSheet('Aging Report')){
    ui.alert("The Aging Report was not found. Please run the \'Build Sheets\' function first.")
    return;
  }
  
  const shtSettings = ssActive.getSheetByName('Settings');
  const shtAging = ssActive.getSheetByName('Aging Report');
  const shtActive = ssActive.getActiveSheet();
  const cAging = shtAging.getActiveCell();
  let strAgingSheetCheck;
  let strPrevClientNm;

  if (shtActive.getSheetName() == 'Aging Report') {    
    if (cAging.getColumn() == 1
        && cAging.getValue() != 'Aging Report' 
        && !cAging.isBlank()
        && cAging.getRow()>5){

      strAgingSheetCheck = true;
      strPrevClientNm = cAging.getValue();

    } else {
      
      ui.alert(String.fromCharCode(9785) + " Please select a client name before proceeding.");
      return;
    }
  } else {

    //Must be a client sheet, NOT Aging Report sheet
      strAgingSheetCheck = false;  
      strPrevClientNm = ssActive.getActiveSheet().getSheetName();
  }

  const arPrompt = ui.alert("Are you sure you want to delete " + strPrevClientNm + "?",ui.ButtonSet.YES_NO);
  if (arPrompt == ui.Button.YES) {

    //delete Sheet ID lookup entry
      let arRng = shtSettings.createTextFinder(strPrevClientNm).findAll();
      let rngFound = arRng[0];

      if(rngFound){
       //--create client sheet object before deleting lookup
        const strClientShtID = rngFound.offset(0,1).getValue();
        var shtClient = ssActive.getSheetById(strClientShtID);
        
        shtSettings.getRange('D7:' + 'E' + rngFound.getRow()).deleteCells(SpreadsheetApp.Dimension.ROWS);
      }else{
        ui.alert("'" + strPrevClientNm + "' was not found in the lookup table. Exiting.")
        return;
      }
    
    //Delete client entry from Aging Report
      if (strAgingSheetCheck){
        
        //--Initiated from Aging Report
        shtAging.deleteRows(cAging.getRow(), 1);
        ssActive.toast(strPrevClientNm + " deleted");

      } else {   

        //--Initiated from a client sheet  
        shtAging.activate();    
        arRng = shtAging.createTextFinder(strPrevClientNm).findAll();
        rngFound = arRng[0];
        if (rngFound){
          shtAging.deleteRows(rngFound.getRow(), 1);
        }
      }

    //delete all named ranges on the client sheet
      const arNmRng = shtClient.getNamedRanges();    
      const cnt_ar = arNmRng.length;

      for (i=0; i<=cnt_ar;i++){
        if (arNmRng[i]){
          ssActive.removeNamedRange(arNmRng[i].getName());
        }
      }
 
    //delete the client sheet
      ssActive.deleteSheet(shtClient);
      ssActive.toast(strPrevClientNm + " deleted");
  }
}

function buildSheets(){

  //Erase the sheets if they already exist
      const arShtNms = ['Settings','Template','Aging Report'];
      const arBool = arShtNms.map(str => isSheet(str)); 
      if (arBool.includes(true)){
        const arPrompt = ui.alert("Are you sure you want to proceed? This will delete the Aging Report & all supporting data! Client sheets created with this add-on will be preserved.",ui.ButtonSet.YES_NO);
        if (arPrompt == ui.Button.YES) {
          //const keepClients = true;

          try{
            arShtNms.map(str => ssActive.deleteSheet(ssActive.getSheetByName(str)));    
          }catch{
            for (const strShtNm of arShtNms){
              if (isSheet(strShtNm)){
                ssActive.deleteSheet(ssActive.getSheetByName(strShtNm));  
              }      
            }
          }
        } else {
            return;
        }
      }


  //----------------------------------------  
  //1st sheet is the hidden "Settings" sheet  
  //----------------------------------------
      ssActive.insertSheet(0);
      const shtSettings = ssActive.getActiveSheet().setName('Settings');
      
  //add DSUM criteria to Settings sheet
  
    //--heading
      let rngSettings = shtSettings.getRange('A1:B1');
      
      rngSettings.mergeAcross()
        .setValue('DSUM Criteria')
        .setHorizontalAlignment('center')
        .setFontWeight('bold')
        .setFontColor('#efefef')
        .setBackground('#3c78d8');
    
    //--values      
      const arA1loc = ["0",">=1",">30",">60",">90"];
      const arValues = ["","<=30","<=60","<=90",""];
      let ar2dCriteria = new Array(4);

      for(i=0;i<=8;i++){
        ar2dCriteria[i] = new Array(2);
        ar2dCriteria[i][0] = arA1loc[i];
        ar2dCriteria[i][1] = arValues[i];
      }
      shtSettings.getRange(2,1,ar2dCriteria.length,ar2dCriteria[0].length).setValues(ar2dCriteria);
    
    //-- DSUM headings
      for (x=0; x<11; x+=2){
        if(x!=0){
          shtSettings.insertRowsBefore(x, 1);
          let rng = shtSettings.getRange("A" + x)
          rng.setValue("Past Due");
          
          let destRng = rng.offset(0, 0, 1, 2);
          rng.autoFill(destRng, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
        }
      }

    //--named ranges
      const arRangeNm = ["Crit_0","Crit_1_30","Crit_31_60","Crit_61_90","Crit_90"];
      let intTopofRng = 2;
      let intBotofRng = 3;
      
      for(let cnt = 0;cnt < 5;cnt++){
        ssActive.setNamedRange(
          arRangeNm[cnt], shtSettings.getRange('A' + intTopofRng + ':B' + intBotofRng));

        intTopofRng = ++intBotofRng; 
        intBotofRng = intTopofRng + 1; 
      }

    //--create lookup of sheet Name --> sheetIDs  
      rngSettings = shtSettings.getRange("D1:E1")
      rngSettings.mergeAcross()
        .setValue('Sheet ID Lookups')
        .setHorizontalAlignment('center')
        .setFontWeight('bold')
        .setFontColor('#efefef')
        .setBackground('#783f04');

      rngSettings = shtSettings.getRange("D2");
        rngSettings.setValue('Sheet Name')
          .setHorizontalAlignment('center');

      rngSettings = shtSettings.getRange("E2");
        rngSettings.setValue('Sheet ID')
            .setHorizontalAlignment('center');

    //-- add newly created settings sheet to lookup
      shtSettings.getRange('D3').setValue(shtSettings.getSheetName());
      shtSettings.getRange('E3').setValue(shtSettings.getSheetId());

    //--protect sheet
      lockit(shtSettings,'SettingsSheetLock')
      shtSettings.hideSheet();

  //-----------------------------
  //2nd sheet is the Aging Report
  //-----------------------------

      ssActive.insertSheet(1);
      const shtAging = ssActive.getActiveSheet();
      const strAgingSheetID = shtAging.getSheetId();
      shtAging.setName('Aging Report');

    //-- Add new Aging Report sheet to Sheet IDs lookup
      rngSettings = shtSettings.getRange('D1:E1');
      let intBtm = rngSettings.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()+1;
      shtSettings.getRange("D" + intBtm).setValue('Aging Report');
      shtSettings.getRange("E" + intBtm).setValue(strAgingSheetID);

    //-- Report title
      rngSettings = shtAging.getRange('A1:B1')
      rngSettings.mergeAcross()
        .setHorizontalAlignment('center')
        .setValue('Aging Report')
        .setFontFamily('Lexend')
        .setFontSize(24)
        .setBackground('#073763')
        .setFontColor('#ffffff');
      ssActive.getActiveSheet().setColumnWidth(2, 156);

    //-- create money totals formulas header & formatting   
        const rwAgingEnd = 7;
        let rngHeaderFormula = shtAging.getRange('B4');
        rngHeaderFormula.setFormula('=sum(B' + rwAgingStart + ':B' + rwAgingEnd + ')');

        let destRng = rngHeaderFormula.offset(0, 0, 1, 6);
        rngHeaderFormula.autoFill(destRng, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
        rngHeaderFormula = shtAging.getRange('B4:G4');
        rngHeaderFormula.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
          .setHorizontalAlignment('center')
          .setNumberFormat('"$"#,##0.00')
          .setBackground('#d9d9d9')
          .setFontWeight('bold')
          .setFontFamily('Lexend')
          .setFontSize(13);
        shtAging.getRange('G4')
          .setBackground('#274e13')
          .setFontColor('#ffffff');

      //-- Column titles
        let rngAnyHeading = shtAging.getRange('A5:F5');
        rngAnyHeading.setBackground('#38761d')
          .setFontColor('#ffffff')
          .setFontFamily('Lexend')
          .setFontSize(13)
          .setFontWeight('bold')
          .setFontColor('#ffffff')
          .setHorizontalAlignment('center');

        rngAnyHeading = shtAging.getRange('G5');
        rngAnyHeading.setBackground('#274e13')
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setHorizontalAlignment('center')
          .setFontFamily('Lexend')
          .setFontSize(13);

        //-- Grand total cell
        rngAnyHeading = shtAging.getRange('G6');
        rngAnyHeading.setBackground('#274e13')
          .setNumberFormat('"$"#,##0.00')
          .setFontColor('#ffffff')
          .setHorizontalAlignment('center')
          .setFontFamily('Lexend')
          .setFontSize(11);

        shtAging.getRange('A5').setFormula('= "Client [" & COUNTA(A6:A7) & "]"');
        shtAging.getRange('B5').setValue('Current');
        shtAging.getRange('C5').setValue('1-30 DAYS');
        shtAging.getRange('D5').setValue('31-60 DAYS');
        shtAging.getRange('E5').setValue('61-90 DAYS');
        shtAging.getRange('F5').setValue('>90 DAYS');
        shtAging.getRange('G5').setValue('Grand Totals');

      //--font of first 2 rows
        rngAnyHeading = shtAging.getRange('A6:G7');
        rngAnyHeading.setHorizontalAlignment('center')
          .setFontFamily('Lexend')
          .setNumberFormat('"$"#,##0.00')
          .setFontSize(11);
      
      //--column widths
        let intMaxCol = shtAging.getMaxColumns();
        rngAnyHeading = shtAging.getRange(1, 1, shtAging.getMaxRows(), intMaxCol);
        rngAnyHeading.activate();
        shtAging.setColumnWidths(1,intMaxCol,120)
          .setFrozenRows(5);
      
        shtAging.getRange('A6').activate();

      //--protect sheet
        lockit(shtAging,'AgingSheetLock')
    
  //-------------------------
  //3rd sheet is the Template
  //-------------------------

      ssActive.insertSheet(0);
      const shtTemplate = ssActive.getActiveSheet();
      shtTemplate.setName('Template');
      
    //--Add new template sheet to Sheet IDs lookup
      rngSettings = shtSettings.getRange('D1:E1');
      intBtm = rngSettings.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()+1;
      shtSettings.getRange("D" + intBtm).setValue('Template');
      shtSettings.getRange("E" + intBtm).setValue(shtAging.getSheetId());

    //--Report title
      rngAnyHeading = shtTemplate.getRange('B1:F1');
      rngAnyHeading.mergeAcross()
        .setHorizontalAlignment('center')
        .setValue('Template')
        .setFontFamily('Lexend')
        .setFontSize(24)
        .setBackground('#073763')
        .setFontColor('#ffffff');
      ssActive.getActiveSheet().setColumnWidth(2, 156);

    //Home "button"  
      rngAnyHeading = shtTemplate.getRange("A1")  
      rngAnyHeading.setFormula('=CHAR(127968)')
        .setFontSize(24)
        .setBackground('#073763')
        .setHorizontalAlignment('left');

      link(rngAnyHeading, strAgingSheetID);

    //--Summary section
      const arSummaryLabels =[['TOTAL INVOICE'],['TOTAL PAID'],['TOTAL DUE'],['AVG DELINQUENCY']]; //setValues expects a 2D array

      for (i=2; i<6; i++){
        rngAnyHeading = shtTemplate.getRange('A' + i + ':B' + i);  
        rngAnyHeading.mergeAcross()
          .setHorizontalAlignment('right')
          .setFontFamily('Lexend')
          .setFontSize(11)
          .setFontWeight('bold')
          .setBackground('#0b5394')
          .setFontColor('#ffffff');
      }
      shtTemplate.getRange(2,1,arSummaryLabels.length,1).setValues(arSummaryLabels);

        rngAnyHeading = shtTemplate.getRange('C2:C5');  
        rngAnyHeading.setHorizontalAlignment('center')
          .setFontFamily('Lexend')
          .setFontSize(11)
          .setFontWeight('bold')
          .setNumberFormat('"$"#,##0.00')
          .setBackground('#0b5394')
          .setFontColor('#ffffff');

        shtTemplate.getRange("C2").setFormula('=SUM(E8:E11)');
        shtTemplate.getRange("C3").setFormula('=SUM(D8:D11)');
        shtTemplate.getRange("C4").setFormula('=C2-C3');
        shtTemplate.getRange("C5").setFormula('=IF(F8<>"",ROUNDUP(AVERAGE(F8:F11)),"-")');

        rngAnyHeading = shtTemplate.getRange('A2:C5');  
        rngAnyHeading.setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);

    //--Headings
        const arHeadings = [['Invoice No.'],['Invoice Date'],['Invoice Amt.'],['Amount Paid'],['Balance'],['Past Due'],['Terms'],['Item/Service Description']];
      
        const ar = transpose(arHeadings);
        rngAnyHeading = shtTemplate.getRange(7,1,1,ar[0].length);
        rngAnyHeading.setValues(ar)
          .setFontFamily('Lexend')
          .setFontSize(13)
          .setFontWeight('bold')
          .setHorizontalAlignment('center')
          .setBackground('#b6d7a8')
          .setFontColor('#356854'); 

    //--column widths
        intMaxCol = shtTemplate.getMaxColumns();
        rngAnyHeading = shtTemplate.getRange(1, 1, shtTemplate.getMaxRows(), intMaxCol);
        rngAnyHeading.activate();
        shtTemplate.setColumnWidths(1,intMaxCol,140)
          .setFrozenRows(7);
        ssActive.getActiveSheet().setColumnWidth(8, 263);

    //--create bandings + other formats    
        let rngMain = shtTemplate.getRange('A8:H11');
        ssActive.setNamedRange(
          'Template', rngMain);  
        
        rngMain.setFontFamily('Lexend')
          .setFontSize(10)
          .setHorizontalAlignment('center');

        shtTemplate.getRange('C8:D11').setNumberFormat('"$"#,##0.00');
        shtTemplate.getRange('B8:B11').setNumberFormat('M/d/yyyy');
        
        createBandings_Clients(shtTemplate,11);
        ssActive.removeNamedRange('Template');

    //--formulas
        rngHeaderFormula = shtTemplate.getRange('E8');
        rngHeaderFormula.setFormula('=IF(C8<>"",C8-D8,"")');
        destRng = rngHeaderFormula.offset(0, 0, 4, 1);
        rngHeaderFormula.autoFill(destRng, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

        rngHeaderFormula = shtTemplate.getRange('F8');
        rngHeaderFormula.setFormula('=IF(B8<>"",IF(TODAY()-B8 > 0, TODAY()-B8,0),"")');
        destRng = rngHeaderFormula.offset(0, 0, 4, 1);
        rngHeaderFormula.autoFill(destRng, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
      
    //--format balance column    
        rngAnyHeading = shtTemplate.getRange('E8:E11');  
        rngAnyHeading.setHorizontalAlignment('center')
          .setFontFamily('Lexend')
          .setFontSize(10)
          .setNumberFormat('"$"#,##0.00')
          .setBackground('#38761d')
          .setFontColor('#ffffff');

        shtTemplate.getRange('A8').activate();
    //--protect sheet
        lockit(shtTemplate,'TemplateSheetLock')
        shtTemplate.hideSheet();

  //--------------------------------------------
  //Add any existing clients to the Aging Report
  //--------------------------------------------
      keepClients();
}