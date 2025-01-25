function keepClients(){

  //Preserve any sheets created with my function
    const shtSettings = ssActive.getSheetByName('Settings');
    const shtAging = ssActive.getSheetByName('Aging Report');
    const arAllSheets = ssActive.getSheets();
    let strRangeNm;

    //--outer loop iterates all tabs in the file
    for (const sht of arAllSheets){
      
      const strClientNm = sht.getSheetName();
      const arNmRng = sht.getNamedRanges();

      if (arNmRng[0]){

        const strSheetID = sht.getSheetId();  
        
        for (const nmRng of arNmRng){
          strRangeNm = nmRng.getName()
          //--then i've found one of my sheets!
          if (strRangeNm.indexOf(strSheetID) > 0){  

            //--Insert row for existing client into aging report
            const headerRow = '5'; //\\\\gotta make this a public value!\\\\\

            const cAging = shtAging.getRange('A' + startRow);
            if (!cAging.isBlank()){
              shtAging.insertRowsBefore(6, 1);
            }    

            //--Load array with Criteria's named range objects for DSUM formulas
            const arCriteria = getDSUMcriteria(strRangeNm);

            //--transpose array to paste it in horizontally
            const arDSUMformulas = transposeArray(arCriteria);

            //--place DSUM formulas
            cAging.offset(0, 1, 1, 5).setFormulas(arDSUMformulas);

            //--place grand total for newly added client on the same row
            shtAging.getRange("G" + startRow).setFormula('=SUM(B' + startRow + ':F' + startRow + ')');

            //update header formulas
              const rwAgingStart = startRow;
              const rwAgingEnd = shtAging.getDataRange().getLastRow()+1;
              
              //-- update money totals formula
                let rngHeaderFormula = shtAging.getRange('B4');
                rngHeaderFormula.setFormula('=sum(B' + rwAgingStart + ':B' + rwAgingEnd + ')');

              //-- Autofill money totals formula
                let destRng = rngHeaderFormula.offset(0, 0, 1, 6);
                rngHeaderFormula.autoFill(destRng, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);  

              //-- update client count
                rngHeaderFormula = shtAging.getRange('A5');
                rngHeaderFormula.setFormula('= \"Client [\" & COUNTA(A' + rwAgingStart + ':A' + rwAgingEnd + ') & \"]\"');  

            //place client name
              shtAging.getRange("A" + startRow).setValue(strClientNm);
              createBandings_Aging(shtAging,headerRow);

            //Create link to new client sheet
              link(shtAging.getRange("A" + startRow),strSheetID);
              shtAging.autoResizeColumn(1);
              mySort(shtAging, startRow);

            //Add client sheet to Sheet Ids lookup
              const intBtm = shtSettings.getRange("D1:E1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()+1;
              shtSettings.getRange("D" + intBtm).setValue(strClientNm);
              shtSettings.getRange("E" + intBtm).setValue(strSheetID);
          }
        }
      }  
    } 
}

function getDSUMcriteria(strRangeNm){
  //Load array with Criteria's named range objects for DSUM formulas

      const shtSettings = ssActive.getSheetByName('Settings');
      const arNmRng = shtSettings.getNamedRanges();
      let arRngNms = [];
      let arCriteria = [];
      let x = 0;

      //--load array with the names of the named range objects
      let cnt_ar = arNmRng.length;
      for (d = 0; d<cnt_ar; d++){
        arRngNms[d] = arNmRng[d].getName();  
      }

      //--now i can sort the named ranges names so that they paste in the correct order
      let strDsumFormula;
      arRngNms.sort()
      
      //--load up DSUM forumlas
      cnt_ar = arRngNms.length;
      for (d = 0; d<cnt_ar; d++){
        strDsumFormula = '=dsum(' + strRangeNm + ',"Balance",' + arRngNms[d] + ')';
        arCriteria[x] = strDsumFormula;
        x++;
      }
      return arCriteria;
}

function findSheetID(shtSearch, strSheetNm){
   
  const arRng = shtSearch.createTextFinder(strSheetNm).findAll();
  const rngFound = arRng[0];

    if(rngFound){
      return rngFound.getValue();
    }
}

function mySort(sht, rw){
  sht.getRange("A" + rw).activate();
  sht.sort(1,true);
}

function createBandings_Aging(sht,intHeaderRow){
  
  //for testing
  //const sht = SpreadsheetApp.getActiveSheet(); 
  //const intHeaderRow = 7;

  // Get the active sheet and data range.
  const fullDataRange = sht.getDataRange();

  // Apply row banding to the data, excluding the header 
    const noHeadersRange = fullDataRange.offset(intHeaderRow, 0,
    fullDataRange.getNumRows() - intHeaderRow,
    fullDataRange.getNumColumns() - 1);
  
  if (noHeadersRange.getBandings()[0]){
    let bandings = noHeadersRange.getBandings()[0]
    bandings.remove()
  } 
  noHeadersRange.applyRowBanding(
        SpreadsheetApp.BandingTheme.LIGHT_GREY,
        false, false); 
}

function createBandings_Clients(sht,lastRw){

  const arNmRng = sht.getNamedRanges();  
  let noHeadersRange = arNmRng[0].getRange().offset(1,0); 
  
  if (noHeadersRange.getBandings()[0]){
    let bandings = noHeadersRange.getBandings()[0]
    bandings.remove()
  } 

  noHeadersRange = sht.getRange('A8:H' + lastRw);
  noHeadersRange.applyRowBanding(
        SpreadsheetApp.BandingTheme.LIGHT_GREY,
        false, false); 

  //reapply background to Balance column
  sht.getRange('E8:E' + noHeadersRange.getLastRow()).setBackground('#38761d');
  
}

function onEdit(e){
  
  const rngClicked = e.range;
  const rw = rngClicked.getRow();
  const shtActive = ssActive.getActiveSheet();
  
  //test for client sheet. This should not run on Aging Report sheet!
  if (shtActive.getName() != 'Aging Report' && shtActive.getName() != 'Template'){
    
    //Only triggered on edits to the invoice date column
    if (rngClicked.getColumn()==2 && row > 7){
      
      //apply past due formula
      const rngPastDue = shtActive.getRange('F' + rw);
      rngPastDue.setFormula('=TODAY()-B' + rw);

      //Expand named range and formatting if necessary
      const arNmRng = shtActive.getNamedRanges();
      
      if (arNmRng[0]){
        
        const intCurrNmRngLastRow = arNmRng[0].getRange().getLastRow();

        //test if user is on the last row of the named range and requires expanding
        if (shtActive.getCurrentCell().getRow() == intCurrNmRngLastRow){
            
            const intNewNmRngLastRow = intCurrNmRngLastRow+1;
            ssActive.setNamedRange(
              arNmRng[0].getName(), ssActive.getRange('A7:H' + intNewNmRngLastRow));
            
            //copy row
            const rwActive = shtActive.getActiveRange().getRow();
            const rngPaste = shtActive.getRange(rwActive + 1, 1, 1, shtActive.getMaxColumns());

            //paste row
            shtActive.getRange(
              ssActive.getCurrentCell().getRow(), 1, 1, shtActive.getMaxColumns()).copyTo(
                rngPaste, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
            
            ssActive.toast("Expand Named Range " + rwActive + ' ' + intNewNmRngLastRow);

            //clear pasted row of prior data
            shtActive.getRange(rwActive + 1,1,1,3).clearContent();
            shtActive.getRange(rwActive + 1,6).clearContent();

            createBandings_Clients(shtActive,intNewNmRngLastRow); 

            //Update summmary formulas
            const rngTotInvoice = shtActive.getRange('C2');
            rngTotInvoice.setFormula('=SUM(E8:E' + intNewNmRngLastRow + ')');

            const rngAmtPaid = shtActive.getRange('C3');
            rngAmtPaid.setFormula('=SUM(D8:D' + intNewNmRngLastRow + ')');
            
        }
      }
    }
  } 
}

function transposeArray(arr) {
  try{
    return arr[0].map((_, colIndex) => arr.map(row => row[colIndex]));
  } catch {
    return Array(arr)}
}

function link(rangeToBeLinked, strSheetID){
  //const rangeToBeLinked = SpreadsheetApp.getActiveSheet().getActiveRange();
  const strDestSheetNm = rangeToBeLinked.getValue();

  const richText = SpreadsheetApp.newRichTextValue()
      .setText(strDestSheetNm)
      .setLinkUrl('#gid=' + strSheetID)
      .build();
      
  rangeToBeLinked.setRichTextValue(richText);
}

function myPrompt(strOldClientNm, strCaller){

    let strTitle;
    let strPrompt;

    switch(strCaller){
      case 'rename':
        strTitle = "Rename " + strOldClientNm + "?";
        strPrompt = "Please enter a new client name:";
        break;

      case 'new':
        strTitle = "Create new client register";
        strPrompt = "Please enter the client's name:";
        break;
    }
    //Prompt user for new client name
    const inputBx = ui.prompt(
      strTitle,
      strPrompt,
      ui.ButtonSet.OK_CANCEL
    );

    const strResponse = inputBx.getResponseText();
    return [inputBx.getSelectedButton(),strResponse];
}

function isSheet(strSheetNm){

    const arAllSheets = ssActive.getSheets();
    const arShtNms = arAllSheets.map(sht => sht.getSheetName()); 
    if(arShtNms.includes(strSheetNm)){
      return true;
    }else{
      return false;      
    }
}

function lockit(sht,strDesc) {
  sht.protect()
    .setDescription(strDesc)
    .setWarningOnly(true); 
}

function unlockit(sht) {
  sht.protect().remove();
}

function prepRangeNm(strNewClientNm,strClientShtID){
    
  //remove spaces
    let strRangeNm = strNewClientNm.split(" ").join("");

  //remove punctuation
    const arPunctuation = [".","?","-","*", "/",","];
    for(const chr of arPunctuation){
      strRangeNm = strRangeNm.replaceAll(chr,"");  
    } 

  //remove leading numbers
    let arRangeNm = strRangeNm.split("",); 
    let x=0;
    
    for (const chr of arRangeNm){
        if(!Number(chr) && chr.valueOf() != 0){
          arRangeNm = arRangeNm.toSpliced(0,x);
          strRangeNm = arRangeNm.join('').trim();
          break;
        }
        x++;
    }

    strRangeNm = strRangeNm + '_id_' + strClientShtID;

  return strRangeNm;
}