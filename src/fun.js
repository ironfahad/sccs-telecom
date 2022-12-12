//WITH THE NAME OF ALLAH, THE MOST MERCIFUL, THE MOST BENEFICIENT! 

const fun = {

  tasksSheetName: 'Tasks',

  printMessage: function (targetCellRange, message) {

    targetCellRange.setValue(message); 


  }, l: function (desc1, value1, desc2, value2, desc3, value3, desc4, value4, desc5, value5, desc6, value6) {


    // A powerful logging function for diagnostic purpose 

    Logger.log(desc1); 
    Logger.log(value1); 

    Logger.log(desc2); 
    Logger.log(value2);

    Logger.log(desc3); 
    Logger.log(value3);

    Logger.log(desc4); 
    Logger.log(value4);

    Logger.log(desc5); 
    Logger.log(value5);

    Logger.log(desc6); 
    Logger.log(value6);

  }, decideStatus: function (e) {


    const activeRowArray = fun.getEventData(e).activeDataRowArray; 

    // Create variables for case comparison 

    const meetingValue = activeRowArray[0][11]; 
    const productInterest = activeRowArray[0][9]; 
    const clientResponse = activeRowArray[0][8]; 

    this.l('meeting value is', meetingValue, 'product interst is', productInterest, 'client response is', clientResponse); 

    let statusValue; 


    switch(true) {
      

      case meetingValue == 'No' && productInterest == 'Low Interest' && clientResponse == 'Interested In Business' : 
      statusValue = 'Lead'; 
      break; 

      case meetingValue == 'No' && productInterest == 'Medium Interest' && clientResponse == 'Interested In Business' : 
      statusValue = 'Strong Lead'; 
      break; 

      case meetingValue == 'No' && productInterest == 'High Interest' && clientResponse == 'Interested In Business' : 
      statusValue = 'Opportunity'; 
      break;

      case meetingValue == 'Yes' && productInterest == 'Low Interest' && clientResponse == 'Interested In Business' : 
      statusValue = 'Low Opportunity'; 
      break;

      case meetingValue == 'Yes' && productInterest == 'Medium Interest' && clientResponse == 'Interested In Business' : 
      statusValue = 'Opportunity'; 
      break;

      case meetingValue == 'Yes' && productInterest == 'High Interest' && clientResponse == 'Interested In Business' : 
      statusValue = 'Strong Opportunity'; 
      break;

      case meetingValue == 'No' && productInterest == 'Undecided' && clientResponse == 'Call Back Later' : 
      statusValue = 'Call Back Later'; 
      break;

      case meetingValue == 'No' && productInterest == 'Undecided' && clientResponse == 'Call Back At Specified Time' : 
      statusValue = 'Call Back At Specified Time'; 
      break;

      default: statusValue = 'Condition Not Found'; 

      // Al Humdulillah! our switch statement has started working correctly! 
      

    }

    return statusValue; 




  }, findFirstEmptyRow: function (spreadsheetID, sheetName) {

    const ss = SpreadsheetApp.openById(spreadsheetID); 
    const sheet = ss.getSheetByName(sheetName); 
    const totalIdRange = sheet.getRange(2, 1, sheet.getLastRow(), 1); 
    const totalIdArray = totalIdRange.getValues(); 

    Logger.log(`The total id array is `); 
    Logger.log(totalIdArray); 


    // Find empty index 

    const emptyRowIndex = totalIdArray.findIndex ((index) => index[0].length === 0); 

    Logger.log(`Using findIndex formula, the index number is ${emptyRowIndex}`); 

    // Find empty row index

    // const lastRowIndex = totalIdArray.filter( id => {

    //   return id[0].length != 0; 

    // })

    // Logger.log(`findFirstEmptyRow Function: emptyRowIndex number is ${emptyRowIndex}`); 

    
    // Find the row number of the empty row 
    
    const firstEmptyRowNumber = emptyRowIndex + 2;  

    
    // Return the row number 

    return firstEmptyRowNumber; 

    Logger.log(`findFirstEmptyRow function : ${firstEmptyRowNumber}`); 


  }, findTargetEmployeeSpreadSheet: function (employeeName) {

    const nameOfEmployee = employeeName; 
    const employeesArray = resources.hrSS().employeesDataValues; 
    Logger.log('test employees array will come here!!')
    Logger.log(employeesArray); 
    const targetEmployeeDataRow = employeesArray.filter( row => {
      return row[1].includes(nameOfEmployee); 
    }); 

    Logger.log(targetEmployeeDataRow); 

    const employeeSpreadsheetId = targetEmployeeDataRow[0][6]; 

    return {employeeSpreadsheetId, targetEmployeeDataRow}; 
    
    },

    findRowNumber: function (spreadSheetId, sheetName, rowID) {
       
      const iDofSpreadsheet = spreadSheetId; 
      Logger.log('id of the test spreadsheet will come here')
      Logger.log(iDofSpreadsheet); // correct spreadsheetID is being extracted~ we are close to our problem! 
      const _sheetName = sheetName; 
      Logger.log('name of sheet will come here'); 
      Logger.log(_sheetName); // correct name is being extracted! 
      const iDofRow = rowID;
      Logger.log(iDofRow); 

      const tasksSheet = SpreadsheetApp.openById(iDofSpreadsheet).getSheetByName(_sheetName); 

      const targetSheetArray = tasksSheet.getRange(2, 1, tasksSheet.getLastRow() - 1, tasksSheet.getLastColumn()).getValues(); 

      const targetRow = targetSheetArray.filter( row => {
        return row[0] === iDofRow; 
      }); 

      Logger.log('target Row will come here'); 
      Logger.log(targetRow); 

      const indexOfTargetRow = targetSheetArray.indexOf(targetRow[0]); 

      const actualRowNumber = indexOfTargetRow + 2; 

      Logger.log('Actual Row Number will come here'); 

      Logger.log(actualRowNumber); // There seems to be some sort of an issue going on here as well! 


      return actualRowNumber; 

    },

    findtargetTaskRow: function (spreadSheetId, sheetName, rowID) {
       
      const iDofSpreadsheet = spreadSheetId; 
      Logger.log('id of the test spreadsheet will come here')
      Logger.log(iDofSpreadsheet); // correct spreadsheetID is being extracted~ we are close to our problem! 
      const _sheetName = sheetName; 
      Logger.log('name of sheet will come here'); 
      Logger.log(_sheetName); // correct name is being extracted! 
      const iDofRow = rowID;
      Logger.log(iDofRow); 

      const tasksSheet = SpreadsheetApp.openById(iDofSpreadsheet).getSheetByName(_sheetName); 

      const targetSheetArray = tasksSheet.getRange(2, 1, tasksSheet.getLastRow() - 1, tasksSheet.getLastColumn()).getValues(); 

      const targetRow = targetSheetArray.filter( row => {
        return row[0] === iDofRow; 
      }); 

      Logger.log('target Row will come here'); 
      Logger.log(targetRow); 

      const indexOfTargetRow = targetSheetArray.indexOf(targetRow[0]); 

      const actualRowNumber = indexOfTargetRow + 2; 

      Logger.log('Actual Row Number will come here'); 

      Logger.log(actualRowNumber); // There seems to be some sort of an issue going on here as well! 


      return targetRow; 

    },

  toggleTargetSheetToActiveColor: function (spreadSheetId, sheetName, rowNumber) {

    const ssId = spreadSheetId; 
    const sheet = sheetName; 
    const numberRow = rowNumber; 
    Logger.log('Row number will come here'); 
    Logger.log(numberRow); 

    const targetSheet = SpreadsheetApp.openById(ssId).getSheetByName(sheet); 
    
    const targetRange = targetSheet.getRange(numberRow, 1, 1, targetSheet.getLastColumn()); 
    targetRange.setBackground("#ffe28a"); 


    Logger.log('Color Function Executed Successfully!');
  }, 

  reprioritizeTaskRow: function (taskRowArray, priorityValue, totalTasksArray, targetSheet, indexOfDataRow ) {

    const _taskRow = taskRowArray[0];
    const _indexOfDataRow = indexOfDataRow; 
    Logger.log(' _task row will come here ')
    Logger.log(_taskRow); 

    const _priorityValue = priorityValue; // suppose requested priority value is 1 what will happen then? 

    const _totalTasksArray = totalTasksArray; 

    Logger.log('reprioritizetTaskRow function Total Tasks Array will come here'); 
    Logger.log(_totalTasksArray); 

    const _targetSheet = targetSheet; 

    const range = _targetSheet.getRange(2, 1, _targetSheet.getLastRow() - 1, _targetSheet.getLastColumn()); 

   

     _totalTasksArray.splice(_indexOfDataRow, 1); 
    Logger.log('after first splice'); 
    Logger.log(_indexOfDataRow); 
    Logger.log(_totalTasksArray); 

    _totalTasksArray.splice(_priorityValue, 0, _taskRow); 

   

    Logger.log('after second splice'); 
    Logger.log(_totalTasksArray); 

    for ( let i = 0; i < _totalTasksArray.length; i++) {

      _totalTasksArray[i].splice(5, 1, i); 

    }

    Logger.log(_totalTasksArray); 

    range.setValues(_totalTasksArray); 

  },

  highLightColor: function (spreadSheetId, sheetName, rowNumber) {

    const ssId = spreadSheetId; 
    const sheet = sheetName; 
    const numberRow = rowNumber; 
    Logger.log('Row number will come here'); 
    Logger.log(numberRow); 

    const targetSheet = SpreadsheetApp.openById(ssId).getSheetByName(sheet); 
    
    const targetRange = targetSheet.getRange(numberRow, 1, 1, targetSheet.getLastColumn()); 
    targetRange.setBackground("#fffeb3"); 


    Logger.log('Highlight Color Function Executed Successfully!');
  }, 

  unHighLightColor: function (spreadSheetId, sheetName, rowNumber) {

    const ssId = spreadSheetId; 
    const sheet = sheetName; 
    const numberRow = rowNumber; 
    Logger.log('Row number will come here'); 
    Logger.log(numberRow); 

    const targetSheet = SpreadsheetApp.openById(ssId).getSheetByName(sheet); 
    
    const targetRange = targetSheet.getRange(numberRow, 1, 1, targetSheet.getLastColumn()); 
    targetRange.setBackground("white"); 


    Logger.log('unHighlight Color Function Executed Successfully!');
  }, 

  getEventData: function (e) {

    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
    const activesheet = activeSpreadsheet.getActiveSheet(); 
    const sheetName = activesheet.getSheetName(); 
    Logger.log('Event data Sheet Name Check will come here ')
    Logger.log(sheetName); // correct sheet is being extracted! 
    const dataRange = activesheet.getRange(2, 1, activesheet.getLastRow() - 1, activesheet.getLastColumn());
    const totalDataArray = dataRange.getValues(); 
    const activeRowRange = activesheet.getRange(e.range.getRow(), 1, 1, activesheet.getLastColumn()); 
    const activeDataRowArray = activeRowRange.getValues(); 

    // const activeRowNumber = activesheet.getRow(); 
    // const activeRowArrayIndex = activeRowNumber - 2; 



    const companyID = activeDataRowArray[0][0]; 
    const companyName = activeDataRowArray[0][1]; 
    const companyAddress = activeDataRowArray[0][2]; 
    const companyCity = activeDataRowArray[0][3]; 
    const companyPersonName = activeDataRowArray[0][4]; 
    const companyPersonMobile = activeDataRowArray[0][5]; 
    const companyLandline = activeDataRowArray[0][6]; 
    const companyEmail = activeDataRowArray[0][7]; 
    const callResponse = activeDataRowArray[0][8]; 
    const contactNameVerification = activeDataRowArray[0][9];
    const actualContactName = activeDataRowArray[0][10];
    const companyNameVerification = activeDataRowArray[0][11];
    const actualCompanyName = activeDataRowArray[0][12];
    const clientResponse = activeDataRowArray[0][13];
    const meetingGranted = activeDataRowArray[0][14];
    const meetingTime = activeDataRowArray[0][15];
    const meetingDate = activeDataRowArray[0][16];
    const askedQuestions = activeDataRowArray[0][17];
    const givenComplements = activeDataRowArray[0][18];
    const needProductData = activeDataRowArray[0][19];
    const interestLevel = activeDataRowArray[0][20];
    const specificPackageInquiry = activeDataRowArray[0][21];
    const companyStatus = activeDataRowArray[0][22];
    const remarks = activeDataRowArray[0][23];
    const competitionType = activeDataRowArray[0][24];
    const followUpStatus = activeDataRowArray[0][25];
    const callBackDate = activeDataRowArray[0][26];
    const callBackTime = activeDataRowArray[0][27];
    const negativeCounterScore = activeDataRowArray[0][28];



    Logger.log('Event Object Executed Successfully'); 


    return {activeSpreadsheet, activesheet, dataRange, 
      totalDataArray, activeRowRange, activeDataRowArray, companyID, companyName, companyAddress, 
      companyCity, companyPersonName, companyPersonMobile, companyLandline, companyEmail, callResponse, 
      contactNameVerification, actualContactName, companyNameVerification, actualCompanyName, clientResponse, 
      meetingGranted, meetingTime, meetingDate, askedQuestions, givenComplements, needProductData, interestLevel, 
      specificPackageInquiry, companyStatus, remarks, competitionType}; 
    

  },  

  generateUniqueArray: function (targetArray) {

    const duplicateEliminatedArray = targetArray.filter((name, index, array) => {

      return array.indexOf(name) === index; 

    }); 

    return duplicateEliminatedArray; 

  },

  getTargetEmployeeData: function (employeeName) {

    const ss = ''; 

    const taskSheet = ''; 
    const taskSheetName = ''; 

    const trainingSheet = ''; 
    const trainingSheetName = ''; 
    

    const taskSheetDataRange = ''; 
    const taskSheetTotalDataArray = ''; 

    const taskSingleRowArray = ''; 


  }, checkDuplicateRow: function ( dataComparisonId, sourceSheet, targetSheet) {

    // Check the id of the row to match 

    // Get the source and target ranges 

    const targetSheetIdRange = targetSheet.getRange( 2, 2, targetSheet.getLastRow() + 1, 1); // There is no data row already and only the header row. which has already been subtracted. So number of rows is zero
    const targetSheetIdValues = targetSheetIdRange.getValues(); 

    Logger.log("the id sheet values are: "); 
    Logger.log(targetSheet.getName()); 
    Logger.log(`data comparison id is ${dataComparisonId}`); 

    Logger.log(targetSheetIdValues); 

    // Filter the row ID from the target range 

    const duplicateFoundId = targetSheetIdValues.filter( value => {

      return value[0] == dataComparisonId; 
    }); 

    Logger.log(`The duplicate found Id is ${duplicateFoundId}`); 

    return duplicateFoundId; 


  }, loadBalancerCompany: function (jobTitle, taskID, e) {

    // Acquire functional parameters 

    const designation = jobTitle; 
    const projectId = taskID; 

    // Acquire Resources Data

    const employeesSheet = resources.strategicSS().ss.getSheetByName('HRM'); 
    const employeesDataRange = employeesSheet.getRange(2, 1, employeesSheet.getLastRow() - 1, employeesSheet.getLastColumn()); 
    const employeesDataArray = employeesDataRange.getValues(); 
    const operationsSheet = resources.strategicSS().operationsSheet; 
    const projectsRange = operationsSheet.getRange(2, 1, operationsSheet.getLastRow() - 1, 11);
    const operationsTotalProjectsArray = projectsRange.getValues();  
    const strategicTasksSheet = resources.strategicSS().ss.getSheetByName('Strategic Management');

    // Verify the accuracy of the resources data
     
    Logger.log('load balancer projectID will come here'); 
    Logger.log(projectId); 
 
    Logger.log('Total projects Array will come here'); 
    Logger.log(operationsTotalProjectsArray);  

    // Find employees with similar designation 

    const matchingEmployeesArray = employeesDataArray.filter( employee => {
      return employee[4] === designation; 
    }); 

    // Find the designated employee with lowest load value

    const currentLoadValues = []; 

    matchingEmployeesArray.forEach(employee => {
      currentLoadValues.push(employee[9]); 
    })

    Logger.log('the current load values are'); 
    Logger.log(currentLoadValues); 

    const lowestLoadValue = Math.min(...currentLoadValues); 
    Logger.log('The lowest load value is:'); 
    Logger.log(lowestLoadValue); 
    const indexOfLowestLoadValue = currentLoadValues.indexOf(lowestLoadValue); 
    Logger.log('The index of lowest load value is: '); 
    Logger.log(indexOfLowestLoadValue); 
    const lowestLoadEmployeeArray = matchingEmployeesArray[indexOfLowestLoadValue]; 
    Logger.log('The data array of employee with lowest load value'); 
    Logger.log(lowestLoadEmployeeArray); 

    const lowestLoadEmployeeId = lowestLoadEmployeeArray[0]; 
    Logger.log("primary index No. is: "); 

    const primaryIndexOfLowestLoadEmployeeArray = employeesDataArray.indexOf(lowestLoadEmployeeArray); 
    Logger.log(primaryIndexOfLowestLoadEmployeeArray);

    // Update the load value of employee with lowest load value

    lowestLoadEmployeeArray[9] = lowestLoadEmployeeArray[9] + 1; 
    const sSLowestLoadEmployeeRange = employeesSheet.getRange(primaryIndexOfLowestLoadEmployeeArray + 2, 10); 
    sSLowestLoadEmployeeRange.setValue(lowestLoadEmployeeArray[9]); 
    

    // Alhumdulillah Excellent work so far! the load balancer is on its way to become a power function InshAllah!

    // Find the campaign ID of the designated employee

    const designatedEmployeeCampaignId = employeesSheet.getRange(primaryIndexOfLowestLoadEmployeeArray + 2, 12).getValue(); 

    // Find spreadSheet & campaigns sheet of the lowest load employee

    const targetEmployeeSpreadsheetId = lowestLoadEmployeeArray[6]; 
    const targetEmployeeSS = SpreadsheetApp.openById(targetEmployeeSpreadsheetId); 
    const targetEmployeeCampaignSheet = targetEmployeeSS.getSheetByName('Campaigns');

    // Find the campaign record matching the campaign ID of the designated employee 

    const targetEmployeeCampSheetRange = targetEmployeeCampaignSheet.getRange(2, 1, targetEmployeeCampaignSheet.getLastRow() -1,targetEmployeeCampaignSheet.getLastColumn()); 
    const targetEmployeeCampaignDataArray = targetEmployeeCampSheetRange.getValues(); 

    const activeCampaignRow = targetEmployeeCampaignDataArray.filter( campaign => {

      return campaign[0] == designatedEmployeeCampaignId; 

    }); 

    Logger.log('The active campaign row is'); 

    Logger.log(activeCampaignRow); 

    // Extract the url from the targetlist cell 

    const indexOfActiveCampaignRow = targetEmployeeCampaignDataArray.indexOf(activeCampaignRow[0]); 

    Logger.log(`Index of active campaign row is ${indexOfActiveCampaignRow}`); 

    SpreadsheetApp.getActive().toast(`Index of active campaign row is ${indexOfActiveCampaignRow}`);
    
    const campaignTargetListCellRange = targetEmployeeCampaignSheet.getRange(indexOfActiveCampaignRow + 2, 5); 
    const campaignTargetListCellArray = campaignTargetListCellRange.getFormula().split('/');
    Logger.log('the split formula is'); 
    Logger.log(campaignTargetListCellArray); 

    // Extract the file ID from the URL formula

    const campaignTargetListFileLinkValue = campaignTargetListCellArray[5].split(','); 
    Logger.log('The campaign link value array is '); 
    Logger.log(campaignTargetListFileLinkValue); 
    const campaignTargetListFileId = campaignTargetListFileLinkValue[0];  

    Logger.log(`The campaign target list file ID is ${campaignTargetListFileId}`); // verified successfully! Alhumdulillah

    // Open spreadsheet and companies list sheet of the campaign file

    const designatedCampaignSpreadSheet = SpreadsheetApp.openById(campaignTargetListFileId); 
    const companiesListSheet = designatedCampaignSpreadSheet.getSheetByName('Sheet1'); 

    // find the first empty row 

    const emptyRowNumber = this.findFirstEmptyRow(campaignTargetListFileId, 'Sheet1'); 


    const companiesListLastRowRange = companiesListSheet.getRange( emptyRowNumber, 1, 1, companiesListSheet.getLastColumn());  

    // Get the active sheet data row 

    const activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet(); 
    const activeSheet = activeSpreadSheet.getActiveSheet(); 
    const activeSheetName = activeSheet.getName(); 
    const activeRowArray = fun.getEventData(e).activeDataRowArray; 
    const companylistSheet = activeSpreadSheet.getSheetByName('Sheet1'); 
    const callsSheet = activeSpreadSheet.getSheetByName('Calls'); 
    const meetingsSheet = activeSpreadSheet.getSheetByName('Meetings'); 

    if ( activeSheetName == 'Sheet1') {


      const ActiveCompanyDataArray = activeRowArray; 

      Logger.log(`Active company data array is`); 

      Logger.log(ActiveCompanyDataArray); 

      companiesListLastRowRange.setValues(ActiveCompanyDataArray);

      Logger.log("loadBalancerCompany function has been executed successfully!");


    } else if ( activeSheetName == 'Calls') {
 

     const companyListRowNumber = this.findRowNumber(activeSpreadSheet.getId(), 'Sheet1', activeRowArray[0][1]); 

     const companyListRowRange = companylistSheet.getRange(companyListRowNumber, 1, 1, companylistSheet.getLastColumn()); 
     const companyListRowArray = companyListRowRange.getValues(); 

     companiesListLastRowRange.setValues(companyListRowArray); 

     Logger.log("loadBalancerCompany function has been executed successfully!");


    } else if ( sheetName == 'Meetings') {



    } else if ( sheetName == 'Tasks') {



    }

    // Here we need to design a custom array that can fit to the companies list or we need to find the company array from the sheet1 .. lets see what can we do
    

    // Set the data to the target sheet 

    // Logger.log(`Active company data array is`); 

    // Logger.log(ActiveCompanyDataArray); 

    // companiesListLastRowRange.setValues(ActiveCompanyDataArray);

    // Logger.log("loadBalancerCompany function has been executed successfully!"); 

  }, 

  loadBalancerActivity: function (jobTitle, taskID, e, activityType, remarksHistory) {

    // Acquire functional parameters 

    const designation = jobTitle; 
    const projectId = taskID; 
    const typeOfActivity = activityType; 



    // Acquire Resources Data

    const employeesSheet = resources.strategicSS().ss.getSheetByName('HRM'); 
    const employeesDataRange = employeesSheet.getRange(2, 1, employeesSheet.getLastRow() - 1, employeesSheet.getLastColumn()); 
    const employeesDataArray = employeesDataRange.getValues(); 
    const operationsSheet = resources.strategicSS().operationsSheet; 
    const projectsRange = operationsSheet.getRange(2, 1, operationsSheet.getLastRow() - 1, 11);
    const operationsTotalProjectsArray = projectsRange.getValues();  
    const strategicTasksSheet = resources.strategicSS().ss.getSheetByName('Strategic Management'); 

    // Verify the accuracy of the resources data
     
    Logger.log('load balancer projectID will come here'); 
    Logger.log(projectId); 
 
    Logger.log('Total projects Array will come here'); 
    Logger.log(operationsTotalProjectsArray);  

    // Find employees with similar designation 

    const matchingEmployeesArray = employeesDataArray.filter( employee => {
      return employee[4] === designation; 
    }); 

    // Find the designated employee with lowest load value

    const currentLoadValues = []; 

    matchingEmployeesArray.forEach(employee => {
      currentLoadValues.push(employee[9]); 
    })

    Logger.log('the current load values are'); 
    Logger.log(currentLoadValues); 

    const lowestLoadValue = Math.min(...currentLoadValues); 
    Logger.log('The lowest load value is:'); 
    Logger.log(lowestLoadValue); 
    const indexOfLowestLoadValue = currentLoadValues.indexOf(lowestLoadValue); 
    Logger.log('The index of lowest load value is: '); 
    Logger.log(indexOfLowestLoadValue); 
    const lowestLoadEmployeeArray = matchingEmployeesArray[indexOfLowestLoadValue]; 
    Logger.log('The data array of employee with lowest load value'); 
    Logger.log(lowestLoadEmployeeArray); 

    const lowestLoadEmployeeId = lowestLoadEmployeeArray[0]; 
    Logger.log("primary index No. is: "); 

    const primaryIndexOfLowestLoadEmployeeArray = employeesDataArray.indexOf(lowestLoadEmployeeArray); 
    Logger.log(primaryIndexOfLowestLoadEmployeeArray);

    // Update the load value of employee with lowest load value

    lowestLoadEmployeeArray[9] = lowestLoadEmployeeArray[9] + 1; 
    const sSLowestLoadEmployeeRange = employeesSheet.getRange(primaryIndexOfLowestLoadEmployeeArray + 2, 10); 
    sSLowestLoadEmployeeRange.setValue(lowestLoadEmployeeArray[9]); 
    

    // Alhumdulillah Excellent work so far! the load balancer is on its way to become a power function InshAllah!

    // Find the campaign ID of the designated employee

    const designatedEmployeeCampaignId = employeesSheet.getRange(primaryIndexOfLowestLoadEmployeeArray + 2, 12).getValue(); 

    // Find spreadSheet & campaigns sheet of the lowest load employee

    const targetEmployeeSpreadsheetId = lowestLoadEmployeeArray[6]; 
    const targetEmployeeSS = SpreadsheetApp.openById(targetEmployeeSpreadsheetId); 
    const targetEmployeeCampaignSheet = targetEmployeeSS.getSheetByName('Campaigns');

    // Find the campaign record matching the campaign ID of the designated employee 

    const targetEmployeeCampSheetRange = targetEmployeeCampaignSheet.getRange(2, 1, targetEmployeeCampaignSheet.getLastRow() -1,targetEmployeeCampaignSheet.getLastColumn()); 
    const targetEmployeeCampaignDataArray = targetEmployeeCampSheetRange.getValues(); 

    const activeCampaignRow = targetEmployeeCampaignDataArray.filter( campaign => {

      return campaign[0] == designatedEmployeeCampaignId; 

    }); 

    Logger.log('The active campaign row is'); 

    Logger.log(activeCampaignRow); 

    // Extract the url from the targetlist cell 

    const indexOfActiveCampaignRow = targetEmployeeCampaignDataArray.indexOf(activeCampaignRow[0]); 

    Logger.log(`Index of active campaign row is ${indexOfActiveCampaignRow}`); 

    SpreadsheetApp.getActive().toast(`Index of active campaign row is ${indexOfActiveCampaignRow}`);
    
    const campaignTargetListCellRange = targetEmployeeCampaignSheet.getRange(indexOfActiveCampaignRow + 2, 5); 
    const campaignTargetListCellArray = campaignTargetListCellRange.getFormula().split('/');
    Logger.log('the split formula is'); 
    Logger.log(campaignTargetListCellArray); 

    // Extract the file ID from the URL formula

    const campaignTargetListFileLinkValue = campaignTargetListCellArray[5].split(','); 
    Logger.log('The campaign link value array is '); 
    Logger.log(campaignTargetListFileLinkValue); 
    const campaignTargetListFileId = campaignTargetListFileLinkValue[0];  

    Logger.log(`The campaign target list file ID is ${campaignTargetListFileId}`); // verified successfully! Alhumdulillah

    // Open spreadsheet and activity sheet of the campaign file 



    if(typeOfActivity == 'Call') {

      this.l('detected type of activity is Call'); 

      const activeRowArray = this.getEventData(e).activeDataRowArray; 
      const remarksHistory = activeRowArray[0][23]; 
      const callResponse = activeRowArray[0][7]; 
      const previousCallRemarks = activeRowArray[0][15]; 

      const designatedEmployeeSS = SpreadsheetApp.openById(campaignTargetListFileId); 

      const designatedEmployeeCallSheet = designatedEmployeeSS.getSheetByName('Calls'); 
      const designatedEmployeeCallSheetRange = designatedEmployeeCallSheet.getRange(designatedEmployeeCallSheet.getLastRow() + 1, 1, 1, designatedEmployeeCallSheet.getLastColumn()); 
      const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const activeSheetName = activeSheet.getName();

      // const callDataArray = []; 

      // // Populate the row date for calls

      // callDataArray[0] = Math.floor(Math.random() * 10000000000); 
      // callDataArray[1] = fun.getEventData(e).companyID; 
      // callDataArray[2] = fun.getEventData(e).companyPersonMobile; 
      // callDataArray[3] = fun.getEventData(e).companyLandline; 
      // callDataArray[4] = fun.getEventData(e).companyPersonName; 
      // callDataArray[5] = fun.getEventData(e).companyName; 
      // callDataArray[6] = fun.getEventData(e).remarks; 
      // callDataArray[7] = fun.getEventData(e).needProductData; 
      // callDataArray[8] = 'Outbound'; 
      // callDataArray[9] = '';
      // callDataArray[10] = '';
      // callDataArray[11] = '';
      // callDataArray[12] = '';
      // callDataArray[13] = ''; 
       

      // Logger.log('the Call Data array is'); 
      // Logger.log(callDataArray); 

      // Check duplicate value

      const duplicateValueCheck = this.checkDuplicateRow(activeRowArray[0][0], '' , designatedEmployeeCallSheet);  

      Logger.log(`the duplicate value check is ${duplicateValueCheck}`); 
      Logger.log(`the length of duplicate value check is ${duplicateValueCheck.length}`); 
      Logger.log('Hello World!'); 

      Logger.log('name of the active sheet is '); 
      Logger.log(activeSheetName); 

      // end of duplicate value check function 

      // Call history processing algorithm... 

      const existingHistoryValue = activeSheet.getRange(e.range.getRow(), 7).getValue(); 
      Logger.log(`history value is ${existingHistoryValue}`); 
      const existingHistoryValueArray = existingHistoryValue.split('-'); 
      Logger.log('History value array is ');
      Logger.log(existingHistoryValueArray)
      Logger.log('array length'); 
      Logger.log(existingHistoryValueArray.length);
      existingHistoryValueArray.push(callResponse); 
      existingHistoryValueArray.push(previousCallRemarks); 
      Logger.log(existingHistoryValueArray); 
      const newStringofHistory = existingHistoryValueArray.join("--"); 
      Logger.log(newStringofHistory);

      // End of history algorithm 

      if( duplicateValueCheck.length == 0 && activeSheetName == 'Sheet1') {

        this.l('sheet1 condition detected!')

        const callDataArray = []; 

        callDataArray[0] = Math.floor(Math.random() * 10000000000); 
        callDataArray[1] = activeRowArray[0][0]; // company ID 
        callDataArray[2] = activeRowArray[0][5]; // cell phone number 
        callDataArray[3] = activeRowArray[0][6]; // company landline number 
        callDataArray[4] = activeRowArray[0][4]; // person name 
        callDataArray[5] = activeRowArray[0][1]; // company name 
        callDataArray[6] = newStringofHistory; // Call History 
        callDataArray[7] = ''; 
        callDataArray[8] = ''; 
        callDataArray[9] = activeRowArray[0][19]; // need product data
        callDataArray[10] = 'Outbound'; 
        callDataArray[11] = '';
        callDataArray[12] = '';
        callDataArray[13] = '';
        callDataArray[14] = '';
        callDataArray[15] = ''; 
        callDataArray[16] = 'Planned'; // follow up status 
        callDataArray[17] = new Date(activeRowArray[0][16]).toDateString(); 
        callDataArray[18] = new Date(activeRowArray[0][15]).toLocaleTimeString('en-US'); 
        callDataArray[19] = activeRowArray[0][28] + 1; // negative counter score 
        

        Logger.log('the Call Data array is'); 
        Logger.log(callDataArray);

        designatedEmployeeCallSheetRange.setValues([callDataArray]);
        designatedEmployeeCallSheet.getRange(2, 1, designatedEmployeeCallSheet.getLastRow() - 1, designatedEmployeeCallSheet.getLastColumn()).sort([{ column: 11, ascending: true }]);
        // designatedEmployeeCallSheet.sort(1);  


      } else if ( duplicateValueCheck == 0 && activeSheetName == 'Calls') {

        

        const callDataArray = activeRowArray; 

        callDataArray[0][6] = newStringofHistory; 
        callDataArray[0][7] = ''; 
        callDataArray[0][8] = '';
        callDataArray[0][9] = '';
        callDataArray[0][11] = '';
        callDataArray[0][12] = '';
        callDataArray[0][13] = '';
        callDataArray[0][14] = '';
        callDataArray[0][15] = '';

        this.l('Active Sheet detected from telecom executive is Calls Sheet and the data is', activeRowArray); 

        // callDataArray[0] = Math.floor(Math.random() * 10000000000); 
        // callDataArray[1] = activeRowArray[0][0]; // company ID 
        // callDataArray[2] = activeRowArray[0][5]; // cell phone number 
        // callDataArray[3] = activeRowArray[0][6]; // company landline number 
        // callDataArray[4] = activeRowArray[0][4]; // person name 
        // callDataArray[5] = activeRowArray[0][1]; // company name 
        // callDataArray[6] = remarksHistory; // Call History 
        // callDataArray[7] = ''; 
        // callDataArray[8] = ''; 
        // callDataArray[9] = activeRowArray[0][19]; // need product data
        // callDataArray[10] = 'Outbound'; 
        // callDataArray[11] = '';
        // callDataArray[12] = '';
        // callDataArray[13] = '';
        // callDataArray[14] = '';
        // callDataArray[15] = ''; 
        // callDataArray[16] = 'Planned'; // follow up status 
        // callDataArray[17] = new Date(activeRowArray[0][16]).toDateString(); 
        // callDataArray[18] = new Date(activeRowArray[0][15]).toLocaleTimeString('en-US'); 
        // callDataArray[19] = activeRowArray[0][28] + 1; // negative counter score 
        

        Logger.log('the Call Data array is'); 
        Logger.log(callDataArray);

        designatedEmployeeCallSheetRange.setValues(callDataArray);
        designatedEmployeeCallSheet.getRange(2, 1, designatedEmployeeCallSheet.getLastRow() - 1, designatedEmployeeCallSheet.getLastColumn()).sort([{ column: 11, ascending: true }]);
        // designatedEmployeeCallSheet.sort(1);

      }
      

    } else if( typeOfActivity == 'Meeting'){

      const activeRowArray = this.getEventData(e).activeDataRowArray; 
      const remarksHistory = activeRowArray[0][23]; 

      const designatedEmployeeMeetingSheet = SpreadsheetApp.openById(campaignTargetListFileId).getSheetByName('Meetings'); 
      const designatedEmployeeMeetingSheetRange = designatedEmployeeMeetingSheet.getRange(designatedEmployeeMeetingSheet.getLastRow() + 1, 1, 1, designatedEmployeeMeetingSheet.getLastColumn()); 

      // const callDataArray = []; 

      // // Populate the row date for calls

      // callDataArray[0] = Math.floor(Math.random() * 10000000000); 
      // callDataArray[1] = fun.getEventData(e).companyID; 
      // callDataArray[2] = fun.getEventData(e).companyPersonMobile; 
      // callDataArray[3] = fun.getEventData(e).companyLandline; 
      // callDataArray[4] = fun.getEventData(e).companyPersonName; 
      // callDataArray[5] = fun.getEventData(e).companyName; 
      // callDataArray[6] = fun.getEventData(e).remarks; 
      // callDataArray[7] = fun.getEventData(e).needProductData; 
      // callDataArray[8] = 'Outbound'; 
      // callDataArray[9] = '';
      // callDataArray[10] = '';
      // callDataArray[11] = '';
      // callDataArray[12] = '';
      // callDataArray[13] = ''; 
       

      // Logger.log('the Call Data array is'); 
      // Logger.log(callDataArray); 

      // Check duplicate value

      const duplicateValueCheck = this.checkDuplicateRow(activeRowArray[0][0], '' , designatedEmployeeMeetingSheet);  

      Logger.log(`the duplicate value check is ${duplicateValueCheck}`); 
      Logger.log(`the length of duplicate value check is ${duplicateValueCheck.length}`); 

      // end of duplicate value check function 

      if( duplicateValueCheck.length === 0) {

        const callDataArray = []; 

        callDataArray[0] = Math.floor(Math.random() * 10000000000); 
        callDataArray[1] = activeRowArray[0][0]; // company ID 
        callDataArray[2] = activeRowArray[0][5]; // cell phone number 
        callDataArray[3] = activeRowArray[0][6]; // company landline number 
        callDataArray[4] = activeRowArray[0][4]; // person name 
        callDataArray[5] = activeRowArray[0][1]; // company name 
        callDataArray[6] = remarksHistory; // Call History 
        callDataArray[7] = ''; 
        callDataArray[8] = ''; 
        callDataArray[9] = activeRowArray[0][19]; // need product data
        callDataArray[10] = 'Outbound'; 
        callDataArray[11] = '';
        callDataArray[12] = '';
        callDataArray[13] = '';
        callDataArray[14] = '';
        callDataArray[15] = ''; 
        callDataArray[16] = 'Planned'; // follow up status 
        callDataArray[17] = new Date(activeRowArray[0][16]).toDateString(); 
        callDataArray[18] = new Date(activeRowArray[0][15]).toLocaleTimeString('en-US'); 
        callDataArray[19] = activeRowArray[0][28] + 1; // negative counter score 
        

        Logger.log('the Call Data array is'); 
        Logger.log(callDataArray);

        designatedEmployeeMeetingSheetRange.setValues([callDataArray]);
        designatedEmployeeMeetingSheet.getRange(2, 1, designatedEmployeeMeetingSheet.getLastRow() - 1, designatedEmployeeMeetingSheet.getLastColumn()).sort([{ column: 11, ascending: true }]);
        // designatedEmployeeCallSheet.sort(1);  


      }


    } else if( typeOfActivity == 'Task') {

      const designatedEmployeeCallSheet = SpreadsheetApp.openById(campaignTargetListFileId).getSheetByName('Tasks'); 
      const designatedEmployeeCallSheetRange = designatedEmployeeCallSheet.getRange(designatedEmployeeCallSheet.getLastRow() + 1, 1, 1, designatedEmployeeCallSheet.getLastColumn()); 

      const callDataArray = []; 
      callDataArray[0] = Math.floor(Math.random() * 10000000000); 
      callDataArray[1] = fun.getEventData(e).companyID; 
      callDataArray[2] = fun.getEventData(e).companyPersonMobile; 
      callDataArray[3] = fun.getEventData(e).companyLandline; 
      callDataArray[4] = fun.getEventData(e).companyPersonName; 
      callDataArray[5] = fun.getEventData(e).companyName; 
      callDataArray[6] = fun.getEventData(e).remarks; 
      callDataArray[7] = fun.getEventData(e).needProductData; 
      callDataArray[8] = 'Outbound'; 
      callDataArray[9] = '';
      callDataArray[10] = '';
      callDataArray[11] = '';
      callDataArray[12] = '';
      callDataArray[13] = '';
      

      Logger.log('the Call Data array is'); 
      Logger.log(callDataArray); 

      designatedEmployeeCallSheetRange.setValues([callDataArray]); 
      designatedEmployeeCallSheet.sort(1); 
    }

    Logger.log("loadBalancerActivity function has been executed successfully!");
    

  }, extractData: function (projectId, NoOfRecords, dataAdditionMethod) {

    // Find the project row array

    SpreadsheetApp.getActive().toast('Extract Data function activated successfully!'); 

    const strategicOperationsSheet = resources.strategicSS().operationsSheet; 
    const strategicOperationsSheetRange = strategicOperationsSheet.getRange(2, 1, strategicOperationsSheet.getLastRow() - 1, strategicOperationsSheet.getLastColumn()); 
    const strategicOperationsSheetArray = strategicOperationsSheetRange.getValues(); 

    const projectRowArray = strategicOperationsSheetArray.filter(project => {

      return project[0] == projectId; 
    })

    // Get the data source cell 

    const indexOfProjectRowArray = strategicOperationsSheetArray.indexOf(projectRowArray[0]); 

    Logger.log(`Index of project row array is ${indexOfProjectRowArray}`); 

    SpreadsheetApp.getActive().toast(`Index of project row array is ${indexOfProjectRowArray}`);
    
    const projectTargetListCellRange = strategicOperationsSheet.getRange(indexOfProjectRowArray + 2, 5); 
    const projectTargetListCellArray = projectTargetListCellRange.getFormula().split('/');
    Logger.log('the split formula is'); 
    Logger.log(projectTargetListCellArray); 

    // Extract the file ID from the URL formula

    const projectTargetListFileLinkValue = projectTargetListCellArray[5].split(','); 
    Logger.log('The project link value array is '); 
    Logger.log(projectTargetListFileLinkValue); 
    const projectTargetListFileId = projectTargetListFileLinkValue[0];  

    Logger.log(`The project target list file ID is ${projectTargetListFileId}`); // verified successfully! Alhumdulillah

    SpreadsheetApp.getActive().toast('File Id seems to be extracted successfully!');

    // Source File 

    const sourceTargetlistSheet = SpreadsheetApp.openById(projectTargetListFileId).getSheetByName('Sheet1'); //verified 
    const sourceTargetListRange = sourceTargetlistSheet.getRange(2, 1, NoOfRecords, sourceTargetlistSheet.getLastColumn()); 
    const sourceTargetListArray = sourceTargetListRange.getValues(); 
    const sourceTargetListTotalDataRange = sourceTargetlistSheet.getRange(2, 1, sourceTargetlistSheet.getLastRow(), sourceTargetlistSheet.getLastColumn()); 

    const remainingTargetListRange = sourceTargetlistSheet.getRange(NoOfRecords + 2, 1, sourceTargetlistSheet.getLastRow() - NoOfRecords, sourceTargetlistSheet.getLastColumn()); 
    const remainingTargetListDataArray = remainingTargetListRange.getValues(); 
    const updatedTargetListRange = sourceTargetlistSheet.getRange(2, 1, remainingTargetListDataArray.length, remainingTargetListDataArray[0].length); 


    // Target File

    const targetFileId = SpreadsheetApp.getActiveSpreadsheet().getId(); 

    const targetSpreadSheetDataSheet = SpreadsheetApp.openById(targetFileId).getSheetByName('Sheet1'); 
    const targetSpreadSheetDataSheetRange = targetSpreadSheetDataSheet.getRange(2, 1, NoOfRecords, sourceTargetlistSheet.getLastColumn()); 
    const targetEmployeeTelecomDataArray = targetSpreadSheetDataSheetRange.getValues(); 

    // Data addition Methodology
    
    if (dataAdditionMethod == 'Replace') {

      // Delete all existing data in the target sheet and replace with fresh data 

      targetSpreadSheetDataSheetTotalDataRange = targetSpreadSheetDataSheet.getRange(2, 1, targetSpreadSheetDataSheet.getLastRow() + 1, targetSpreadSheetDataSheet.getLastColumn()); 
      targetSpreadSheetDataSheetTotalDataRange.clearContent(); 

      targetSpreadSheetDataSheetRange.setValues(sourceTargetListArray);

      SpreadsheetApp.getActive().toast('Data seems to be copied successfully!')
  
      const totaltargetlistDataRange = sourceTargetlistSheet.getRange(2, 1, sourceTargetlistSheet.getLastRow() + 1, sourceTargetlistSheet.getLastColumn()); 
  
      totaltargetlistDataRange.clearContent();
  
      updatedTargetListRange.setValues(remainingTargetListDataArray);
  
      SpreadsheetApp.getActive().toast('extractData function executed successfully - Alhumdulillah!'); 

    } else if ( dataAdditionMethod == 'Add') {

      // Add the new data after the last row of the target sheet 

      const targetSpreadSheetAddDataRange = targetSpreadSheetDataSheet.getRange(targetSpreadSheetDataSheet.getLastRow() + 1, 1, NoOfRecords, sourceTargetlistSheet.getLastColumn());

      targetSpreadSheetAddDataRange.setValues(sourceTargetListArray);

      SpreadsheetApp.getActive().toast('Data seems to be copied successfully!')
  
      const totaltargetlistDataRange = sourceTargetlistSheet.getRange(2, 1, sourceTargetlistSheet.getLastRow() + 1, sourceTargetlistSheet.getLastColumn()); 
  
      totaltargetlistDataRange.clearContent();
  
      updatedTargetListRange.setValues(remainingTargetListDataArray);
  
      SpreadsheetApp.getActive().toast('extractData function executed successfully - Alhumdulillah!'); 

    }


  }, rescheduleActivity: function (activityType, e, activityDate, ActivityTime, callResponse) {

    const ss = SpreadsheetApp.getActiveSpreadsheet(); 
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();  
    const callsSheet = ss.getSheetByName('Calls'); 
    const meetingsSheet = ss.getSheetByName('Meetings'); 
    const tasksSheet = ss.getSheetByName('tasks'); 
    



    if( activityType == 'Call' && activeSheetName == 'Sheet1') {

      const activeRowRange = this.getEventData(e).activeRowRange; 
      const activeRowArray = activeRowRange.getValues(); 

      const historyData = callResponse; 

      // Check duplicate value

      const duplicateValueCheck = this.checkDuplicateRow(activeRowArray[0][0], '' , callsSheet);  

      Logger.log(`the duplicate value check is ${duplicateValueCheck}`); 
      Logger.log(`the length of duplicate value check is ${duplicateValueCheck.length}`); 

      // end of duplicate value check function 

      if( duplicateValueCheck.length === 0) {

        const callDataArray = []; 

      callDataArray[0] = Math.floor(Math.random() * 10000000000); 
      callDataArray[1] = activeRowArray[0][0]; // company ID 
      callDataArray[2] = activeRowArray[0][5]; // cell phone number 
      callDataArray[3] = activeRowArray[0][6]; // company landline number 
      callDataArray[4] = activeRowArray[0][4]; // person name 
      callDataArray[5] = activeRowArray[0][1]; // company name 
      callDataArray[6] = historyData; // Call History 
      callDataArray[7] = ''; 
      callDataArray[8] = ''; 
      callDataArray[9] = activeRowArray[0][19]; // need product data
      callDataArray[10] = 'Outbound'; 
      callDataArray[11] = '';
      callDataArray[12] = '';
      callDataArray[13] = '';
      callDataArray[14] = '';
      callDataArray[15] = ''; 
      callDataArray[16] = 'Planned'; // follow up status 
      callDataArray[17] = new Date(activityDate).toDateString(); 
      callDataArray[18] = new Date(activityDate).toLocaleTimeString('en-US'); 
      callDataArray[19] = activeRowArray[0][28] + 1; // negative counter score 
      

      Logger.log('the Call Data array is'); 
      Logger.log(callDataArray); 

      callsSheet.getRange(callsSheet.getLastRow() + 1, 1, 1, callsSheet.getLastColumn()).setValues([callDataArray]);     
      callsSheet.getRange(2, 1, callsSheet.getLastRow() - 1, callsSheet.getLastColumn()).sort([{ column: 18, ascending: true }, { column: 17, ascending: true}]);




      } else {

        SpreadsheetApp.getActive().toast('Trran trran - Duplicate entry detected successfully!'); 
        Logger.log('Duplicate value detected successfully therefore skipping this operations!'); 
          
        }; 


      
    } else if (activityType == 'Call' && activeSheetName == 'Calls') {

      // All operations related to calls sheet 

      const activeRowRange = this.getEventData(e).activeRowRange; // Now the event is occuring in Calls Sheet 
      const activeRowArray = activeRowRange.getValues(); 
      const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
      const followUpStatusRange = activeSheet.getRange(e.range.getRow(), 17); 
      const callsSheetNegativeCounter = activeRowArray[0][19]; 
      let callSheetFollowUpStatus = ''; 

      // Conditional statements for followup status update! - On only those follow-up

      if ((callResponse == 'Busy' || callResponse == 'Not Answering') && callsSheetNegativeCounter < 4) {

        SpreadsheetApp.getActive().toast('Not held status detected based on condition!'); 

        followUpStatusRange.setValue('Not Held'); 


      } else if ( (callResponse == 'Busy' || callResponse == 'Not Answering') && callsSheetNegativeCounter == 4) {

        SpreadsheetApp.getActive().toast('Dead company detected!');

        followUpStatusRange.setValue('Dead'); 
        

      } else {

        followUpStatusRange.setValue('Held'); 
        
        

      }; 

      callSheetFollowUpStatus = followUpStatusRange.getValue(); // Lets test this 

      // Call history processing algorithm... 

      const existingHistoryValue = activeSheet.getRange(e.range.getRow(), 7).getValue(); 
      Logger.log(`history value is ${existingHistoryValue}`); 
      const existingHistoryValueArray = existingHistoryValue.split('-'); 
      Logger.log('History value array is ');
      Logger.log(existingHistoryValueArray)
      Logger.log('array length'); 
      Logger.log(existingHistoryValueArray.length);

      existingHistoryValueArray.push(callResponse); 
      Logger.log(existingHistoryValueArray); 
      const newStringofHistory = existingHistoryValueArray.join("--"); 
      Logger.log(newStringofHistory); 

      if (callsSheetNegativeCounter < 4) { 

        const callDataArray = []; 

      callDataArray[0] = Math.floor(Math.random() * 10000000000); 
      callDataArray[1] = activeRowArray[0][1]; // company ID 
      callDataArray[2] = activeRowArray[0][2]; // cell phone number 
      callDataArray[3] = activeRowArray[0][3]; // company landline number 
      callDataArray[4] = activeRowArray[0][4]; // person name 
      callDataArray[5] = activeRowArray[0][5]; // company name 
      callDataArray[6] = newStringofHistory; // Call History 
      callDataArray[7] = ''; // call response 
      callDataArray[8] = 'Interested in Business'; // client response 
      callDataArray[9] = ''; // product interest
      callDataArray[10] = 'Outbound'; 
      callDataArray[11] = ''; // meeting granted 
      callDataArray[12] = ''; // meeting date 
      callDataArray[13] = ''; // meeting time 
      callDataArray[14] = ''; // status 
      callDataArray[15] = ''; // remarks 
      callDataArray[16] = 'Planned'; // follow up status 
      callDataArray[17] = new Date(activityDate).toDateString(); 
      callDataArray[18] = new Date(activityDate).toLocaleTimeString('en-US'); 
      callDataArray[19] = activeRowArray[0][19] + 1; // negative counter score 
      

      Logger.log('the Call Data array is'); 
      Logger.log(callDataArray); 

      callsSheet.getRange(callsSheet.getLastRow() + 1, 1, 1, callsSheet.getLastColumn()).setValues([callDataArray]);     
       
    }

    const activityHistorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('activityhistory'); 
    const activityHistorySheetLastRowRange = activityHistorySheet.getRange(activityHistorySheet.getLastRow() + 1, 1, 1, activeRowArray[0].length); 
    activeRowArray[0][16] = callSheetFollowUpStatus; 
    activityHistorySheetLastRowRange.setValues(activeRowArray); 

    activeRowRange.clearContent(); 
    callsSheet.getRange(2, 1, callsSheet.getLastRow() - 1, callsSheet.getLastColumn()).sort([{ column: 18, ascending: true }, { column: 17, ascending: true}]);

  

    } else if (activityType == 'Task') {


    } else {

      SpreadsheetApp.getActive().toast('You have mentioned an invalid activity')

    }; 

  }, setStatusHighlighting: function (Sheet, rowNum, rowStatus) {

    // const targetSheet = spreadSheet.getActiveSheet(); 
    const targetRange = Sheet.getRange(rowNum, 1, 1, Sheet.getLastColumn()); 
    Logger.log("Status highling function in progress! ;-)")
    Logger.log("RowStatus Value will come here") 
    Logger.log(rowStatus); 
    Logger.log("Row number will come here") 
    Logger.log(rowNum); 
    targetRange.clearFormat(); 

    if (rowStatus === "Not Contacted") {
      targetRange
      .setBackground("white")
      .setFontStyle('normal')
      .setFontColor('black')
      .setFontFamily("Ropa Sans")
      .setHorizontalAlignment("center")

    } else if ( rowStatus === "In Progress") {
      targetRange
      .setBackground("#ffe28a")
      .setFontStyle("bold")
      .setFontColor("#666547")
      .setFontFamily("Ropa Sans")
      .setHorizontalAlignment("center")
    } else if (rowStatus === "Call Back Later") {
      targetRange
      .setBackground("#fffeb3")
      .setFontStyle("italic")
      .setFontColor("#666547")
      .setFontFamily("Ropa Sans")
      .setHorizontalAlignment("center"); 
    } else if(rowStatus === "Competition Limit") {
      targetRange
      .setBackground("grey")
      .setFontColor("white")
      .setFontWeight('bold')
      .setFontFamily("Ropa Sans")
      .setHorizontalAlignment("center")
      .setFontLine('line-through'); 
    } else if( rowStatus === "Opportunity" || rowStatus == "Strong Opportunity") {
      targetRange
      .setBackground("#6fcb9f")
      .setFontColor("#fb2e01")
      .setFontStyle("italic")
      .setFontLine('line-through')
      .setFontFamily("Ropa Sans")
      .setHorizontalAlignment("center"); 
    } else if(rowStatus === "Not Interested") {
      targetRange
      .setBackground("grey")
      .setFontColor("white")
      .setFontStyle("italic")
      .setFontLine('line-through')
      .setFontFamily("Ropa Sans")
      .setHorizontalAlignment("center"); 
    } else if ( rowStatus === "Lead" || rowStatus == "Strong Lead" || rowStatus == "Potential Lead") {
      targetRange
      .setBackground("#ffe28a")
      .setFontColor('#666547')
      .setFontWeight('bold')
      .setFontFamily("Ropa Sans")
      .setHorizontalAlignment("center")
      .setFontLine('line-through'); 
    } else if ( rowStatus === "Responded But No Interest") {
      targetRange
      .setBackground('grey')
      .setFontColor('white')
      .setFontStyle('italic')
      .setFontFamily("Ropa Sans")
      .setHorizontalAlignment("center")
      .setFontLine('line-through');
    } else if ( rowStatus === "Unresponsive"){
      targetRange
      .setBackground('grey')
      .setFontColor('white')
      .setFontStyle('italic')
      .setFontFamily("Ropa Sans")
      .setHorizontalAlignment("center")
      .setFontLine('line-through');
    } else if ( rowStatus === "Follow-Up Contact" || rowStatus === "Follow Up At Specified Time"){
      targetRange
      .setBackground('#fffeb3')
      .setFontColor('#666547')
      .setFontStyle('italic')
      .setFontFamily("Ropa Sans")
      .setHorizontalAlignment("center"); 
    }
    
    else {
      Logger.log("No matching criteria for status highlighting found"); 
    }; 

    Logger.log("Status Highlighting successfully applied! Al Humdulillah!"); 

  }

}; 

