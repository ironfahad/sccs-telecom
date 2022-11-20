//WITH THE NAME OF ALLAH, THE MOST MERCIFUL, THE MOST BENEFICIENT! 

const fun = {

    tasksSheetName: 'Tasks',
  
    printMessage: function (targetCellRange, message) {
  
      targetCellRange.setValue(message); 
  
  
    },
  
    findTargetEmployeeSpreadSheet: function (employeeName) {
  
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
  
  
    },
  
    loadBalancer: function (jobTitle, taskID) {

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

      // Update the status of project record in operations sheet 
  
      const targetProjectArray = operationsTotalProjectsArray.filter( project => {
      return project[0] == projectId; 
  
      })
      Logger.log('target project array will come here'); 
      Logger.log(targetProjectArray); 
  
  
      const indexOfTargetProject = operationsTotalProjectsArray.indexOf(targetProjectArray[0]); 
      Logger.log(indexOfTargetProject); 
  
      const targetProjectStatusRange = operationsSheet.getRange(indexOfTargetProject + 2, 9);
      targetProjectStatusRange.setValue(`Assigned to ${designation}`); 
  
      // Need to update this status futher when market research updates its status to Accepted 
  
  
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

      // Find spreadSheet & projects sheet of the lowest load employee

      const targetEmployeeSpreadsheetId = lowestLoadEmployeeArray[6]; 
      const targetEmployeeSS = SpreadsheetApp.openById(targetEmployeeSpreadsheetId); 
      const targetEmployeeProjectsSheet = targetEmployeeSS.getSheetByName('Projects');
      
      // Add a new project after the last row at lowest load employee's project's sheet 

      const LastRowRange = targetEmployeeProjectsSheet.getRange(targetEmployeeProjectsSheet.getLastRow() + 1, 1, 1, targetEmployeeProjectsSheet.getLastColumn()); 
      Logger.log('The array of target project that will come at the end is'); 
  
      Logger.log(targetProjectArray); 
  
      LastRowRange.setValues(targetProjectArray);  
  
      // Get the project folder link from the operations sheet's cell URL 
  
      const targetEmployeeEmailAddress = lowestLoadEmployeeArray[7];
      const projectFolderCellArray = operationsSheet.getRange(indexOfTargetProject + 2, 4).getFormula().split('/'); 
      Logger.log(projectFolderCellArray); 
      const folderLinkValue = projectFolderCellArray[5].split(','); 
  
      const folderlinkAddress = folderLinkValue[0].slice(0, -1); 
      Logger.log(`folder link address is ${folderlinkAddress}`); 

      // Share the folder with the lowest load employee

      


  
      // Ends here ... 
  
  
    }
  
  }; 
  
  