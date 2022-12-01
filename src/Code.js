// WITH THE NAME OF ALLAH, THE MOST MERCIFUL, THE MOST BENEFICIENT
// PRAISE BE TO ALLAH, WHO HAS TAUGHT US WITH A PEN



// /**
//  * Creates a trigger for when a spreadsheet opens.
//  * @see https://developers.google.com/apps-script/guides/triggers/installable
//  */
//  function createSpreadsheetOpenTrigger() {
//     const ss = SpreadsheetApp.getActive();
//     ScriptApp.newTrigger('checkAuthorization')
//         .forSpreadsheet(ss)
//         .onOpen()
//         .create();
//   }

function onOpen(e) {
    // This creates the Organize Menu Item
    var ss = SpreadsheetApp.getUi(); 
    ss.createMenu('Organize')
    .addItem('Authorize Your Account', 'checkAuthorization')
    .addToUi(); 
  
  }

function editTrigger1() {
    ScriptApp.newTrigger('telecomEventProcessing')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create()
}; 

function checkAuthorization() {

    var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    var status = authInfo.getAuthorizationStatus();
    url = authInfo.getAuthorizationUrl();

    const allTriggers = ScriptApp.getProjectTriggers(); 

    const triggerTypes = []; 

    allTriggers.forEach( trigger => {

        const typeOfTrigger = trigger.getEventType(); 
        triggerTypes.push(typeOfTrigger); 

    }); 

    if ( triggerTypes[0] != 'ON_EDIT') {

        editTrigger1();
        Logger.log('on edit trigger installed based on negative condition!'); 
        Logger.log(triggerTypes[0]); 

    } else {

        Logger.log('No additional trigger installed as they are already installed'); 
        
        Logger.log(triggerTypes[0]);

    }

    Logger.log(status); 

    Logger.log(allTriggers); 

    Logger.log(triggerTypes); 
}


function telecomEventProcessing(e) {

    //Get the value of status from the current row 
    const activeRowArray = fun.getEventData(e).activeDataRowArray; 
    Logger.log('The active data row array is '); 
    Logger.log(activeRowArray); 
    const idOfCompany = activeRowArray[0][0]; 

    let statusValue = activeRowArray[0][22]; // status value 
    const meetingValue = activeRowArray[0][15]; // meeting value
    const callResponse = activeRowArray[0][8]; // call response
    let followUpStatus = activeRowArray[0][25]; 
    let negativeCounterScore = activeRowArray[0][28]; // negative score counter
    const callSheetNegativeCounterScore = activeRowArray[0][19]; // negative score counter for call sheet 
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
    Logger.log(`Calls sheet negative counter score is ${callSheetNegativeCounterScore}`); 
    const activeSheet = activeSpreadsheet.getActiveSheet();  
    const negativeCounterScoreCellRange = activeSheet.getRange(e.range.getRow(), 29); 
    const statusValueCellRange = activeSheet.getRange(e.range.getRow(), 23); 
    const callsSheetStatusCellValueRange = activeSheet.getRange(e.range.getRow(), 15); 
    const callsSheetStatusCellValue = callsSheetStatusCellValueRange.getValue(); 
    const activeSheetName = activeSheet.getName(); 
    const columnValue = activeSheet.getRange(1, e.range.getColumn()).getValue(); 
    const currentCellValue = activeSheet.getRange(e.range.getRow(), e.range.getColumn()).getValue(); 
    const currentCellRange = activeSheet.getRange(e.range.getRow(), e.range.getColumn()); 
    const IdLowerCellValue = activeSheet.getRange(e.range.getRow() + 1, 1).getValue(); 
    

    Logger.log(`the status value is ${statusValue} & the meeting value is ${meetingValue}, and call response value is ${callResponse} and the negative score counter is ${negativeCounterScore}`); // verified 

    // Execute conditional statements 

    if( activeSheetName == 'Sheet1' && currentCellValue == 'Lead' && meetingValue.length == 0 ) {

        SpreadsheetApp.getActive().toast('Condition for lead activated successfully!'); 

        // Transfer the current row record to the lowest load Inside Sales Executive 

        fun.loadBalancerCompany('Inside Sales Executive', fun.getEventData(e).companyID, e); 
        fun.loadBalancerActivity('Inside Sales Executive', fun.getEventData(e).companyID, e, 'Call'); 

        // Highlight the row according to status value 

        statusValueCellRange.setValue('Lead'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);
    

    } else if( activeSheetName == 'Sheet1' && statusValue == 'Opportunity' && meetingValue == 'yes') {

        SpreadsheetApp.getActive().toast('Condition for meeting activated successfully'); 

        // Transfer the current row record to the lowest load Marketing Executive 

        fun.loadBalancerCompany('Marketing Executive', fun.getEventData(e).companyID, e); 
        fun.loadBalancerActivity('Marketing Executive', fun.getEventData(e).companyID, e, 'Meeting'); 

        // Highlight the row according to status value 

        statusValueCellRange.setValue('Strong Opportunity'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);

    } else if ( activeSheetName == 'Sheet1' && statusValue == 'Opportunity' && meetingValue.length == 0 ) {

        SpreadsheetApp.getActive().toast('Condition For Opportunity with no meeting activated successfully'); 

        fun.loadBalancerCompany('Inside Sales Executive', fun.getEventData(e).companyID, e); 
        fun.loadBalancerActivity('Inside Sales Executive', fun.getEventData(e).companyID, e, 'Call'); 

        // Highlight the row according to status value 

        statusValueCellRange.setValue('Opportunity'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);

    } else if( activeSheetName == 'Sheet1' && columnValue == 'Call Response' && IdLowerCellValue.length == 0) {

        const SpreadSheetName = SpreadsheetApp.getActiveSpreadsheet().getName(); 
        const SpreadSheetNameArray = SpreadSheetName.split('-'); 
        const campaignId = SpreadSheetNameArray[4]; 
        const employeeId = SpreadSheetNameArray[3]; 

        Logger.log(`the campaign ID is ${campaignId} & the employee id is ${employeeId}`); 

        fun.extractData(campaignId, 10, 'Add'); 


    } else if ( activeSheetName == 'Sheet1' && columnValue == 'Call Response' && currentCellValue == 'Busy' && negativeCounterScore < 3 ) {

        // Reschedule a new call the next day & add 1 score to negative counter score

        SpreadsheetApp.getActive().toast('Busy Call detected successfully!'); 

        const newCallDate = Date.now() + 1 * 24 * 3600 * 1000; // This should give next day's date 

        Logger.log(`the new call date is ${newCallDate}`); 

        fun.rescheduleActivity('Call', e, newCallDate, '', callResponse); 

        SpreadsheetApp.getActive().toast('Busy Call activity completed successfully! Alhumdulillah!');

        // Highlight the row according to status value 

        statusValueCellRange.setValue('In Progress'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);

    } else if ( activeSheetName == 'Sheet1' && columnValue == 'Call Response' && currentCellValue == 'Not Answering' && negativeCounterScore < 3) {
        
        // Reschedule call after 3 days & add 1 score to negative counter score

        SpreadsheetApp.getActive().toast('Not Answering Call detected successfully!'); 

        const newCallDate = Date.now() + 3 * 24 * 3600 * 1000; // This should give next day's date 

        fun.rescheduleActivity('Call', e, newCallDate, '', callResponse); 

        SpreadsheetApp.getActive().toast('Not Answering Call activity completed successfully! Alhumdulillah!');

        // Highlight the row according to status value 

        statusValueCellRange.setValue('In Progress'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);


    } else if( activeSheetName == 'Sheet1' && callResponse == 'Picked Up' && followUpStatus == 'Call Back Later' && negativeCounterScore < 4) {

        // Reschedule call after 3 days & add 1 score to negative counter score

        SpreadsheetApp.getActive().toast('Call Back Later detected successfully!'); 

        const newCallDate = Date.now() + 3 * 24 * 3600 * 1000; // This should give next day's date 

        fun.rescheduleActivity('Call', e, newCallDate); 

        SpreadsheetApp.getActive().toast('Call Back Later activity completed successfully! Alhumdulillah!');

        // Highlight the row according to status value 

        statusValueCellRange.setValue('In Progress'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);

    } else if ( activeSheetName == 'Sheet1' && callResponse == 'Picked Up' && followUpStatus == 'Call Back in Specific Time') {

        // Reschedule call at the specific date & time and reduce 1 score from the negative counter score 

        const personCallBackDate = ''; 
        const personCallBackTime = ''; 

        fun.rescheduleActivity('Call', e, personCallBackDate, personCallBackTime); 

        SpreadsheetApp.getActive().toast('Call Back in specific date and time activated successfully!'); 

        // Highlight the row according to status value 

        statusValueCellRange.setValue('In Progress'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);

    } else if ( activeSheetName == 'Calls' && columnValue == 'Call Response' && currentCellValue == 'Busy' && callSheetNegativeCounterScore < 4) {

        SpreadsheetApp.getActive().toast('Call Sheet Call Response Busy Value detected successfully !')

        const newCallDate = Date.now() + 1 * 24 * 3600 * 1000; // This should give next day's date 

        Logger.log(`the new call date is ${newCallDate}`); 

        fun.rescheduleActivity('Call', e, newCallDate, '', currentCellValue); 

        SpreadsheetApp.getActive().toast('Busy Call activity completed successfully! Alhumdulillah!');

        // Highlight the row according to status value 

        callsSheetStatusCellValueRange.setValue('In Progress'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        fun.setStatusHighlighting(activeSheet, targetRow, callsSheetStatusCellValue);


    } else if ( activeSheetName == 'Calls' && columnValue == 'Call Response' && currentCellValue == 'Not Answering' && callSheetNegativeCounterScore < 4 ) {

        SpreadsheetApp.getActive().toast('Call Sheet Call Response Not Answering Value detected successfully !')

        const newCallDate = Date.now() + 3 * 24 * 3600 * 1000; // This should give next day's date 

        Logger.log(`the new call date is ${newCallDate}`); 

        fun.rescheduleActivity('Call', e, newCallDate, '', currentCellValue); 

        SpreadsheetApp.getActive().toast('Not Answering Call activity completed successfully! Alhumdulillah!');

        // Highlight the row according to status value 

        callsSheetStatusCellValueRange.setValue('In Progress'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        fun.setStatusHighlighting(activeSheet, targetRow, callsSheetStatusCellValue);


    } else if ( activeSheetName == 'Calls' && columnValue == 'Call Response' && (currentCellValue == 'Not Answering' || currentCellValue == 'Busy') && callSheetNegativeCounterScore == 4 ) {

        SpreadsheetApp.getActive().toast('Call Sheet Call Response Not Answering Value detected successfully !'); 
        Logger.log('Dead call with not answering or busy response detected successfully!'); 

        const newCallDate = Date.now(); // This should give today's date

        Logger.log(`the new call date is ${newCallDate}`); 

        fun.rescheduleActivity('Call', e, newCallDate, '', currentCellValue); 

        SpreadsheetApp.getActive().toast('Dead call related to Not Answering Call activity completed successfully! Alhumdulillah!');

        // Highlight the row according to status value 

        callsSheetStatusCellValueRange.setValue('Unresponsive'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        fun.setStatusHighlighting(activeSheet, targetRow, callsSheetStatusCellValue);


    } else if ( activeSheetName == 'Sheet1' && columnValue == 'Call Response' && currentCellValue == 'Invalid Number' ) {

        //invalid number detected successfully!

        Logger.log('Invalid Number detected successfully')

        // Update the negative counter score to 4 

        negativeCounterScore = 4; 

        statusValue = 'Unresponsive'; 

        // get the range and set the value 

        negativeCounterScoreCellRange.setValue(negativeCounterScore); 
        statusValueCellRange.setValue(statusValue); 

        // Update status highlighting to grey and strike through

        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), 'Sheet1', idOfCompany); 

        fun.setStatusHighlighting(activeSheet, targetRow, statusValue); 

    }
    
    else {

        SpreadsheetApp.getActive().toast('No condition satisfied!'); 
    }; 



}