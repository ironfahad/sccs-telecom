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
    const meetingValue = activeRowArray[0][14]; // meeting value
    const callResponse = activeRowArray[0][8]; // call response
    const callSheetCallResponse = activeRowArray[0][7]; 
    let followUpStatus = activeRowArray[0][25]; 
    let clientResponse = activeRowArray[0][13]; 
    let callSheetClientResponse = activeRowArray[0][8]; 
    const callSheetMeetingValue = activeRowArray[0][11]; 
    let negativeCounterScore = activeRowArray[0][28]; // negative score counter
    let callBackTime = activeRowArray[0][27]; 
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
    

    Logger.log(`the status value is ${statusValue} & the meeting value is ${meetingValue}, and call response value is ${callResponse} and the negative score counter is ${negativeCounterScore} & the client response is ${clientResponse} & the column value is ${columnValue} & call back time is ${callBackTime}`); // verified 

    // Execute conditional statements 

    if( activeSheetName == 'Sheet1' && columnValue == 'Remarks' && (statusValue == 'Lead' || statusValue == 'Strong Lead' || statusValue == 'Potential Lead' || statusValue == 'Opportunity' || statusValue == 'Strong Opportunity') && callResponse == 'Call Picked Up' && (meetingValue == 'No' || meetingValue.length == 0)) {

        SpreadsheetApp.getActive().toast('Record detected for inside sales executive'); 

        // Transfer the current row record to the lowest load Inside Sales Executive 

        fun.loadBalancerCompany('Inside Sales Executive', fun.getEventData(e).companyID, e); 
        fun.loadBalancerActivity('Inside Sales Executive', fun.getEventData(e).companyID, e, 'Call'); 

        // Highlight the row according to status value 

        // statusValueCellRange.setValue('Lead'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        statusValue = statusValueCellRange.getValue(); // For acquiring the updated value of the status!
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);
    

    } else if( activeSheetName == 'Sheet1' && columnValue == 'Remarks' && (statusValue == 'Lead' || statusValue == 'Strong Lead' || statusValue == 'Potential Lead' || statusValue == 'Opportunity' || statusValue == 'Strong Opportunity') && callResponse == 'Call Picked Up' && meetingValue == 'Yes') {

        SpreadsheetApp.getActive().toast('Record detected for marketing executive'); 

        // Transfer the current row record to the lowest load Marketing Executive 

        fun.loadBalancerCompany('Marketing Executive', fun.getEventData(e).companyID, e); 
        fun.loadBalancerActivity('Marketing Executive', fun.getEventData(e).companyID, e, 'Meeting'); 

        // Highlight the row according to status value 

        // statusValueCellRange.setValue('Strong Opportunity'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        statusValue = statusValueCellRange.getValue(); // For acquiring the updated value of the status!
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);

    } 
    //  else if( activeSheetName == 'Sheet1' && columnValue == 'Call Response' && IdLowerCellValue.length == 0) {

    //     const SpreadSheetName = SpreadsheetApp.getActiveSpreadsheet().getName(); 
    //     const SpreadSheetNameArray = SpreadSheetName.split('-'); 
    //     const campaignId = SpreadSheetNameArray[4]; 
    //     const employeeId = SpreadSheetNameArray[3]; 

    //     Logger.log(`the campaign ID is ${campaignId} & the employee id is ${employeeId}`); 

    //     fun.extractData(campaignId, 10, 'Add'); 


    // } 
    
    else if ( activeSheetName == 'Sheet1' && columnValue == 'Call Response' && currentCellValue == 'Busy' && negativeCounterScore < 3 ) {

        // Reschedule a new call the next day & add 1 score to negative counter score

        SpreadsheetApp.getActive().toast('Busy Call detected successfully!'); 

        const newCallDate = Date.now() + 1 * 24 * 3600 * 1000; // This should give next day's date 

        Logger.log(`the new call date is ${newCallDate}`); 

        fun.rescheduleActivity('Call', e, newCallDate, '', callResponse); 

        SpreadsheetApp.getActive().toast('Busy Call activity completed successfully! Alhumdulillah!');

        // Highlight the row according to status value 

        statusValueCellRange.setValue('In Progress'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        statusValue = statusValueCellRange.getValue(); // For acquiring the updated value of the status!
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
        statusValue = statusValueCellRange.getValue(); // For acquiring the updated value of the status!
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
        statusValue = statusValueCellRange.getValue(); // For acquiring the updated value of the status!
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);

    } else if ( activeSheetName == 'Sheet1' && callResponse == 'Call Picked Up' && clientResponse == 'Call Back At Specified Time' && activeRowArray[0][27].length != 0) {

        SpreadsheetApp.getActive().toast('Call Back In specific time detected successfully');

        // Reschedule call at the specific date & time and reduce 1 score from the negative counter score 

        const personCallBackDate = activeRowArray[0][26]; 
        const personCallBackTime = activeRowArray[0][27];
        
        Logger.log(`The call back date is ${personCallBackDate} & the call back time is ${personCallBackTime}`); 

        fun.rescheduleActivity('Call', e, personCallBackDate, personCallBackTime); 

        SpreadsheetApp.getActive().toast('Call Back in specific date and time activated successfully!'); 

        // Highlight the row according to status value 

        statusValueCellRange.setValue('In Progress'); // We need to update the setHighlighting function to add more status values into it!
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        statusValue = statusValueCellRange.getValue(); // For acquiring the updated value of the status!
        fun.setStatusHighlighting(activeSheet, targetRow, statusValue);

        // Starting from below all conditions are related to calls sheet - this code should gradually shift to another file for seperation of concerns

    } else if ( activeSheetName == 'Calls' && columnValue == 'Call Response' && currentCellValue == 'Busy' && callSheetNegativeCounterScore < 4) {

        SpreadsheetApp.getActive().toast('Call Sheet Call Response Busy Value detected successfully !')

        const newCallDate = Date.now() + 1 * 24 * 3600 * 1000; // This should give next day's date 

        Logger.log(`the new call date is ${newCallDate}`); 

        fun.rescheduleActivity('Call', e, newCallDate, '', currentCellValue); 

        SpreadsheetApp.getActive().toast('Busy Call activity completed successfully! Alhumdulillah!');

        // Highlight the row according to status value 

        callsSheetStatusCellValueRange.setValue('In Progress'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        statusValue = statusValueCellRange.getValue(); // For acquiring the updated value of the status!
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
        statusValue = statusValueCellRange.getValue(); // For acquiring the updated value of the status!
        fun.setStatusHighlighting(activeSheet, targetRow, callsSheetStatusCellValue);


    } else if ( activeSheetName == 'Calls' && columnValue == 'Call Response' && (currentCellValue == 'Not Answering' || currentCellValue == 'Busy') && callSheetNegativeCounterScore == 4 ) {

        SpreadsheetApp.getActive().toast('Call Sheet Call Response Not Answering Value detected successfully !'); 
        Logger.log('Dead call with not answering or busy response detected successfully!'); 

        const newCallDate = Date.now(); // This should give today's date

        // Is this code transfering the record to History sheet ? check that out! 

        Logger.log(`the new call date is ${newCallDate}`); 

        fun.rescheduleActivity('Call', e, newCallDate, '', currentCellValue); 

        SpreadsheetApp.getActive().toast('Dead call related to Not Answering Call activity completed successfully! Alhumdulillah!');

        // Highlight the row with line through and color according to status value 

        callsSheetStatusCellValueRange.setValue('Unresponsive'); 
        const targetRow = fun.findRowNumber(activeSpreadsheet.getId(), "Sheet1", idOfCompany); 
        statusValue = statusValueCellRange.getValue(); // For acquiring the updated value of the status!
        fun.setStatusHighlighting(activeSheet, targetRow, callsSheetStatusCellValue);

        // Update Sheet1's negativeCounterScore to 4 




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
        statusValue = statusValueCellRange.getValue(); // For acquiring the updated value of the status! 

        fun.setStatusHighlighting(activeSheet, targetRow, statusValue); 


    } else if ( activeSheetName == 'Calls' && columnValue == 'Remarks' && callSheetCallResponse == 'Call Picked Up' && callSheetClientResponse == 'Interested in talking') {

        // reschedule call after 3 days 

        SpreadsheetApp.getActive().toast('Calls Sheet call back detected successfully!');

        const activityDate = Date.now() + 3 * 24 * 3600 * 1000; 

        fun.rescheduleActivity('Call', e, activityDate, '', callResponse); 


        SpreadsheetApp.getActive().toast('Calls Sheet call back completed successfully!');

        // update negative counter value to +1 



        // Move current call to history - I think this is done by the reschedule call function as well . 



    } else if (activeSheetName == 'Calls' && columnValue == 'Call Back Time' && callSheetCallResponse == 'Call Picked Up' && (callsSheetStatusCellValue == 'Lead' || callsSheetStatusCellValue == 'Strong Lead' || callsSheetStatusCellValue == 'Potential Lead' || callsSheetStatusCellValue == 'Opportunity' || callsSheetStatusCellValue == 'Strong Opportunity') && callSheetMeetingValue == 'No') {

        // Transfer record to Inside Sales Executive

        const currentRowRemarks = activeRowArray[0][15]; 

        fun.loadBalancerCompany('Inside Sales Executive', activeRowArray[0][0], e); 

        fun.loadBalancerActivity('Inside Sales Executive', activeRowArray[0][0], e, 'Call', currentRowRemarks); 

        // Update negativeCounterValue to 4 in calls sheet 

        // Update negativeCounter value to 4 in sheet1 

        const relatedCompanyRowNumber = fun.findRowNumber(activeSpreadsheet.getId(), 'Sheet1', activeRowArray[0][1]); 
        const targetRowCellSheet = activeSpreadsheet.getSheetByName('Sheet1'); 
        const negativeCounterRowCellRange = targetRowCellSheet.getRange(relatedCompanyRowNumber, targetRowCellSheet.getLastColumn());
        const sheet1StatusCellValue = activeSheet.getRange(e.range.getRow(), 15).getValue(); 
        // const targetRowCellValue = targetRowCellRange.getValue(); 

        negativeCounterRowCellRange.setValue(4); 

        // Set status highlighting with line through in sheet1 for the related company 


        fun.setStatusHighlighting(targetRowCellSheet, relatedCompanyRowNumber, sheet1StatusCellValue); 



    } else if (activeSheetName == 'Calls' && columnValue == 'Call Back Time' && callSheetCallResponse == 'Call Picked Up' && (callsSheetStatusCellValue == 'Lead' || callsSheetStatusCellValue == 'Strong Lead' || callsSheetStatusCellValue == 'Potential Lead' || callsSheetStatusCellValue == 'Opportunity' || callsSheetStatusCellValue == 'Strong Opportunity') && callSheetMeetingValue == 'Yes') {

        // Transfer record to Inside Sales Executive

        const currentRowRemarks = activeRowArray[0][15]; 

        fun.loadBalancerCompany('Marketing Executive', activeRowArray[0][0], e); 

        fun.loadBalancerActivity('Marketing Executive', activeRowArray[0][0], e, 'Call', currentRowRemarks);

        // Update negativeCounterValue to 4 in calls sheet 

        // Update negativeCounter value to 4 in sheet1 

        const relatedCompanyRowNumber = fun.findRowNumber(activeSpreadsheet.getId(), 'Sheet1', activeRowArray[0][1]); 
        const targetRowCellSheet = activeSpreadsheet.getSheetByName('Sheet1'); 
        const negativeCounterRowCellRange = targetRowCellSheet.getRange(relatedCompanyRowNumber, targetRowCellSheet.getLastColumn());
        const sheet1StatusCellValue = activeSheet.getRange(e.range.getRow(), 15).getValue(); 
        // const targetRowCellValue = targetRowCellRange.getValue(); 

        negativeCounterRowCellRange.setValue(4);

        // Set status highlighting with line through in sheet1 for the related company

        fun.setStatusHighlighting(targetRowCellSheet, relatedCompanyRowNumber, sheet1StatusCellValue);

        // Move the current call to history 



    } else if (activeSheetName == 'Calls' && columnValue == 'Call Back Time' && callSheetCallResponse == 'Call Picked Up' && callsSheetStatusCellValue == 'Call Back At Specified Time') {

        // Reschedule a new call at the specified time 

        const followUpCallDate = activeRowArray[0][17]; 
        const followUpCallTime = activeRowArray[0][18];
        
        fun.rescheduleActivity('Call', e, followUpCallDate, followUpCallTime, callSheetCallResponse); 

        // Move the current call to history 

        // Condition verified successfully - Alhumdulillah! 


    } else if (activeSheetName == 'Calls' && columnValue == 'Client Response' && callSheetClientResponse == 'Call Back Later') {

        // Reschedule a new call 3 days from now 

        const nextFollowUpDate = Date.now() + 3 * 24 * 3600 * 1000; 

        fun.rescheduleActivity('Call', e, nextFollowUpDate, '', callSheetCallResponse); 

        // Update the negative counter value to +1 

        // Move the current call to history 

        // Condition verified successfully! Al Humdulillah! 


    } else if ( activeSheetName == 'Calls' && columnValue == 'Meeting Granted' && callSheetCallResponse == 'Call Picked Up' && activeRowArray[0][9].length != 0) {

        fun.l('status value autocompletion condition detected!'); 


        // Update the value of status based on a function in fun library 

        const newStatusRange = activeSheet.getRange(e.range.getRow(), 15); 

        const newStatusValue = fun.decideStatus(e); 

        fun.l('The status value based on switch statement is', newStatusValue); 

        newStatusRange.setValue(newStatusValue); 


        // Condition verified - Alhumdulillah! 



    }
    
    else {

        SpreadsheetApp.getActive().toast('No condition satisfied!'); 
    }; 



}