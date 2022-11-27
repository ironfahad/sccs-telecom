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
    editTrigger1(); 
    
    Logger.log(status); 

}


function telecomEventProcessing(e) {

    //Get the value of status from the current row 
    const activeRowArray = fun.getEventData(e).activeDataRowArray; 
    Logger.log('The active data row array is '); 
    Logger.log(activeRowArray); 

    const statusValue = activeRowArray[0][22]; // status value 
    const meetingValue = activeRowArray[0][15]; // meeting value
    const callResponse = activeRowArray[0][8]; // call response 
    const negativeCounterScore = activeRowArray[0][28]; // negative score counter 
    const columnValue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange(1, e.range.getColumn()).getValue(); 
    const companyIdLowerCellValue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange(e.range.getRow() + 1, 1).getValue(); 

    Logger.log(`the status value is ${statusValue} & the meeting value is ${meetingValue}, and call response value is ${callResponse} and the negative score counter is ${negativeCounterScore}`); // verified 

    // Execute conditional statements 

    if(statusValue == 'Lead' && meetingValue.length == 0 ) {

        SpreadsheetApp.getActive().toast('Condition for lead activated successfully!'); 

        // Transfer the current row record to the lowest load Inside Sales Executive 

        fun.loadBalancerCompany('Inside Sales Executive', fun.getEventData(e).companyID, e); 
        fun.loadBalancerActivity('Inside Sales Executive', fun.getEventData(e).companyID, e, 'Call'); 
    

    } else if( statusValue == 'Opportunity' && meetingValue == 'yes') {

        SpreadsheetApp.getActive().toast('Condition for meeting activated successfully'); 

        // Transfer the current row record to the lowest load Marketing Executive 

        fun.loadBalancerCompany('Marketing Executive', fun.getEventData(e).companyID, e); 
        fun.loadBalancerActivity('Marketing Executive', fun.getEventData(e).companyID, e, 'Meeting'); 

    } else if ( statusValue == 'Opportunity' && meetingValue.length == 0 ) {

        SpreadsheetApp.getActive().toast('Condition For Opportunity with no meeting activated successfully'); 

        fun.loadBalancerCompany('Inside Sales Executive', fun.getEventData(e).companyID, e); 
        fun.loadBalancerActivity('Inside Sales Executive', fun.getEventData(e).companyID, e, 'Call'); 

    } else if( columnValue == 'Call Response' && companyIdLowerCellValue.length == 0) {

        const SpreadSheetName = SpreadsheetApp.getActiveSpreadsheet().getName(); 
        const SpreadSheetNameArray = SpreadSheetName.split('-'); 
        const campaignId = SpreadSheetNameArray[4]; 
        const employeeId = SpreadSheetNameArray[3]; 

        Logger.log(`the campaign ID is ${campaignId} & the employee id is ${employeeId}`); 

        fun.extractData(campaignId, 10, 'Add'); 


    } else if ( callResponse == 'Busy' && negativeCounterScore < 3) {

        // Reschedule a new call the next day & add 1 score to negative counter score

        SpreadsheetApp.getActive().toast('Busy Call detected successfully!'); 

        const newCallDate = Date.now() + 1 * 24 * 3600 * 1000; // This should give next day's date 

        Logger.log(`the new call date is ${newCallDate}`); 

        fun.rescheduleActivity('Call', e, newCallDate); 

        SpreadsheetApp.getActive().toast('Busy Call activity completed successfully! Alhumdulillah!');

    } else if ( callResponse == 'Not Answering' && negativeCounterScore < 3) {
        
        // Reschedule call after 3 days & add 1 score to negative counter score

        SpreadsheetApp.getActive().toast('Not Answering Call detected successfully!'); 

        const newCallDate = Date.now() + 3 * 24 * 3600 * 1000; // This should give next day's date 

        fun.rescheduleActivity('Call', e, newCallDate, '', callResponse); 

        SpreadsheetApp.getActive().toast('Not Answering Call activity completed successfully! Alhumdulillah!');

    } else if( callResponse == 'Picked Up' && fun.getEventData(e).followUpStatus == 'Call Back Later' && fun.getEventData(e).negativeCounterScore < 3) {

        // Reschedule call after 3 days & add 1 score to negative counter score

        SpreadsheetApp.getActive().toast('Call Back Later detected successfully!'); 

        const newCallDate = Date.now() + 3 * 24 * 3600 * 1000; // This should give next day's date 

        fun.rescheduleActivity('Call', e, newCallDate); 

        SpreadsheetApp.getActive().toast('Call Back Later activity completed successfully! Alhumdulillah!');

    } else if ( callResponse == 'Picked Up' && fun.getEventData(e).followUpStatus == 'Call Back in Specific Time') {

        // Reschedule call at the specific date & time and reduce 1 score from the negative counter score 

        const personCallBackDate = ''; 
        const personCallBackTime = ''; 

        fun.rescheduleActivity('Call', e, personCallBackDate, personCallBackTime); 

        SpreadsheetApp.getActive().toast('Call Back in specific date and time activated successfully!'); 

    } else if ( negativeCounterScore == '3') {

        // select the active row range and grey that contact out 

        const activeRowSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
        const activeRowRange = activeRowSheet.getRange(e.range.getRow(), 1, 1, activeRowSheet.getLastColumn()); 

        activeRowRange.setBackground('grey').setFontStyle('italic'); 

        
    }
    
    else {

        SpreadsheetApp.getActive().toast('No condition satisfied!'); 
    }; 



}