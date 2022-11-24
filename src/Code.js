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

    const statusValue = fun.getEventData(e).companyStatus;
    const meetingValue = fun.getEventData(e).meetingGranted; 
    const columnValue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange(1, e.range.getColumn()).getValue(); 
    const callResponseLowerCellValue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange(e.range.getRow() + 1, e.range.getColumn()).getValue(); 

    Logger.log(`the status value is ${statusValue} & the meeting value is ${meetingValue}`); // verified 

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

    } else if( columnValue == 'Call Response' && callResponseLowerCellValue.length == 0) {

        

        // Find the campaign ID associated with current user 

        // Get the operations sheet of the strategic department 

        // Get the cell containing the URL of the duplicate target list 

        // Extract the file ID of the duplicate target list 

        // Get the Source Data 

        // Get the target Data 

        // Delete 25 records from the target sheet 

        // Copy 25 records to the target sheet 

        // Delete 25 records from the duplicate target list sheet 



    }
    
    else {

        SpreadsheetApp.getActive().toast('No condition satisfied!'); 
    }; 



}