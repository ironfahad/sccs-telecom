// WITH THE NAME OF ALLAH, THE MOST MERCIFUL, THE MOST BENEFICIENT
// PRAISE BE TO ALLAH, WHO HAS TAUGHT US WITH A PEN


function editTrigger1() {
    ScriptApp.newTrigger('telecomEventProcessing')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create()
}; 


function telecomEventProcessing(e) {

    //Get the value of status from the current row 

    const statusValue = fun.getEventData(e).companyStatus;
    const meetingValue = fun.getEventData(e).meetingGranted; 
    Logger.log(`the status value is ${statusValue} & the meeting value is ${meetingValue}`); // verified 

    // Execute conditional statements 

    if(statusValue == 'Lead' && meetingValue.length == 0 ) {

        SpreadsheetApp.getActive().toast('Condition for lead activated successfully!'); 

        // Transfer the current row record to the lowest load Inside Sales Executive 

        fun.loadBalancerCompany('Inside Sales Executive', fun.getEventData(e).companyID, e); 

    

    } else if( statusValue == 'Opportunity' && meetingValue == 'yes') {

        SpreadsheetApp.getActive().toast('Condition for meeting activated successfully'); 

        // Transfer the current row record to the lowest load Marketing Executive 

        fun.loadBalancer('Marketing Executive', fun.getEventData(e).companyID); 

    } else {

        SpreadsheetApp.getActive().toast('No condition satisfied!'); 
    }



}