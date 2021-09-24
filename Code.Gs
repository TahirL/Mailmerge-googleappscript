
function sendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //DataSheet
  var dataSheet = ss.getSheetByName("Sheet1")
  //Resonse Sheet
  var responseSheet=ss.getSheetByName("reponsetrack")
  // [START apps_script_gmail_email_data_range]
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, 4);
  // [END apps_script_gmail_email_data_range]

  // [START apps_script_gmail_email_template]
  var templatefile=DriveApp.getFileById("1loPE8XvUBDWEB1pklKSxBVcTvxTkiSS_Il_fC7uyg1Y")

  // [END apps_script_gmail_email_template]

  // [START apps_script_gmail_email_objects]

  var data = dataSheet.getDataRange().getValues()
  var header=data.splice(0,1)[0]
  //email Settings
  var emailText="Please review attached body"
  var emailSubject='New Doc';


  //check Email Column index
  for(var i=0;i<header.length;i++){if(header[i]=="DRM_Email"){var emailCol=i}}

  // [END apps_script_gmail_email_objects]
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  var oldKeys=responseSheet.getRange("A:A").getValues().flat()
  for (var i = 0; i < data.length; ++i) {
    
    // Get a row object
    var rowData = data[i];
    var emailSubject=rowData[1];

    //Check if the response Shet Contains the Key
    if(!oldKeys.includes(rowData[1])){
    // [START apps_script_gmail_email_text]
    // Generate a personalized email.
    var newFile=templatefile.makeCopy();
    newFile.setName(rowData[1] + " Billing Letter")
    var folder = DriveApp.getFolderById("1bn3GaFX_Z4tcEFgU8WHROLbyuJs5p71q")
    newFile.moveTo(folder)
    var id = newFile.getMimeType();
    var doc = newFile.getId();
    //var doc = UrlFetchApp.fetch(url+'download/documents/Export?exportFormat=doc&format=doc&id='+id,googleOAuth_('docs',url)).getBlob();
    // Given a template string, replace markers (for instance ${"First Name"}) with
    fillInTemplateFromObject(newFile, rowData,header);
    

    // [END apps_script_gmail_email_text]
    // [START apps_script_gmail_send_email]
   MailApp.sendEmail(rowData[emailCol], emailSubject, emailText+"\n\n"+newFile.getUrl(),{
          name: 'Telus - DRM Billing Letter Automation',
          cc: 'tahir.latif@telus.com'
          //attachments: [doc]
   })
  
    responseSheet.appendRow([rowData[1],newFile.getUrl()])
    // [END apps_script_gmail_send_email]
    }
  }
}


function fillInTemplateFromObject(newFile, data,header) {
  var email =  DocumentApp.openById(newFile.getId())
  var body=email.getBody();
  
  for (var i = 0; i < header.length; i++) {
    //replaces Each Header with text
    body.replaceText('{{'+header[i]+'}}',data[i])
    
  }
  body=email.getHeader();
  
  for (var i = 0; i < header.length; i++) {
    //replaces Each Header with text
    body.replaceText('{{'+header[i]+'}}',data[i])
    
  }
  body=email.getFooter();
  
  for (var i = 0; i < header.length; i++) {
    //replaces Each Header with text
    body.replaceText('{{'+header[i]+'}}',data[i])
    
  }
  email.saveAndClose();
}
// [END apps_script_gmail_mail_merge]
