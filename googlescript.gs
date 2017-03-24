function uploadNielsenReports() {
 var d = new Date();
 var today = d.toDateString(); 
 var attachments, attachment, zip, zipFolder, zipName, attachmentName, newSheet, newSheetName, sheetId, rows;
  
 var allSheets = new Array(); 
 
 var emailBody = '<b>Most Recent Nielsen Validation Email Report</b><br/><br/>';
 
  // Nielsen data folder in Ben Upham's Drive
 var nielsenFolder = new Array('0B4irovCPs1GEUU95NTl0T0c5elU');
 
  //Gmail labels in Ben Upham's gmail account
 var nielsenEmails = GmailApp.getUserLabelByName("Nielsen Data Emails");
 var processed = GmailApp.getUserLabelByName("Processed Nielsen Reports"); 
 var threads = nielsenEmails.getThreads(0,20);
 
 for (var i = 0; i < threads.length; i++) {
   var labels = new Array();
   for (var q = 0; q < threads[i].getLabels().length; q++) {
     labels.push(threads[i].getLabels()[q].getName());
   }
   if (threads[i].getMessages()[0].getAttachments() && threads[i].getFirstMessageSubject() == 'Destini Product Locator Order Form' && labels.indexOf('Processed Nielsen Reports') == -1 ) {
     
     var attachments = threads[i].getMessages()[0].getAttachments();
     
     //for .zip attachments
     for (var j = 0; j < attachments.length; j++) {
       if (attachments[j].getContentType() == 'application/zip') {
         zip = attachments[j].copyBlob();
         zipFolder = Utilities.unzip(zip);
         for (var k = 0; k < zipFolder.length; k++) {
           attachment = zipFolder[k].copyBlob();
           var emailData = new Object();
           emailData.name = zipFolder[k].getName() + " " + today;
           newSheet = convertExcel2Sheets(attachment, emailData.name, nielsenFolder);
           emailData.rows = getRows(newSheet);

           allSheets.push(emailData);
           //Logger.log(newSheetName + " " + rows);
           //  Logger.log('invalid: '+zipFolder[k].getName()); 
         }        
         //Logger.log(zipFolder[0].getName());
       } else if (attachments[j].getContentType() == 'application/vnd.ms-excel') {
           attachment = attachments[j].copyBlob();
           var emailData = new Object();
           emailData.name = attachments[j].getName() + " " + today;
           newSheet = convertExcel2Sheets(attachment, emailData.name, nielsenFolder);
           emailData.rows = getRows(newSheet);
           allSheets.push(emailData);
          // Logger.log(newSheetName+' - rows as determined by getRows: '+ rows);
       } 
     }
   
   threads[i].addLabel(processed);

   }
 }
  //sort the data alphabetically
  allSheets.sort(function(a, b) {
    var nameA = a.name.toUpperCase(); // ignore upper and lowercase
    var nameB = b.name.toUpperCase(); // ignore upper and lowercase
    if (nameA < nameB) {
      return -1;
    }
    if (nameA > nameB) {
      return 1;
    }
    
    // names must be equal
    return 0;
  });
  //create email body html
  for (var z = 0; z < allSheets.length; z++)
    if (allSheets[z].name.indexOf('INVALID') > 0) {
      emailBody += '<a href="'+ newSheet.getUrl() + '">' + allSheets[z].name + '</a> <span style="color:red">Invalid</span> UPCs: ' + allSheets[z].rows + '<br/><br/>'; 
    } else {
      emailBody += '<a href="'+ newSheet.getUrl() + '">' + allSheets[z].name + '</a> <span style="color:green">Valid</span> UPCs: ' + allSheets[z].rows + '<br/><br/>'; 
    }
  //Logger.log(emailBody);
  if (allSheets.length > 0) {
    GmailApp.sendEmail('bupham@destinilocators.com,jeff@destinilocators.com', 'Nielsen Report', emailBody, { htmlBody: emailBody});
  }
}


//because the excel from Nielsen sometimes has blank rows filled with spaces
function getRows(sheet){
  var range = sheet.getDataRange();
  var values = range.getValues();
  var rows = 0;
  for (var i=0; i < values.length; i++) {
    // count rows that have actual data in them
    
    if (typeof values[i][0] == 'number') {
      //Logger.log(typeof values[i][0]);
      rows++;
    }
  }
  //Logger.log(rows);
  return rows;
  //Logger.log('rows '+rows);
  //Logger.log('values '+values);
  
  //Logger.log('values length '+values.length);
  //Logger.log(typeof ss.getLastRow());
}
  
  
//var ssNew = SpreadsheetApp.create("Finances");
//application/vnd.ms-excel

/**
 * Convert Excel file to Sheets
 * @param {Blob} excelFile The Excel file blob data; Required
 * @param {String} filename File name on uploading drive; Required
 * @param {Array} arrParents Array of folder ids to put converted file in; Optional, will default to Drive root folder
 * @return {Spreadsheet} Converted Google Spreadsheet instance
 **/

function convertExcel2Sheets(excelFile, filename, arrParents) {
  
  var parents  = arrParents || []; // check if optional arrParents argument was provided, default to empty array if not
  
  //if ( !parents.isArray ) parents = []; // make sure parents is an array, reset to empty array if not
 
  // Parameters for Drive API Simple Upload request (see https://developers.google.com/drive/web/manage-uploads#simple)
  var uploadParams = {
    method:'post',
    contentType: 'application/vnd.ms-excel', // works for both .xls and .xlsx files
    contentLength: excelFile.getBytes().length,
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    payload: excelFile.getBytes()
  };
  
  // Upload file to Drive root folder and convert to Sheets
  var uploadResponse = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true', uploadParams);
    
  // Parse upload&convert response data (need this to be able to get id of converted sheet)
  var fileDataResponse = JSON.parse(uploadResponse.getContentText());

  // Create payload (body) data for updating converted file's name and parent folder(s)
  var payloadData = {
    title: filename, 
    parents: []
  };
  if ( parents.length ) { // Add provided parent folder(s) id(s) to payloadData, if any
    for ( var i=0; i<parents.length; i++ ) {
      try {
        var folder = DriveApp.getFolderById(parents[i]); // check that this folder id exists in drive and user can write to it
        payloadData.parents.push({id: parents[i]});
      }
      catch(e){} // fail silently if no such folder id exists in Drive
    }
  }
  // Parameters for Drive API File Update request (see https://developers.google.com/drive/v2/reference/files/update)
  var updateParams = {
    method:'put',
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    contentType: 'application/json',
    payload: JSON.stringify(payloadData)
  };
  
  // Update metadata (filename and parent folder(s)) of converted sheet
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/'+fileDataResponse.id, updateParams);
  
  return SpreadsheetApp.openById(fileDataResponse.id);
}

