
// upload document into google spreadsheet
    // and put link to it into current cell

function onOpen(e) {
      var ss = SpreadsheetApp.getActiveSpreadsheet()
      var menuEntries = [];
      menuEntries.push({name: "Attach File...", functionName: "doGet"});
      menuEntries.push({name: "Request Peer Review", functionName: "Review_sendEmail"});
      menuEntries.push({name: "Back To Developer", functionName: "Back_to_Developer"});
      menuEntries.push({name: "Request Lead Review", functionName: "Review_Lead"});
      menuEntries.push({name: "Review Completed", functionName: "Review_done"});
      menuEntries.push({name: "Show Only In-Progress Status", functionName: "hideCompleted"});
      menuEntries.push({name: "Show All", functionName: "showAll"});
      ss.addMenu("", menuEntries);
}

function test()
{
  /*var files = DriveApp.getFolderById('FOLDERID').getFiles();
   for(var i=0; i < files.length; i++)
   {
      if(files[i].getName()=="SQL.txt")
      {
        Browser.msgBox(files[i].getUrl());
        break;
      }
   }*/
   
      var emailAddress = "EMAILID";
      var subject = "Test Message to see the format"
      var message =  "<html><head></head><body>"+
                  "<div style='width: 560px; height: 580px; background-color: #F1FAFF;border-radius: 10px;'><center><br /><br />"+
                  "<div style='border-radius: 10px;border: 1px solid #7DACC6;font-family: Times New Roman;font-size: 14px;font-weight: normal;padding: 25px 25px 15px 25px;width: 460px;height: 410px;background-color: White;text-align:left;'>"+
                  "<img src='" + "BACKGROUND IMAGE URL" + "' style='border-style: none' />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
                  "<br />"+
                  "<div style='font-size: 24px;font-weight: strong;padding: 45px 25px 15px 25px;'> <center>  Peer Review</center> </div>"+
                  "<br /><br />"+ 
                  "Dear "+ "Tilak Dhar" + "," +
                    "<br /><br /><br />" + " Please review the following component. Below are the details:"+
                  "<br /><br />" + "Name:  "+ "<strong>" +"Test Name" +" </strong>" +                
                  "<br /><br />" + "Type:  "+ "<b>" + "Test Type"+"</b>"+
                  "<br /><br />" +"Sheet:  "+ "<b>" + "Test Sheet"+"</b>"+
                  "<br /><br />" +"Comment:  "+ "<b>" + "Test Comment"+"</b>"+
                  "<p>" +"<br>" +"<br>" +"You can access the spreadsheet "+ "<A HREF='" + "Test URL" + "'>here</A>"+
                  "<br>" +"<br>" +"<br>" +
                    
                  "Thanks,"+"<p>"+"Tilak"+ " " +"Dhar"+
                  "</div></center></div></body></html>";
  
   MailApp.sendEmail(emailAddress, subject, message, {htmlBody: message});
}

function doGet(e) {
      var app = UiApp.createApplication().setTitle(" - Add Attachment");
      SpreadsheetApp.getActiveSpreadsheet().show(app);
      var form = app.createFormPanel().setId('frm').setEncoding('multipart/form-data');
      var formContent = app.createVerticalPanel();
      form.add(formContent);  
      formContent.add(app.createFileUpload().setName('thefile'));

      // these parameters need to be passed by form
      // in doPost() these cannot be found out anymore
      formContent.add(app.createHidden("activeCell", SpreadsheetApp.getActiveRange().getA1Notation()));
      formContent.add(app.createHidden("activeSheet", SpreadsheetApp.getActiveSheet().getName()));
      formContent.add(app.createHidden("activeSpreadsheet", SpreadsheetApp.getActiveSpreadsheet().getId()));
      formContent.add(app.createSubmitButton('Submit'));
      app.add(form);
      SpreadsheetApp.getActiveSpreadsheet().show(app);
      return app;
}

function doPost(e) {
      var app = UiApp.getActiveApplication();
      app.createLabel('saving...');
      var fileBlob = e.parameter.thefile;
  // Changed by Tilak for testing:
  // var doc = DriveApp.getFolderById('FOLDERID').createFile(fileBlob);
  var doc = DriveApp.getFolderById('FOLDERID').createFile(fileBlob);
      var label = app.createLabel('File uploaded successfully');

      // write value into current cell
      var value = 'hyperlink("' + doc.getUrl() + '","' + doc.getName() + '")'
      var activeSpreadsheet = e.parameter.activeSpreadsheet;
      var activeSheet = e.parameter.activeSheet;
      var activeCell = e.parameter.activeCell;
      var label = app.createLabel('File uploaded successfully');
      app.add(label);
      SpreadsheetApp.openById(activeSpreadsheet).getSheetByName(activeSheet).getRange(activeCell).setFormula(value);
      app.close();
      return app;
}


  function hideCompleted()
  {
    var sheet = SpreadsheetApp.getActiveSheet();
    var maxRows = sheet.getMaxRows();
    sheet.showRows(1, maxRows);
    for(var i=1; i< maxRows+1; i++){      
       if(sheet.getRange("A" + Number(i)).getValue() == 'Completed'){
         sheet.hideRows(i);
      }
    } 
  }

  function showAll()
  {
    var sheet = SpreadsheetApp.getActiveSheet();     
    var maxRows = sheet.getMaxRows(); 
    sheet.showRows(1, maxRows);
  }

function Review_sendEmail() {
  var sheet = SpreadsheetApp.getActiveSheet();
   //SpreadsheetApp.openById("FOLDERID").toast("My message to the end user.","Message Title");
  //var sheet = SpreadsheetApp.openById("FOLDERID").getActiveSheet();
  
  var cell =  sheet.getActiveCell();
  var cellR = cell.getRow();
  var cellC = cell.getColumn();
  var cellValue = cell.getValue();
  var rowIndex = cell.getRowIndex();
  
  if(sheet.getRange("K" + Number(cell.getRowIndex())).getValue().length==0 && cell.getColumnIndex()==11)
  {
       var myMessage = Browser.msgBox('Comments Not Saved',
      'You have not saved the last entered comment. Please click outside the comment box to save it.' +
      ' ',
       Browser.Buttons.OK);
      if (myMessage == 'ok') {
        return;
      }
  }
  
  //Logger.log(Session.getEffectiveUser());
  
  if(Session.getEffectiveUser()!=sheet.getRange("O" + Number(cell.getRowIndex())).getValue())
  {
    var myMessage = Browser.msgBox('Wrong row selection',
      'You have selected a different row than you last edited.' +
      ' Do you want to continue ?',
      Browser.Buttons.YES_NO);
    if (myMessage == 'no') {
      return;
    }
    else
    {
      //Browser.msgBox(cell.getRowIndex());
      rowIndex = cell.getRowIndex();
    }
  }
  else
  {
     //rowIndex = sheet.getRange("Z1").getValue();
     //rowIndex = cell.getRowIndex();
    if(cell.getRowIndex()!=sheet.getRange("Z1").getValue())
    {
       var myMessage = Browser.msgBox('Wrong row selection',
      'You have either selected a different row than you last edited or someone else has edited the sheet.' +
      ' Do you want to continue ?',
       Browser.Buttons.YES_NO);
      if (myMessage == 'no') {
        return;
      }
      else
      {
        rowIndex = cell.getRowIndex();
      }
    }
  }
  var colIndex = cell.getColumnIndex();
  
 
  //var dataRange = sheet.getRange(rowIndex+1,1,1,colIndex);
  var dataRange = sheet.getRange(rowIndex,1,1,12);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
 for (i in data) {
    var row = data[i];


    var names = row[7].split(" ");
    var emailAddress = names[0].toLowerCase() +"."+names[1].toLowerCase()+"@gmail.com";
    if(emailAddress=="EMAILID")
    {
      emailAddress = "EMAILID";
    }
   
   var senderNameTemp = (Session.getEffectiveUser().toString().split("@"));
   var senderName = senderNameTemp[0].split(".");
   var subject = sheet.getName() + " Peer Review - Please review "+ row[1] + " - " + row[2];
   
   var ssURL = 'SHEET URL';
  
  // Changed by Tilak for testing
   // var files = DriveApp.getFolderById('FOLDERID').getFiles();
   var files = DriveApp.getFolderById('FOLDERID').getFiles();
  var attchURL = row[3];
   for(var p=0; p < files.length; p++)
   {
      if(files[p].getName()==row[3].toString())
      {
        //Browser.msgBox(files[p].getUrl());
        attchURL = files[p].getUrl();
        break;
      }
   }
   
   var message =  "<html><head></head><body>"+
                  "<div style='width: 560px; height: 580px; background-color: #F1FAFF;border-radius: 10px;'><center><br /><br />"+
                  "<div style='border-radius: 10px;border: 1px solid #7DACC6;font-family: Times New Roman;font-size: 14px;font-weight: normal;padding: 25px 25px 15px 25px;width: 460px;height: 410px;background-color: White;text-align:left;'>"+
                  "<img src='" + "BACKGROUND IMAGE URL" + "' style='border-style: none' />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
                  "<br />"+
                  "<div style='font-size: 24px;font-weight: strong;padding: 45px 25px 15px 25px;'> <center>  Peer Review</center> </div>"+
                  "<br /><br />"+ 
                  "Dear "+ row[7] + "," +
                  "<br /><br /><br />" + " Please review the following component. Below are the details:"+
                  "<br /><br />" + "Name:  "+ "<b>" +row[2] +"</b>" +                
                  "<br /><br />" + "Type:  "+ "<b>" + row[1]+"</b>"+
                  "<br /><br />" +"Attachment:  "+ "<b>" + "<A HREF='" + attchURL + "'>"+row[3]+"</A>" +"</b>"+
                  "<br /><br />" +"Sheet:  "+ "<b>" + sheet.getName()+"</b>"+
                  "<br /><br />" +"Comment:  "+ "<b>" + row[10]+"</b>"+
                  "<p>" +"<br>" +"<br>" +"You can access the spreadsheet "+ "<A HREF='" + ssURL + "'>here</A>"+
                  "<br>" +"<br>" +"<br>" +
                    
                  "Thanks,"+"<p>"+senderName[0].toUpperCase()+ " " +senderName[1].toUpperCase()+
                  "</div></center></div></body></html>";
 
    MailApp.sendEmail(emailAddress, subject, message, {htmlBody: message});
   
    var statusCheck = "A" + Number(rowIndex);
     sheet.getRange(statusCheck).setValue("Peer Review");
   
    var assignedToCheck = "F" + Number(rowIndex);
    sheet.getRange(assignedToCheck).setValue(row[7]);
   
    var historyCommentsCol = "L" + Number(rowIndex);
    var historyCommentsOld = sheet.getRange(historyCommentsCol).getValue();
    var todaysDate = new Date();
    var historyComments =  "Peer Review"+ "\n"+todaysDate + " - " + "\n"+ Session.getEffectiveUser()+ " added the following comments " + "\n"+row[10]+
                           "\n"+"****************"+"\n"+historyCommentsOld;
    sheet.getRange(historyCommentsCol).setValue(historyComments);
   
    var historyCheck = "N" + Number(rowIndex);
    //sheet.setActiveSelection(historyCheck);
    var historyDetailsOld = sheet.getRange(historyCheck).getValue();//sheet.getActiveCell().getValue();
    var todaysDate = new Date();
    var historyDetails =  "Peer Review"+ "\n"+todaysDate + " - " + "\n"+ Session.getEffectiveUser() + " requested review from " + row[7] + "\n"+ "Status - "+row[0] + 
                          ", Location - "+row[3]+ ", Version - "+row[4]+ ", Assigned To - "+row[5]+ 
                          ", Peer Reviewer - "+row[7]+", Lead Reviewer - "+row[8]+ "\n"+ "Comments - "+row[10]+"\n"+"\n"+"****************"+"\n"+historyDetailsOld;
    //sheet.getActiveCell().setValue(historyDetails);
    sheet.getRange(historyCheck).setValue(historyDetails);
   
   //Making the comments cell as null in the end.
   sheet.getRange("K" + Number(rowIndex)).setValue(null);
   
   //Set Peer Review value
   sheet.getRange("P" + Number(rowIndex)).setValue(row[7]);
   
  }
}


function Back_to_Developer() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell =  sheet.getActiveCell();
  var cellR = cell.getRow();
  var cellC = cell.getColumn();
  var cellValue = cell.getValue();
  var rowIndex = cell.getRowIndex();
  
  if(sheet.getRange("K" + Number(cell.getRowIndex())).getValue().length==0 && cell.getColumnIndex()==11)
  {
       var myMessage = Browser.msgBox('Comments Not Saved',
      'You have not saved the last entered comment. Please click outside the comment box to save it.' +
      ' ',
       Browser.Buttons.OK);
      if (myMessage == 'ok') {
        return;
      }
  }
  
  if(Session.getEffectiveUser()!=sheet.getRange("O" + Number(cell.getRowIndex())).getValue())
  {
    var myMessage = Browser.msgBox('Wrong row selection',
      'You have selected a different row than you last edited.' +
      ' Do you want to continue ?',
      Browser.Buttons.YES_NO);
    if (myMessage == 'no') {
      return;
    }
    else
    {
      rowIndex = cell.getRowIndex();
    }
  }
  else
  {
    if(cell.getRowIndex()!=sheet.getRange("Z1").getValue())
    {
       var myMessage = Browser.msgBox('Wrong row selection',
      'You have either selected a different row than you last edited or someone else has edited the sheet.' +
      ' Do you want to continue ?',
       Browser.Buttons.YES_NO);
      if (myMessage == 'no') {
        return;
      }
      else
      {
        rowIndex = cell.getRowIndex();
      }
    }
  }
  var colIndex = cell.getColumnIndex();
  

  var dataRange = sheet.getRange(rowIndex,1,1,12);
  var data = dataRange.getValues();
 for (i in data) {
    var row = data[i];

     if(!(sheet.getRange("L" + Number(rowIndex)).getValue().indexOf("Peer Review")>-1))
     {
       var myMessage = Browser.msgBox('Incorrect Selection',
      'You want to Send Back to Developer without Peer Review being done.' +
      ' Do you want to continue ?',
       Browser.Buttons.YES_NO);
      if (myMessage == 'no') {
        return;
      }
     }

    var names = row[6].split(" ");
    var emailAddress = names[0].toLowerCase() +"."+names[1].toLowerCase()+"@gmail.com";
    if(emailAddress=="EMAILID")
    {
      emailAddress = "EMAILID";
    }
    
   var senderNameTemp = (Session.getEffectiveUser().toString().split("@"));
   var senderName = senderNameTemp[0].split(".");
   
   var subject = sheet.getName() + " - Back to Developer - Review Comments for "+ row[1] + " - " + row[2];
   
   //Changed by Tilak for testing 
   
   var ssURL = 'SPREADSHEET URL';
   
   
   // Changed by Tilak for testing
   
   var files = DriveApp.getFolderById('FOLDERID').getFiles();
   
   var attchURL = row[3];
   for(var p=0; p < files.length; p++)
   {
      if(files[p].getName()==row[3].toString())
      {
        //Browser.msgBox(files[p].getUrl());
        attchURL = files[p].getUrl();
        break;
      }
   }
   
   var message =  "<html><head></head><body>"+
                  "<div style='width: 560px; height: 580px; background-color: #F1FAFF;border-radius: 10px;'><center><br /><br />"+
                  "<div style='border-radius: 10px;border: 1px solid #7DACC6;font-family: Times New Roman;font-size: 14px;font-weight: normal;padding: 25px 25px 15px 25px;width: 460px;height: 410px;background-color: White;text-align:left;'>"+
                  "<img src='" + "BACKGROUND IMAGE URL" + "' style='border-style: none' />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
                  "<br />"+
                  "<div style='font-size: 24px;font-weight: strong;padding: 45px 25px 15px 25px;'> <center>  - Back To Developer</center> </div>"+
                  "<br /><br />"+ 
                  "Dear "+ row[6] + "," +
                  "<br /><br /><br />" + "I have reviewed the following component. Please see my comments below."+
                  "<br /><br />" + "Name:  "+ "<b>" +row[2] +"</b>" +                
                  "<br /><br />" + "Type:  "+ "<b>" + row[1]+"</b>"+
                  "<br /><br />" +"Attachment:  "+ "<b>" + "<A HREF='" + attchURL + "'>"+row[3]+"</A>" +"</b>"+
                  "<br /><br />" +"Sheet:  "+ "<b>" + sheet.getName()+"</b>"+
                  "<br /><br />" +"Comment:  "+ "<b>" + row[10]+"</b>"+
                  "<p>" +"<br>" +"<br>" +"You can access the spreadsheet "+ "<A HREF='" + ssURL + "'>here</A>"+
                  "<br>" +"<br>" +"<br>" +
                    
                  "Thanks,"+"<p>"+senderName[0].toUpperCase()+ " " +senderName[1].toUpperCase()+
                  "</div></center></div></body></html>";
 
    MailApp.sendEmail(emailAddress, subject, message, {htmlBody: message});
   
    var statusCheck = "A" + Number(rowIndex);
     sheet.getRange(statusCheck).setValue("In Progress");
   
    var assignedToCheck = "F" + Number(rowIndex);
    sheet.getRange(assignedToCheck).setValue(row[6]);
   
    var historyCommentsCol = "L" + Number(rowIndex);
    var historyCommentsOld = sheet.getRange(historyCommentsCol).getValue();
    var todaysDate = new Date();
    var historyComments =  "Back To Developer"+ "\n"+todaysDate + " - " + "\n"+ Session.getEffectiveUser()+ " added the following comments " + "\n"+row[10]+
                           "\n"+"****************"+"\n"+historyCommentsOld;
    sheet.getRange(historyCommentsCol).setValue(historyComments);
   
    var historyCheck = "N" + Number(rowIndex);
    //sheet.setActiveSelection(historyCheck);
    var historyDetailsOld = sheet.getRange(historyCheck).getValue();//sheet.getActiveCell().getValue();
    var todaysDate = new Date();
    var historyDetails =  "Back To Developer"+ "\n"+todaysDate + " - " + "\n"+ Session.getEffectiveUser() + " requested review from " + row[7] + "\n"+ "Status - "+row[0] + 
                          ", Location - "+row[3]+ ", Version - "+row[4]+ ", Assigned To - "+row[5]+ 
                          ", Peer Reviewer - "+row[7]+", Lead Reviewer - "+row[8]+ "\n"+ "Comments - "+row[10]+"\n"+"\n"+"****************"+"\n"+historyDetailsOld;
    //sheet.getActiveCell().setValue(historyDetails);
    sheet.getRange(historyCheck).setValue(historyDetails);
   
   //Making the comments cell as null in the end.
   sheet.getRange("K" + Number(rowIndex)).setValue(null);
   
  }
}

function Review_Lead() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var cell =  sheet.getActiveCell();
  var cellR = cell.getRow();
  var cellC = cell.getColumn();
  var cellValue = cell.getValue();
  var rowIndex = cell.getRowIndex();
  
  if(sheet.getRange("K" + Number(cell.getRowIndex())).getValue().length==0 && cell.getColumnIndex()==11)
  {
       var myMessage = Browser.msgBox('Comments Not Saved',
      'You have not saved the last entered comment. Please click outside the comment box to save it.' +
      ' ',
       Browser.Buttons.OK);
      if (myMessage == 'ok') {
        return;
      }
  }
  
  if(Session.getEffectiveUser()!=sheet.getRange("O" + Number(cell.getRowIndex())).getValue())
  {
    var myMessage = Browser.msgBox('Wrong row selection',
      'You have selected a different row than you last edited.' +
      ' Do you want to continue ?',
      Browser.Buttons.YES_NO);
    if (myMessage == 'no') {
      return;
    }
    else
    {
      rowIndex = cell.getRowIndex();
    }
  }
  else
  {

    if(cell.getRowIndex()!=sheet.getRange("Z1").getValue())
    {
       var myMessage = Browser.msgBox('Wrong row selection',
      'You have either selected a different row than you last edited or someone else has edited the sheet.' +
      ' Do you want to continue ?',
       Browser.Buttons.YES_NO);
      if (myMessage == 'no') {
        return;
      }
      else
      {
        rowIndex = cell.getRowIndex();
      }
    }
  }
  
   if(sheet.getRange("I" + Number(cell.getRowIndex())).getValue().length==0)
  {
       var myMessage = Browser.msgBox('No Reviewer Selected',
      'You have not selected any Lead Reviewer. Please select a Lead Reviewer from the list.' +
      ' ',
       Browser.Buttons.OK);
      if (myMessage == 'ok') {
        return;
      }
  }
  

  var colIndex = cell.getColumnIndex();
  
  var dataRange = sheet.getRange(rowIndex,1,1,12);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
 for (i in data) {
    var row = data[i];

     if(sheet.getRange("P" + Number(rowIndex)).getValue()==row[7])
     {
       var myMessage = Browser.msgBox('Select Reviewer',
      'You have not changed the Reviewer. This reviewer already did the Peer Review' +
      ' Do you want to continue ?',
       Browser.Buttons.YES_NO);
      if (myMessage == 'no') {
        return;
      }
     }
   
     if(!(sheet.getRange("L" + Number(rowIndex)).getValue().indexOf("Peer Review")>-1))
     {
       var myMessage = Browser.msgBox('Incorrect Selection',
      'You are requesting for Lead Review without Peer Review being done.' +
      ' Do you want to continue ?',
       Browser.Buttons.YES_NO);
      if (myMessage == 'no') {
        return;
      }
     }
   
    var names = row[8].split(" ");
    var emailAddress = names[0].toLowerCase() +"."+names[1].toLowerCase()+"@gmail.com";
    if(emailAddress=="EMAILID")
    {
      emailAddress = "EMAILID";
    }
   
   var senderNameTemp = (Session.getEffectiveUser().toString().split("@"));
   var senderName = senderNameTemp[0].split(".");
   var subject = sheet.getName() + " - Lead Review - Please review "+ row[1] + " - " + row[2];
   
   // Changed by Tilak for testing
   
   var ssURL = 'SPREADSHEET URL';
   
   
   // Changed by Tilak for testing 
   
   var files = DriveApp.getFolderById('FOLDERID').getFiles();
   var attchURL = row[3];
   for(var p=0; p < files.length; p++)
   {
      if(files[p].getName()==row[3].toString())
      {
        //Browser.msgBox(files[p].getUrl());
        attchURL = files[p].getUrl();
        break;
      }
   }
   
   var message =  "<html><head></head><body>"+
                  "<div style='width: 560px; height: 580px; background-color: #F1FAFF;border-radius: 10px;'><center><br /><br />"+
                  "<div style='border-radius: 10px;border: 1px solid #7DACC6;font-family: Times New Roman;font-size: 14px;font-weight: normal;padding: 25px 25px 15px 25px;width: 460px;height: 410px;background-color: White;text-align:left;'>"+
                  "<img src='" + "BACKGROUND IMAGE URL" + "' style='border-style: none' />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
                  "<br />"+
                  "<div style='font-size: 24px;font-weight: strong;padding: 45px 25px 15px 25px;'> <center>  - Lead Review</center> </div>"+
                  "<br /><br />"+ 
                  "Dear "+ row[8] + "," +
                  "<br /><br /><br />" + " Please do the Lead Review of the following component. Below are the details:"+
                  "<br /><br />" + "Name:  "+ "<b>" +row[2] +"</b>" +                
                  "<br /><br />" + "Type:  "+ "<b>" + row[1]+"</b>"+
                  "<br /><br />" +"Attachment:  "+ "<b>" + "<A HREF='" + attchURL + "'>"+row[3]+"</A>" +"</b>"+
                  "<br /><br />" +"Sheet:  "+ "<b>" + sheet.getName()+"</b>"+
                   "<br /><br />" +"Peer Review Done by:  "+ "<b>" + row[7]+"</b>"+
                  "<br /><br />" +"Comment:  "+ "<b>" + row[10]+"</b>"+
                  "<p>" +"<br>" +"<br>" +"You can access the spreadsheet "+ "<A HREF='" + ssURL + "'>here</A>"+
                  "<br>" +"<br>" +"<br>" +
                    
                  "Thanks,"+"<p>"+senderName[0].toUpperCase()+ " " +senderName[1].toUpperCase()+
                  "</div></center></div></body></html>";
 
    MailApp.sendEmail(emailAddress, subject, message, {htmlBody: message});
   
    var statusCheck = "A" + Number(rowIndex);
    sheet.getRange(statusCheck).setValue("Lead Review");
   
    var assignedToCheck = "F" + Number(rowIndex);
    sheet.getRange(assignedToCheck).setValue(row[7]);
   
    var historyCommentsCol = "L" + Number(rowIndex);
    var historyCommentsOld = sheet.getRange(historyCommentsCol).getValue();
    var todaysDate = new Date();
    var historyComments = "Lead Review"+ "\n"+todaysDate + " - " + "\n"+ Session.getEffectiveUser()+ " added the following comments " + "\n"+row[10]+
                           "\n"+"****************"+"\n"+historyCommentsOld;
    sheet.getRange(historyCommentsCol).setValue(historyComments);
   
    var historyCheck = "N" + Number(rowIndex);
    //sheet.setActiveSelection(historyCheck);
    var historyDetailsOld = sheet.getRange(historyCheck).getValue();//sheet.getActiveCell().getValue();
    var todaysDate = new Date();
    var historyDetails =  "Lead Review"+ "\n"+todaysDate + " - " + "\n"+ Session.getEffectiveUser() + " requested review from " + row[8] + "\n"+ "Status - "+row[0] + 
                          ", Location - "+row[3]+ ", Version - "+row[4]+ ", Assigned To - "+row[5]+ 
                          ", Peer Reviewer - "+row[7]+", Lead Reviewer - "+row[8]+ "\n"+ "Comments - "+row[10]+"\n"+"\n"+"****************"+"\n"+historyDetailsOld;
    sheet.getRange(historyCheck).setValue(historyDetails);
   
   //Making the comments cell as null in the end.
   sheet.getRange("K" + Number(rowIndex)).setValue(null);
   
  }
}


function Review_done() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell =  sheet.getActiveCell();
  var cellR = cell.getRow();
  var cellC = cell.getColumn();
  var cellValue = cell.getValue();
  var rowIndex = cell.getRowIndex();
    
  if(sheet.getRange("K" + Number(cell.getRowIndex())).getValue().length==0 && cell.getColumnIndex()==11)
  {
       var myMessage = Browser.msgBox('Comments Not Saved',
      'You have not saved the last entered comment. Please click outside the comment box to save it.' +
      ' ',
       Browser.Buttons.OK);
      if (myMessage == 'ok') {
        return;
      }
  }
    
  if(Session.getEffectiveUser()!=sheet.getRange("O" + Number(cell.getRowIndex())).getValue())
  {
    var myMessage = Browser.msgBox('Wrong row selection',
      'You have selected a different row than you last edited.' +
      ' Do you want to continue ?',
      Browser.Buttons.YES_NO);
    if (myMessage == 'no') {
      return;
    }
    else
    {
      rowIndex = cell.getRowIndex();
    }
  }
  else
  {
    if(cell.getRowIndex()!=sheet.getRange("Z1").getValue())
    {
       var myMessage = Browser.msgBox('Wrong row selection',
      'You have either selected a different row than you last edited or someone else has edited the sheet.' +
      ' Do you want to continue ?',
       Browser.Buttons.YES_NO);
      if (myMessage == 'no') {
        return;
      }
      else
      {
        rowIndex = cell.getRowIndex();
      }
    }
  }
  var colIndex = cell.getColumnIndex();
  

  var dataRange = sheet.getRange(rowIndex,1,1,12);
  var data = dataRange.getValues();
 for (i in data) {
    var row = data[i];

     if(!(sheet.getRange("L" + Number(rowIndex)).getValue().indexOf("Lead Review")>-1))
     {
       var myMessage = Browser.msgBox('Incorrect Selection',
      'You are completing this review without Lead Review being done.' +
      ' Do you want to continue ?',
       Browser.Buttons.YES_NO);
      if (myMessage == 'no') {
        return;
      }
     }

    var names = row[6].split(" ");
    var emailAddress = names[0].toLowerCase() +"."+names[1].toLowerCase()+"@gmail.com";
    if(emailAddress=="EMAILID")
    {
      emailAddress = "EMAILID";
    }
    
   var senderNameTemp = (Session.getEffectiveUser().toString().split("@"));
   var senderName = senderNameTemp[0].split(".");
   
   var subject = sheet.getName() + " - Review Completed for "+ row[1] + " - " + row[2];
   
   // Changed by Tilak for testing 
   
   var ssURL = 'SPREADSHEET URL';
   
   // Changed by Tilak for testing 
   
   var files = DriveApp.getFolderById('FOLDERID').getFiles();
   var attchURL = row[3];
   for(var p=0; p < files.length; p++)
   {
      if(files[p].getName()==row[3].toString())
      {
        //Browser.msgBox(files[p].getUrl());
        attchURL = files[p].getUrl();
        break;
      }
   }
   
   var message =  "<html><head></head><body>"+
                  "<div style='width: 560px; height: 580px; background-color: #F1FAFF;border-radius: 10px;'><center><br /><br />"+
                  "<div style='border-radius: 10px;border: 1px solid #7DACC6;font-family: Times New Roman;font-size: 14px;font-weight: normal;padding: 25px 25px 15px 25px;width: 460px;height: 410px;background-color: White;text-align:left;'>"+
                  "<img src='" + "BACKGROUND IMAGE URL" + "' style='border-style: none' />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
                  "<br />"+
                  "<div style='font-size: 24px;font-weight: strong;padding: 45px 25px 15px 25px;'> <center>  - Review Complete</center> </div>"+
                  "<br /><br />"+ 
                  "Dear "+ row[6] + "," +
                  "<br /><br /><br />" + " The review has been completed for the following component."+
                  "<br /><br />" + "Name:  "+ "<b>" +row[2] +"</b>" +                
                  "<br /><br />" + "Type:  "+ "<b>" + row[1]+"</b>"+
                   "<br /><br />" +"Attachment:  "+ "<b>" + "<A HREF='" + attchURL + "'>"+row[3]+"</A>" +"</b>"+
                  "<br /><br />" +"Sheet:  "+ "<b>" + sheet.getName()+"</b>"+
                  "<br /><br />" +"Comment:  "+ "<b>" + row[10]+"</b>"+
                  "<br /><br />" +"<br>" +"<br>" +"You can access the spreadsheet "+ "<A HREF='" + ssURL + "'>here</A>"+
                  "<br>" +"<br>" +"<br>" +
                    
                  "Thanks,"+"<p>"+senderName[0].toUpperCase()+ " " +senderName[1].toUpperCase()+
                  "</div></center></div></body></html>";
 
    MailApp.sendEmail(emailAddress, subject, message, {htmlBody: message});
   
    var statusCheck = "A" + Number(rowIndex);
     sheet.getRange(statusCheck).setValue("Completed");
   
    var assignedToCheck = "F" + Number(rowIndex);
    sheet.getRange(assignedToCheck).setValue("None");
   
    var historyCommentsCol = "L" + Number(rowIndex);
    var historyCommentsOld = sheet.getRange(historyCommentsCol).getValue();
    var todaysDate = new Date();
    var historyComments =  "Completed"+ "\n"+todaysDate + " - " + "\n"+ Session.getEffectiveUser()+ " added the following comments " + "\n"+row[10]+
                           "\n"+"****************"+"\n"+historyCommentsOld;
    sheet.getRange(historyCommentsCol).setValue(historyComments);
   
    var historyCheck = "N" + Number(rowIndex);
    //sheet.setActiveSelection(historyCheck);
    var historyDetailsOld = sheet.getRange(historyCheck).getValue();//sheet.getActiveCell().getValue();
    var todaysDate = new Date();
    var historyDetails =  "Completed"+ "\n"+todaysDate + " - " + "\n"+ Session.getEffectiveUser() + "completed the review requested from " + row[7] + "\n"+ "Status - "+row[0] + 
                          ", Location - "+row[3]+ ", Version - "+row[4]+ ", Assigned To - "+row[5]+ 
                            ", Peer Reviewer - "+row[7]+", Lead Reviewer - "+row[8]+ "\n"+ "Comments - "+row[10]+"\n"+"\n"+"****************"+"\n"+historyDetailsOld;
    //sheet.getActiveCell().setValue(historyDetails);
    sheet.getRange(historyCheck).setValue(historyDetails);
   
   //Making the comments cell as null in the end.
   sheet.getRange("K" + Number(rowIndex)).setValue(null);
   
   //Set the Final Review date
   sheet.getRange("J" + Number(rowIndex)).setValue(Utilities.formatDate(todaysDate, 'MST', 'MM/dd/yyyy'));
   
  }
}

function onEdit(event) {
  
    var sheet = SpreadsheetApp.getActiveSheet();
    var cell =  sheet.getActiveCell();
    var rowIndex = cell.getRowIndex();
   //Browser.msgBox(Session.getEffectiveUser());
    //sheet.setActiveSelection("Z1");
    sheet.getRange("Z1").setValue(rowIndex);
  
    var cell =  sheet.getActiveCell();
    var rowIndex = cell.getRowIndex();
    sheet.getRange("O" + Number(rowIndex)).setValue(Session.getEffectiveUser());
}



/***
Test functions below
*/

function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }

/*** Getting URL of the speadsheet. */

Logger.log('URL : ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());

/*** Getting Id from URL of the speadsheet. */

Logger.log('Id from URL: ' + getIdFromUrl(SpreadsheetApp.getActiveSpreadsheet().getUrl()));


Logger.log(getIdFromUrl('SPREADSHEET URL'));

/*** Getting sheet id 
*/

var id  = SpreadsheetApp.getActiveSheet().getSheetId();// get the actual id
Logger.log('Sheet id : ' + id);// log


//Logger.log('Test');

function getSheetIdtest(){
  var id  = SpreadsheetApp.getActiveSheet().getSheetId();// get the actual id
  Logger.log(id);// log
  var sheets = SpreadsheetApp.getActive().getSheets();
  for(var n in sheets){ // iterate all sheets and compare ids
    
    var sheetids = sheets[n].getSheetId();
    var currentSheetName = SpreadsheetApp.getActive().getSheets()[n].getName();
    Logger.log('Sheet ids : ' + sheetids);
    
    if(sheets[n].getSheetId()==id){break}
  } 
  Logger.log('tab index = '+n);// your result, test below just to confirm the value
  var currentSheetName = SpreadsheetApp.getActive().getSheets()[n].getName();
  Logger.log('current Sheet Name = '+currentSheetName);
}

function t_Tilak() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell =  sheet.getActiveCell();
  var rowIndex = cell.getRowIndex();
  
  var dataRange = sheet.getRange(rowIndex,1,1,12);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
 for (i in data) {
    var row = data[i];

   Logger.log('Row(7) : ' + row[7].split(" "));
   
    var names = row[7].split(" ");
    var emailAddress = names[0].toLowerCase() +"."+names[1].toLowerCase()+"@gmail.com";
    if(emailAddress=="EMAILID")
    {
      emailAddress = "EMAILID";
    }
   
   var senderNameTemp = (Session.getEffectiveUser().toString().split("@"));
   var senderName = senderNameTemp[0].split(".");
   var subject = sheet.getName() + " Peer Review - Please review "+ row[1] + " - " + row[2];
   
   var ssURL = 'SHEET URL';
 }
}




function tFoldername(){
  var folder = DriveApp.getFolderById('1lpkA3Rs2lKosRuUZjPs042hJZbmkhfZRya_IN62Ndjk').getName();
  var files = DriveApp.getFolderById('1lpkA3Rs2lKosRuUZjPs042hJZbmkhfZRya_IN62Ndjk').getFiles();
  Logger.log(folder);
  Logger.log(files);
  Logger.log(files.length);
}


function FnctnMenuUpdateGetLocation() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var SSID=ss.get.getId();
  var file = DocsList.getFileById(SSID);

  Browser.msgBox(SSID);
  Browser.msgBox(add);
}
