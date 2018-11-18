function test()
{   
      var emailAddress = "EMAILID";
      var subject = "Test Message to see the format"
      var message =  "<html><head></head><body>"+"<br />"+
                  "<div style='font-size: 24px;font-weight: strong;padding: 45px 25px 15px 25px;'> <center> Peer Review</center> </div>"+
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
                  "</div></body></html>";
  
   MailApp.sendEmail(emailAddress, subject, message, {htmlBody: message});
}