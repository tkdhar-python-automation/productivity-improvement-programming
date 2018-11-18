/**
 * Serves HTML of the application for HTTP GET requests.
 * If folderId is provided as a URL parameter, the web app will list
 * the contents of that folder (if permissions allow). Otherwise
 * the web app will list the contents of the root folder.
 *
 * @param {Object} e event parameter that can contain information
 *     about any URL parameters provided.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index');
}


/***

Improvement areas of this code as of Jul-02-2016:
=================================================
1. Code clean up.
2. Week number should be presented with Date.
3. Implementation of defaulter case 2 and case 3.
4. Automated email to defaulter with detail.
5. Proper naming convension and traceability.
6. Defaulter email id generation code is not working fine for back dated sheet.
7. Currently the code can automatically validate current week's data. It needs to be extended for past sheet as well.
8. Code to convert XML to XLS
9. Code to upload file automatically taking filename as input.
10. Instead of hard-coding the User list in google sheet, user list can be read from google group - Nov-18-2018


Note: For testing following data manipulated in the Jul-02-2016 sheet

Tilak's WTS entries are made "Open".
Nitasha's one row is removed.
**/

/*** Status of improvements:

Jul-03-3016:

Point 3: Completed.


**/

// Creating a new Spreadsheet if the same name Spreadsheet does not exist.
function createSpreadSheet(spreadSheetName) {
  var files = DriveApp.getFilesByName(spreadSheetName);
  if(!files.hasNext()) {
    var ssNew = SpreadsheetApp.create(spreadSheetName,1,1);
  }
  else{
    // If file exist then delete the file and recreate it
    deleteSpreadsheet(spreadSheetName);
    var ssNew = SpreadsheetApp.create(spreadSheetName,1,1);   
  }
  return ssNew;
}

// Deleting a complete Spreadsheet using name
// This is needed to calculate current week
function deleteSpreadsheet(spreadSheetName){
  var files = DriveApp.getFilesByName(spreadSheetName);
  while (files.hasNext()) {
    var eachFile = files.next();
    
    var idToDLET = eachFile.getId();
    //Logger.log('idToDLET: ' + idToDLET);
    
    var rtrnFromDLET = DriveApp.getFileById(idToDLET).setTrashed(true);
  }
}


// Deleting a sheet from a Spreadsheet using name




// Getting current week number in the form of Week<Space><number>
function getCurrentWeekNum(){
  
  // Creating the Spreadsheet so that we can use spreadsheet formula to calculate current week.
  // Create a google sheet for Current Week Calculation if it is not exist
  var ssNew = createSpreadSheet('Temp');
  
  var sheet = ssNew.getSheets()[0];
  
  var cell = sheet.getRange("A1");
  cell.setFormula("=WEEKNUM(today(),2)");
  var WeekNum = sheet.getRange("A1").getValue();
  
  
  // Deleting the google sheet created Current Week Calculation
  deleteSpreadsheet('Temp');
  
  return ("Week " + WeekNum);
}

// Getting Distinct Project Codes
function getUniqueProjCode(arrayWTS){
  var uniqueProjCode = new Array();
  for(nn in arrayWTS){
    var duplicate = false;
    for(j in uniqueProjCode){
      if(arrayWTS[nn][4] == uniqueProjCode[j][0]){
        duplicate = true;
      }
    }
    if(!duplicate){
      uniqueProjCode.push([arrayWTS[nn][4]]);
    }
  }
  return uniqueProjCode;
}


// Getting Project Wise Report: 
// Input is an array of data
// Output is Name of the people saved or submitted WTS
//Week, People Name, Project Name, Total No. of Hours, 
function getProjectCodeWiseReport(arrayWTS,caseNo){
  
  var totalPerPerson = 0;
  var weeklyPeopleProjStatusTotalNoHours = new Array();
  var peopleFilledWTS = new Array();
  
  // Getting the unique employee name
  var uniqueNameData = getUniqueEmpName(arrayWTS);
  
  //Logger.log('Name of unique emp : ' + uniqueNameData);
  
  for (var j = 0; j < uniqueNameData.length; j++) {
    var name = uniqueNameData[j];
    for (var i = 0; i < arrayWTS.length; i++) {
      if (arrayWTS[i][1] == name){
        totalPerPerson = totalPerPerson + arrayWTS[i][5];
        var empName = arrayWTS[i][1];
        var weekNo = arrayWTS[i][0];
        var ProjectName = arrayWTS[i][4];
        var ApprovalStatus = arrayWTS[i][11];
      }
    }
    weeklyPeopleProjStatusTotalNoHours.push([weekNo,empName,ProjectName,totalPerPerson,ApprovalStatus]);
    peopleFilledWTS.push([empName]);
    totalPerPerson = 0;
  }
  if (caseNo == 1){
  return peopleFilledWTS;
  }
  else{
    return weeklyPeopleProjStatusTotalNoHours;
  }
}

// Copy subset of data from one array to other array
//var weekNum = getCurrentWeekNum();

// Reading subset of data from the newly created spreadsheet
// parameter 1: Name of the sheet
// parameter 2: Current week number calculated above
// parameter 3: CM Team billable members list
function readDataFromSheet(filename, weekNum, cmTeamList) {
  var files = DriveApp.searchFiles('title =' +  '"' + filename + '"');
  while (files.hasNext()) {
    var spreadsheet = SpreadsheetApp.open(files.next());
    var sheet = spreadsheet.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    
    var subsetArray = new Array();
    
    //    var cmTeamList = ['Awasthi, Neha',
    //                      'Bhalerao, Rhishikesh Ramchandra',
    //                      'Chopde, Sanjay',
    //                      'Deshpande, Supriya',
    //                      'Dhar, Tilak',
    //                      'Jadhav, Sushil',
    //                      'Kumar, Saurabh',
    //                      'Lal, Neha',
    //                      'Loknath, Shaw',
    //                      'Patil, Dhanashree',
    //                      'Pawar, Neeraj',
    //                      'Pious, Aldrin',
    //                      'Purohit, Shrinidhi',
    //                      'Rasal, Madhav',
    //                      'Reddy, Maheshwar',
    //                      'Satoor, Mayura',
    //                      'Sharma, Nitasha',
    //                      'Sonawane, Poonam'];
    
    // Test logger
    //    for (var i = 10; i < data.length; i++) {
    //      // Displaying data for current week
    //      if (data[i][0] == weekNum){
    //        Logger.log('Week: ' + data[i][0]);
    //        Logger.log('Name of Employee: ' + data[i][1]);
    //        Logger.log('Project / Task (Full Path): ' + data[i][4]);      
    //        Logger.log('Total Hrs: ' + data[i][5]);
    //        Logger.log('Approval Status: ' + data[i][11]);
    //      }
    //    }
    
    // Defaulter array
    var defaulterArray = new Array();
    
    // Storing subset of data in new array subsetArray
    for (var i = 10; i < data.length; i++) {
      for (var j in cmTeamList) {
        //Logger.log('cmTeamList[j] :' + cmTeamList[j]);
        // Displaying data for current week
        if ((data[i][0] == weekNum) && (data[i][1] == cmTeamList[j])){
          subsetArray.push(data[i]);
        }
      }
    }
  } 
  
  return subsetArray;
}

// Function to get only Customer Master Team's data for current week
function copySubset1Array2OtherArray(fileName, weekNum, cmTeamList){
  var arrayWTS = readDataFromSheet(fileName, weekNum, cmTeamList);
  return arrayWTS;
}


// Getting Distinct Employee Names
function getUniqueEmpName(arrayWTS){
  var uniqueNameData = new Array();
  for(nn in arrayWTS){
    var duplicate = false;
    for(j in uniqueNameData){
      if(arrayWTS[nn][1] == uniqueNameData[j][0]){
        duplicate = true;
      }
    }
    if(!duplicate){
      uniqueNameData.push([arrayWTS[nn][1]]);
    }
  }
  return uniqueNameData;
}

// Function call for getting only Customer Master Team's data for current week
//var arrayWTS = copySubset1Array2OtherArray("AAA Saama Day Level Report Daily-Jul-02-2016.xlsx", weekNum);

// Getting Total Number of Hours per person per week
function getTotalNoHoursPerPersonPerWeek(arrayWTS){
  var totalPerPerson = 0;
  var totalNoHoursPerPersonPerWeek = new Array();
  
  // Getting the unique employee name
  var uniqueNameData = getUniqueEmpName(arrayWTS);
  
  //Logger.log('Name of unique emp : ' + uniqueNameData);
  
  for (var j = 0; j < uniqueNameData.length; j++) {
    var name = uniqueNameData[j];
    for (var i = 0; i < arrayWTS.length; i++) {
      if (arrayWTS[i][1] == name){
        totalPerPerson = totalPerPerson + arrayWTS[i][5];
      }
    }
    totalNoHoursPerPersonPerWeek.push(uniqueNameData[j],totalPerPerson);
    totalPerPerson = 0;
  }
  return totalNoHoursPerPersonPerWeek;
}

// Getting Distinct Approval Status
// Assuming that one employee is filling WTS in one project code and the status of his
// all 5 days timesheet
function getUniqueApprovalStatusPerPerson(arrayWTS){
  var newdata = new Array();
  
  // Getting the unique employee name
  var uniqueNameData = getUniqueEmpName(arrayWTS);
  
  for (var k = 0; k < uniqueNameData.length; k++) {
    var name = uniqueNameData[k];
    for(nn in arrayWTS){
      if (arrayWTS[nn][1] == name){
        newdata.push(name,arrayWTS[nn][11]);
        break;
      }
    }
  }
  return newdata;
}

// Find the difference between two arrays
function arr_diff (a1, a2) {
  
  //  var a1 = [1,2,3];
  //  var a2 = [1,2];
  
  var a = [], diff = [];
  
  for (var i in a1) {
    a[a1[i]] = true;
  }
  
  for (var i in a2) {
    if (a[a2[i]]) {
      delete a[a2[i]];
    } else {
      a[a2[i]] = true;
    }
  }
  
  for (var k in a) {
    diff.push(k);
  }
  
  return diff;
};

// Finding the WTS submission report for the current week.
function wtsSubmissionReport(getTotalNoHoursPerPersonPerWeek, weekNum){
  
  // Populating a reporting Spreadsheet with defaulter detail
  var ssNew = createSpreadSheet('WTS Submission Report ' + weekNum);
  
  var sheet = ssNew.getSheets()[0];
  
  // Create Header Row
  
  sheet.appendRow(["Week No","Emp Name","Project Name","Total Per Person","Approval Status"]); 
  
  for(var i in getTotalNoHoursPerPersonPerWeek){
    //Logger.log('getTotalNoHoursPerPersonPerWeek : ' + getTotalNoHoursPerPersonPerWeek[i][1]);
    //Logger.log('getTotalNoHoursPerPersonPerWeek[i][0], getTotalNoHoursPerPersonPerWeek[i][1], getTotalNoHoursPerPersonPerWeek[i][2], getTotalNoHoursPerPersonPerWeek[i][3]], getTotalNoHoursPerPersonPerWeek[i][4] : ' + getTotalNoHoursPerPersonPerWeek[i][0] + getTotalNoHoursPerPersonPerWeek[i][1] + getTotalNoHoursPerPersonPerWeek[i][2] + getTotalNoHoursPerPersonPerWeek[i][3]] + getTotalNoHoursPerPersonPerWeek[i][4]);
    sheet.appendRow([getTotalNoHoursPerPersonPerWeek[i][0], getTotalNoHoursPerPersonPerWeek[i][1], getTotalNoHoursPerPersonPerWeek[i][2], getTotalNoHoursPerPersonPerWeek[i][3], getTotalNoHoursPerPersonPerWeek[i][4]]);
  }
  
}

// Finding the defaulters: Case1: WTS neither saved or submitted, Case2: WTS saved but not submitted.
function wtsDefaulters(peopleFilledWTS, cmTeamList, weekNum){
  
  // Populating a reporting Spreadsheet with defaulter detail
  var ssNew = createSpreadSheet('WTS Defaulter List ' + weekNum);
  
  var sheet = ssNew.getSheets()[0];
  
  // Appends a new row with 3 columns to the bottom of the
  // spreadsheet containing the values in the array
  // Create Header Row
  
  sheet.appendRow(["Name", "This Week Total Hour", "Approval Status"]); 
  
  var peopleName = new Array();
  
  for(var i=0; i<peopleFilledWTS.length; i++){
    //  Logger.log('wklyPplProjStsTlNoHours : ' + wklyPplProjStsTlNoHours);
    //  Logger.log('peopleFilledWTS : ' + peopleFilledWTS);
    //  Logger.log('peopleFilledWTS[i][1] : ' + peopleFilledWTS[i][1]);
    peopleName.push(peopleFilledWTS[i][1]);
    
    // Defaulter Case2: WTS saved but not submitted
    if(peopleFilledWTS[i][4] == "Open"){
      sheet.appendRow([peopleFilledWTS[i][1], peopleFilledWTS[i][3], peopleFilledWTS[i][4]]);
    }
    Logger.log('peopleFilledWTS[i][3] :' + peopleFilledWTS[i][3]);
    Logger.log('peopleFilledWTS[i][4] :' + peopleFilledWTS[i][4]);
    
    // Defaulter Case3: WTS saved or submitted less than 40 hours
    if(peopleFilledWTS[i][3] < 40){
      sheet.appendRow([peopleFilledWTS[i][1], peopleFilledWTS[i][3], peopleFilledWTS[i][4]]);
    }
  }
  
  // Case1: Finding name of the people who did not fill WTS
  var wtsDefaulters = arr_diff(peopleName,cmTeamList);
  //Logger.log('wtsDefaulters : ' + wtsDefaulters);  
    
  for( var i in wtsDefaulters){
    sheet.appendRow([wtsDefaulters[i], "", ""]);
  }
  
  return wtsDefaulters;
}

function splitArray(str) {
  var splittedArray = [{}];
  
  splittedArray = str.split(", ");
  return splittedArray;
}

// Finding the csaa email id of the defaulters.
function getEmailOfDefaulters(listOfWTSDefaulters){
  var str = listOfWTSDefaulters.toString();
  var splittedArray = splitArray(str);
  var emailID = splittedArray[1] + '.' + splittedArray[0] + '@csaa.com';
  //Logger.log("Defaulters' email id : " + emailID);
  return emailID;
}


/* finds the intersection of 
 * two arrays in a simple fashion.  
 *
 * PARAMS
 *  a - first array, must already be sorted
 *  b - second array, must already be sorted
 *
 * NOTES
 *
 *  Should have O(n) operations, where n is 
 *    n = MIN(a.length(), b.length())
 */
function intersect_2arrays(a, b)
{
  var ai=0, bi=0;
  var result = [];

  while( ai < a.length && bi < b.length )
  {
     if      (a[ai] < b[bi] ){ ai++; }
     else if (a[ai] > b[bi] ){ bi++; }
     else /* they're equal */
     {
       result.push(a[ai]);
       ai++;
       bi++;
     }
  }

  Logger.log('Intersection of a & b is using non-destructive: ' + result);
  
  return result;
}


// Function to test other functions
function test(){

// Getting user input for WTS report file name and week number

//  var wtsReportFilename = Browser.inputBox("Please enter WTS report file name : ");  
//  var weekNum = Browser.inputBox("Please enter the week number you want to run the validation : ");  
//  
  
//  // Getting the weekNum input from user for running the script for week number other than current week number.
//  var weekNum = Browser.inputBox( 'Please enter week number if you want to run this code for week number other than current week number');
//
//  // If user input is empty, then find current week number.
//  if (weekNum == ""){
//    weekNum = getCurrentWeekNum();
//  }
  
  // If user input is blank for week number, current week number will be considered.
  if (weekNum == ""){
    weekNum = getCurrentWeekNum();
  }
  
  //var weekNum = "Week 28";
  //Logger.log('Week Number : ' + getCurrentWeekNum());
  
  var cmTeamList = ['Awasthi, Neha',
                    'Bhalerao, Rhishikesh Ramchandra',
                    'Chopde, Sanjay',
                    'Deshpande, Supriya',
                    'Dhar, Tilak',
                    'Jadhav, Sushil',
                    'Kumar, Saurabh',
                    'Lal, Neha',
                    'Loknath, Shaw',
                    'Patil, Dhanashree',
                    'Pawar, Neeraj',
                    'Pious, Aldrin',
                    'Purohit, Shrinidhi',
                    'Rasal, Madhav',
                    'Reddy, Maheshwar',
                    'Satoor, Mayura',
                    'Sharma, Nitasha',
                    'Sonawane, Poonam'];
  
  //var arrayWTS = copySubset1Array2OtherArray("AAA Saama Day Level Report Daily-Jul-23-2016.xlsx", weekNum, cmTeamList);
  if(wtsReportFilename != ""){
    var arrayWTS = copySubset1Array2OtherArray(wtsReportFilename, weekNum, cmTeamList);
  }
  else{
    Browser.msgBox("WTS report file name is null. Please enter file name correctly and try again");
    return;
  }
  
  
  //Logger.log('Current Week Data : ' + arrayWTS);
  //Logger.log('Total number of rows : ' + arrayWTS.length);
  
  //  var totalPerPerson = getTotalNoHoursPerPersonPerWeek(arrayWTS);
  //  //Logger.log('Current Week Data : ' + arrayWTS);
  //  //Logger.log('totalPerPerson : ' + totalPerPerson);
  //  
  //  var approvalStatus = getUniqueApprovalStatusPerPerson(arrayWTS);
  //  //Logger.log('Distinct Approval Status : ' + approvalStatus);
  //  var UniqueProjCode = getUniqueProjCode(arrayWTS);
  //  //Logger.log('UniqueProjCode : ' + UniqueProjCode);
  //  
  
  //var weeklyPeopleProjStatusTotalNoHours = [];
  var peopleFilledWTS = getProjectCodeWiseReport(arrayWTS,1);
  //Logger.log('Weekly People submitted WTS : ' + peopleFilledWTS);
  
  var getTotalNoHoursPerPersonPerWeek = getProjectCodeWiseReport(arrayWTS,2);
  //Logger.log('Weekly People submitted WTS : ' + getTotalNoHoursPerPersonPerWeek);
  
  // Finding the defaulters: Case1: WTS neither saved or submitted, Case2: WTS saved but not submitted Case3: WTS submitted less than 40 hours.
//  var listOfWTSDefaulters = wtsDefaulters(peopleFilledWTS, cmTeamList);
//  Logger.log('Weekly People NOT saved or submitted WTS : ' + listOfWTSDefaulters);
  
  var listOfWTSDefaulters = wtsDefaulters(getTotalNoHoursPerPersonPerWeek, cmTeamList, weekNum);
  Logger.log('Weekly People NOT saved or submitted WTS : ' + listOfWTSDefaulters);
  
  // Finding the email Id of defaulters
  
  var emailID = getEmailOfDefaulters(listOfWTSDefaulters);
  Logger.log('Email ID of defaulters : ' + emailID);
  
  
 // Weekly WTS report 
  
 wtsSubmissionReport(getTotalNoHoursPerPersonPerWeek, weekNum);
  
}

/***

Below codes are used to unit test new functionality.

***/

// Test function to test intersection

/* destructively finds the intersection of 
 * two arrays in a simple fashion.  
 *
 * PARAMS
 *  a - first array, must already be sorted
 *  b - second array, must already be sorted
 *
 * NOTES
 *  State of input arrays is undefined when
 *  the function returns.  They should be 
 *  (prolly) be dumped.
 *
 *  Should have O(n) operations, where n is 
 *    n = MIN(a.length, b.length)
 */

function intersection_destructive(a, b)
{
  var result = [];
  while( a.length > 0 && b.length > 0 )
  {  
     if      (a[0] < b[0] ){ a.shift(); }
     else if (a[0] > b[0] ){ b.shift(); }
     else /* they're equal */
     {
       result.push(a.shift());
       b.shift();
     }
  }

  Logger.log('Intersection of a & b is using destructive: ' + result);
  return result;
}


/* finds the intersection of 
 * two arrays in a simple fashion.  
 *
 * PARAMS
 *  a - first array, must already be sorted
 *  b - second array, must already be sorted
 *
 * NOTES
 *
 *  Should have O(n) operations, where n is 
 *    n = MIN(a.length(), b.length())
 */
function intersect_safe(a, b)
{
  var ai=0, bi=0;
  var result = [];

  while( ai < a.length && bi < b.length )
  {
     if      (a[ai] < b[bi] ){ ai++; }
     else if (a[ai] > b[bi] ){ bi++; }
     else /* they're equal */
     {
       result.push(a[ai]);
       ai++;
       bi++;
     }
  }

  Logger.log('Intersection of a & b is using non-destructive: ' + result);
  
  return result;
}



function callIntersection(){
  intersection_destructive([1,2,3], [2,3,4,5]);
  intersect_safe([1,2,3], [2,3,4,5]);
}