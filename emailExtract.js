//Get the newest 50 emails from the specific label that we want
//Go through all of the emails that were received today
//Log any valid emails and mark as read
//If we find an email that is not a valid email then we mark it as other/update and leave unread
//Gets logged in a google sheet

function dataExtract() {
  try {
    //Get the date for today
    let curDate = new Date();
    curDate = getDate(curDate.toString());

    //Go through the rows of the table and find the rows where the date is equal to today
    //This is used to find if it has been logged already
    let dateRange = getDateRange(curDate, "A:A");

    const label = GmailApp.getUserLabelByName('LABEL NAME HERE');
    const newLabel = GmailApp.getUserLabelByName('NEW LABEL NAME HERE');
    let junk = false;

    Logger.log("Checking 50 newest emails");

    //Get the 50 newest email threads (arbitrary number can be changed)
    threads = label.getThreads(0, 50);

    //Loop through the list of email threads one at a time
    for (let i = 0; i < threads.length; i++){
      //We only check the first initial message because all of the information is in the subject line, we don't need the body at all
      messages = threads[i].getMessages();
      message = messages[0];

      //Check to see if the date of the email, if it is not dated today then we end early!
      date = getDate(message.getDate().toString())
      if (date != curDate){
        Logger.log("No more new emails");
        break;
      }
      //It it is an email from today we loop through all of the labels and see if it is already marked NEW LABEL
      //If it is labeled as other/update then we skip the email
      else {
        labels = threads[i].getLabels();
        for (let j = 0; j < labels.length; j++){
          if (labels[j].getName() == "NEW LABEL NAME HERE"){
            Logger.log(message.getSubject().toString() + " is already marked NEW LABEL");
            junk = true;
            break;
          }
        }
        //If not marked junk then we get the subject line of the email
        //From the subject we are then able to grab whatever information we want
        if (junk != true){
          subject = message.getSubject().toString();
          //Get other information from subject
	  information = getInformation(subject);

          //If we can't find the information that we need then we mark as NEW LABEL
          if (information = "Could not find it"){
            Logger.log(subject + " is NEW LABEL");
            threads[i].addLabel(newLabel);
            message.markUnread();
          }
          else{
            //If we have the information then we log it
            //We use Compare() to see if it has been logged already
            if (information != "Could not find it"){
              updated = informationCompare(information, dateRange[0], 2, dateRange[1], date);
              if (updated == 1){
                message.markRead();
                //Increase the date range since we added another OC
                dateRange[1] = dateRange[1] + 1;
              }
            }
          }
        }
      }
      junk = false;
    }
  } catch(err){
    console.log('Failed with error %s', err.message);
  }
}

function getDate(dateTime){
  //Slice up the date and time string into month (convert month to number), day and year. Then concatenate them all together.
  try {
    let month = dateTime.slice(4,7);
    switch(month){
      case "Jan":
        month = 1;
        break;
      case "Feb":
        month = 2;
        break;
      case "Mar":
        month = 3;
        break;
      case "Apr":
        month = 4;
        break;
      case "May":
        month = 5;
        break;
      case "Jun":
        month = 6;
        break;
      case "Jul":
        month = 7;
        break;
      case "Aug":
        month = 8;
        break;
      case "Sep":
        month = 9;
        break;
      case "Oct":
        month = 10;
        break;
      case "Nov":
        month = 11;
        break;
      case "Dec":
        month = 12;
        break;
    }
    let day = dateTime.slice(8,10);
    if (day.slice(0,1) == 0){
      day = day.slice(1);
    }
    const year = dateTime.slice(11, 15);
    const date = month + "/" + day + "/" + year;
    return date;

  } catch(err) {
    console.log('Failed with error %s', err.message);
  }
}

//Locate the information from the subject line and subsequently strip it
function getInformation(subject_line) {
  try {
    //We use this regex to search for numbers that are 10 digits long
    const pattern = :WHATEVER PATTERN YOU NEED"
    position = subject_line.search(pattern);

    if (position != -1){
      orderNumber = subject_line.slice(position, position + 10);
    }
    else {
      orderNumber = "Could not find INFORMATION";
      Logger.log("Could not find order INFORMATION");
    }
    return orderNumber;

  } catch(err){
    console.log('Failed with error %s', err.message);
  }
}

function updateSpreadsheet(input1, input2){
  //open spreadsheet and insert data at the top, can be edited to add as many inputs as you want
  try {
    const ss = SpreadsheetApp.openByUrl('SPREADSHEET URL HERE',);
    const sheet = ss.getSheets()[0];
    sheet.insertRows(2);
    data = [[input1, input2, input3]];
    sheet.getRange(2, 1, 1, 2).setValues(data);
    Logger.log("Spreadsheet updated with INFORMATION " + input2 + "!");
    return 1;
  } catch(err) {
    console.log('Failed with error %s', err.message);
  }
}

function getDateRange(date, range){
  //Find the range of data that matches our date, outputs the starting row and how many rows match
  try{
    let start = 2;
    let numRows = 0;

    const ss = SpreadsheetApp.openByUrl('SPREADSHEET URL HERE',);
    const sheet = ss.getSheets()[0];
    columnRange = sheet.getRange(range);
    columnValues = columnRange.getValues();

    for(let i = 1; i < columnValues.length; i++){
      if (columnValues[i] == date){
        numRows = numRows + 1;
      }
      else {
        break;
      }
    }
    return [start, numRows];

  } catch (err){
    console.log('Failed with error %s', err.message);
  }
}

function columnCompare(value, row, column, numRows){
  //Search the data in the selected column to see if any match the value
  try{
    let output = 0;
    const ss = SpreadsheetApp.openByUrl('SPREADSHEET URL HERE',);
    const sheet = ss.getSheets()[0];
    columnRange = sheet.getRange(row, column, numRows);
    columnValues = columnRange.getValues();

    for (let i = 0; i < columnValues.length; i++){
      if (columnValues[i] == value){
        output = 1;
        break;
      }
    }
    return output;

  } catch (err){
    console.log('Failed with error %s', err.message);
  }
}

function informationCompare(information, row, column, numRows, date){
  //Compare the information in the range and return if we get a match or not, if not then add the information
  try{
    let result = 0;

    if (numRows != 0){
      compare = columnCompare(information, row, column, numRows);
    }
    else{
      compare = 0;
    }
    if (compare == 0){
      updateSpreadsheet(information, po);
      result = 1;
    }
    else {
      Logger.log("INFORMATION " + information + " already logged in sheets.");
    }
    return result;

  } catch(err){
    console.log('Failed with error %s', err.message);
  }
}