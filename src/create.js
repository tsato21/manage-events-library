/**
 * Stores values input by the user necessary to create events.
 * 
 * This is a private function and used within the `create_events(pre_values,data)` function.
 * 
 * @return {Object} object The values input by the user
 * 
 */
function pre_create_(){
  //check whether the format of the target table is in the designated one.
  const format_warning = Browser.msgBox('Make sure that the first row comprises as follows: 開始日時(*date format), 開始時刻(*time format), 終了日時(*date format), 終了時刻(*time format), タイトル, 場所, 説明.',Browser.Buttons.YES_NO_CANCEL);
  if (format_warning === 'cancel'){
    Browser.msgBox(`Checking format is cancelled.`);
    return;
  } else if (format_warning === 'no'){
    Browser.msgBox('Modify the first row and do it again.');
    return;
  } else if (format_warning === 'yes') {
      let check_sheet_name = Browser.msgBox('Make sure that the sheet name is "Create".',Browser.Buttons.YES_NO_CANCEL);
      let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Create');
      if(check_sheet_name === 'cancel'){
        Browser.msgBox(`Checking sheet name is cancelled.`);
        return;
      } else if (check_sheet_name === 'no'){
        Browser.msgBox(`Modify the sheet name and try again.`);
        return;
      } else if (check_sheet_name === 'yes') {
        if (sheet === null) {
          Browser.msgBox(`There is no sheet named "Create" in this Spreadsheet. Check the sheet name and try again.`);
          return;
        }
      }
      
      const search_start = sheet.getDataRange().createTextFinder("開始日時").findNext();
      let start_cell = search_start.getA1Notation();
      const last_row = sheet.getRange(start_cell).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
      const search_end = sheet.getDataRange().createTextFinder("説明").findNext();
      const head_end_cell = search_end.getA1Notation();
      const head_end_column = head_end_cell.match(/[A-Z]+/)[0];
      let end_cell = `${head_end_column}${last_row}`;

      let check_cell = Browser.msgBox(`The range (cell of "開始日時" : last cell of "説明" column) is designated as ${start_cell} : ${end_cell}.`,Browser.Buttons.YES_NO_CANCEL);
      if (check_cell === 'no'){
        start_cell = Browser.inputBox('Set cell with "開始日時" manually.',Browser.Buttons.OK_CANCEL);
        end_cell = Browser.inputBox('Set last cell of "説明" column manually.',Browser.Buttons.OK_CANCEL);
      } else if(check_cell === 'cancel'){
        Browser.msgBox(`designating the range is cancelled.`);
        return;
      }
      
      let check_send_email = Browser.msgBox('Do you want to share target events with guests and send an email to them?',Browser.Buttons.YES_NO_CANCEL);

      if (check_send_email === 'cancel'){
        Browser.msgBox('Procedures are cancelled.');
        return;
      } else if (check_send_email === 'no'){
        return {
            'sheet_name' : 'Create',
            'start_cell' : start_cell,
            'end_cell' : end_cell
          }
      } else if (check_send_email === 'yes') {
        let subject = Browser.inputBox('Enter the subject of the email',Browser.Buttons.OK_CANCEL);
        if (subject === 'cancel'){
          Browser.msgBox('Inputting subject was cancelled.');
          return;
        } else {
          while(subject === '') {
            Browser.msgBox('You must enter the subject. Please try again.');
            subject = Browser.inputBox(`Enter the subject of the email`,Browser.Buttons.OK_CANCEL);
            if (subject === 'cancel'){
              Browser.msgBox('Inputting subject was cancelled.');
              return;
            }
          }
        }

        let guest_number = Browser.inputBox('Enter the number of the guests',Browser.Buttons.OK_CANCEL);
        if (guest_number === 'cancel') {
          Browser.msgBox('Inputting guest_number was cancelled.');
          return;
        } else {
          while(guest_number === '' || isNaN(guest_number)) {
            Browser.msgBox('You must enter the valid number. Please try again.');
            guest_number = Browser.inputBox(`Enter the number of the guests`,Browser.Buttons.OK_CANCEL);
            if (guest_number === 'cancel'){
              Browser.msgBox('Inputting guest_number was cancelled.');
              return;
            }
          }
        }
        const guest_info = [];

        for(let i=1;i<=guest_number;i++){
              let email = Browser.inputBox(`Enter the email of guest_${i}`,Browser.Buttons.OK_CANCEL);
              let email_regular_ex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; // Basic regex for email validation
              if (email === 'cancel'){
                Browser.msgBox('Inputting email was cancelled.');
                return;
              } else {
                while(email === '' || !email_regular_ex.test(email)) {
                  if (email === '') {
                    Browser.msgBox(`You must enter the email of guest_${i}. Please try again.`);
                  } else {
                    Browser.msgBox(`You entered an invalid email for guest_${i}. Please try again.`);
                  }
                  email = Browser.inputBox(`Enter the email of guest_${i}`,Browser.Buttons.OK_CANCEL);
                  if (email === 'cancel'){
                    Browser.msgBox('Inputting email was cancelled.');
                    return;
                  }
                }
              }

              let name = Browser.inputBox(`Enter the name of guest_${i}`,Browser.Buttons.OK_CANCEL);
              if (name === 'cancel'){
                Browser.msgBox('Inputting name was cancelled.');
                return;
              } else {
                while(name === '') {
                  Browser.msgBox(`You must enter the name of guest_${i}. Please try again.`);
                  name = Browser.inputBox(`Enter the name of guest_${i}`,Browser.Buttons.OK_CANCEL);
                  if (name === 'cancel'){
                  Browser.msgBox('Inputting name was cancelled.');
                  return;
                  }
                }
              }
              guest_info.push({'guest': i, 'email':email, 'name':name});
        }

        let message = 'Make sure of the input info: Email subject is ' + subject + ' / ' + 'Guest information is ';

        for(let i = 0; i < guest_info.length; i++){
            message += '[ No: ' + guest_info[i].guest + ', ';
            message += 'Email: ' + guest_info[i].email + ', ';
            message += 'Name: ' + guest_info[i].name + '] ';
        }
        const check_guest_info = Browser.msgBox(message,Browser.Buttons.YES_NO);
        if(check_guest_info === 'no'){
          Browser.msgBox('Some of the guest_info was wrong. Try again from the start.');
          return;
        } else if (check_guest_info === 'yes'){
          return {
            'sheet_name' : 'Create',
            'start_cell' : start_cell,
            'end_cell' : end_cell,
            'subject' : subject,
            'guest_info' : guest_info,
          }
        }
      }
  }
}

/**
 * Retrieves data necessary for creating events from a spreadsheet. 
 * 
 * This is a private function and used within the `create_events(pre_values,data)` function.
 * 
 * @param {Object} pre_values The result of `preparation_create()` function
 * @return {Array<Array>} two_array The values from the target range of the target spreadsheet
 * 
 */
function data_create_(pre_values){
  try{
    // Range to read (the first row/ the last column in the target table)
    const sheet_name = pre_values.sheet_name;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
    const start_cell = pre_values.start_cell;
    const end_cell = pre_values.end_cell;
    //Gets target data
    const data = sheet.getRange(`${start_cell}:${end_cell}`).getValues();    
    console.log(data);
    return data;
  } catch(e){
    Browser.msgBox(`Error get_data(pre_values): ${e.message}`);
    console.log(`Error get_data(pre_values): ${e.message}`);
  }
}

/**
 * Sends dates of created events to designated recipients by email. 
 * 
 * This is a private function and used within the `create_events(pre_values,data)` function.
 * 
 * @param {String} subject The subject of the email
 * @param {Array<Array>} event_dates The dates of the target events
 * @param {Array<String>} guest_emails The emails of guests
 * @param {Array<String>} guest_names The names of guests
 */
function send_email_guest_(subject,event_dates,guest_emails,guest_names){
  const template = HtmlService.createTemplateFromFile('email_guest');
  const my_name = PropertiesService.getScriptProperties().getProperty('my_name');

  try{
    for (let i=0;i<guest_emails.length;i++){
    template.guest_name = guest_names[i];
    template.event_dates = event_dates;
    template.my_name = my_name;
    let body = template.evaluate().getContent();
    GmailApp.sendEmail(guest_emails[i],subject,body,{htmlBody:body});
    }
  Browser.msgBox("Successfully made the events in Google Calendar and sent emails to recipients!",Browser.Buttons.OK);
  } catch(e){
    Browser.msgBox(`Error send_email_guest function: ${e.messsage}`,Browser.Buttons.OK);
  }
}

/**
 * Creates events in Google Calendar. 
 * 
 * Must set the `manage_event_library` in the target project in advance.
 * 
 * If guest information is input within the 'pre_create()` function, the target events are shared with the guests in Google Calendar, and informed to them by email.
 * 
 */
function create_events() {

  const pre_values = pre_create_();
  if (!pre_values || !pre_values.sheet_name) {
    return;
  } else {
      const data = data_create_(pre_values);

      // Column number for each item, starting with 0
      const start_date_col = 0;
      const start_time_col = 1;
      const end_date_col = 2;
      const end_time_col = 3;
      const title_col = 4;
      const location_col = 5;
      const description_col = 6;

      const calendar = CalendarApp.getDefaultCalendar();
      const event_dates = []; //This variable will be used for the function to send email

      //Defines subject for the email and email address and name for guests that will be used later in send_email_guest(event_dates,guest_emails,guest_names)
      let subject;
      let guest_emails;
      let guest_names;
      if(pre_values.hasOwnProperty('guest_info')){
          subject = pre_values.subject;
          /*
          guest_emails_string is used to show the event in the calendar of the guest's account
          */
          guest_emails_string = pre_values.guest_info.map(each_guest_info => each_guest_info.email).join(",");
          /*
          guest_emails and guest_names are used to send an email to guests
          */
          guest_emails = pre_values.guest_info.map(each_guest_info => each_guest_info.email);
          guest_names = pre_values.guest_info.map(each_guest_info => each_guest_info.name);
      }


      try {
        //Since the first array is the header of the data, "i" starts with 1. 
        for (i=1; i<data.length; i++) {
          let start_date = new Date(data[i][start_date_col]);
          // let start_date = new Date (Utilities.formatDate(data[i][start_date_col],"GMT","yyyy/MM/dd"));
          
          //Pushes start_date to the variable, "event_date"
          event_dates.push((start_date.getMonth() + 1) + "/" + start_date.getDate()); //store start date for send an email to guests
          
          let end_date = new Date(data[i][end_date_col]);
          // let end_date = new Date (Utilities.formatDate(data[i][end_date_col],"GMT","yyyy/MM/dd"));
          let start_time = data[i][start_time_col];
          let end_time = data[i][end_time_col];
          if(start_time !=='' && end_time !== ''){
            console.log('start_time and end_time do exist.');
            // start_time = new Date (Utilities.formatDate(data[i][start_time_col],"GMT","HH:mm"));
            // end_time = new Date (Utilities.formatDate(data[i][end_time_col],"GMT","HH:mm"));
            start_time = new Date(data[i][start_time_col]);
            end_time = new Date(data[i][end_time_col]);
          }

          console.log(start_date,start_time);

          
          let title = data[i][title_col];
          let location = data[i][location_col];
          let description = data[i][description_col];      
          
          // Sets options for creating calendar
          let options;
          if(pre_values.hasOwnProperty('guest_info')){
            console.log('guest_info does exist.');
            options = { 
                  location: location,
                  description: description,
                  guests: guest_emails_string
                };
          } else {
            console.log('guest_info does not exist.');
            options = { 
                  location: location,
                  description: description,
                };
          }            

          // Checks if the start time and end time are empty
          if (start_time === '' || end_time === '') {
            // If start_date is equal to end_date, an event for one day is created.
            
            if (JSON.stringify(start_date) === JSON.stringify(end_date)){
              console.log('One day event with no specified start and end time');
              calendar.createAllDayEvent(
                title,
                start_date,
                options
              );
            
              // Otherwise, an event for multiple days is created
            } else {
              console.log('Multiple-days event with no specified start and end time');

              //Since end_date is exclusive for createAllDayEvent, although the event ends on the day of the end_date, it does not include this day, and 1 more day should be added.
              end_date.setDate(end_date.getDate() + 1);
              calendar.createAllDayEvent(
                title,
                start_date,
                end_date,
                options
              );
            }
            
          // Sets events that have start_time and end_time
          } else {
            console.log('An event with specified start and end time');
            
            start_date.setHours(start_time.getHours(),start_time.getMinutes());
            end_date.setHours(end_time.getHours(),end_time.getMinutes());
            
            calendar.createEvent(
              title,
              start_date,
              end_date,
              options
            );
          }
          
        }
        // Outputs log if there is an error
      } catch (e) {
        console.log(`Error create_events(pre_values,data) ${e.message}`);
        Browser.msgBox(`Error create_events(pre_values,data) ${e.message}`,Browser.Buttons.OK);
        return;
      }
      try{
        //Checks if the property, "guest_emails" does exist. if so, the function to send an email to guests is executed.
        if (pre_values.hasOwnProperty('guest_info')) {
          send_email_guest_(subject,event_dates,guest_emails,guest_names);
        } else {
          Browser.msgBox("Successfully made the events in Google Calendar!",Browser.Buttons.OK);
        }
      } catch (e){
        console.log(`Error send_email_guest(event_dates) ${e.message}`);
      }
  }
}