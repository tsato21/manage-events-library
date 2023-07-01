/**
 * Stores values input by the user necessary to update & delete events.
 * 
 * This is a private function and used in the `display_events()`.
 * 
 * @return {Object} object The values input by the user
 * 
 */
function pre_update_delete_(){
  const format_warning = Browser.msgBox('Make sure that the first row comprises as follows: 開始日時	開始時刻	終了日時	終了時刻	タイトル	場所	説明	ID	Update/Delete.',Browser.Buttons.YES_NO_CANCEL);
  if (format_warning === 'cancel'){
    Browser.msgBox('Checking format is cancelled.');
    return;
  } else if (format_warning === 'no'){
    Browser.msgBox('Modify the first row and do it again.');
    return;
  } else if (format_warning === 'yes') {
      let check_sheet_name = Browser.msgBox('Make sure that the sheet name is "Update-Delete".',Browser.Buttons.YES_NO_CANCEL);
      let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Update-Delete');
      if(check_sheet_name === 'cancel'){
        Browser.msgBox(`Inputting sheet name is cancelled.`);
        return;
      } else if (check_sheet_name === 'yes') {
        if (sheet === null) {
          Browser.msgBox(`There is no sheet named "Update-Delete" in this Spreadsheet. Check the sheet name and try again.`);
          return;
        }
      }

      const find_start = sheet.getDataRange().createTextFinder("開始日時").findNext();
      const find_id = sheet.getDataRange().createTextFinder("ID").findNext();
      const find_validate = sheet.getDataRange().createTextFinder("Update/Delete").findNext();

      const start_head_cell = find_start.getA1Notation();
      const head_column = start_head_cell.match(/[A-Z]+/)[0];
      let head_row = start_head_cell.match(/\d+/)[0];
      let start_cell = `${head_column}${parseInt(head_row) + 1}`;
      const id_head_cell = find_id.getA1Notation();
      let id_column = id_head_cell.match(/[A-Z]+/)[0];
      const validate_head_cell = find_validate.getA1Notation();
      let validate_column = validate_head_cell.match(/[A-Z]+/)[0];

      let check_cell_column = Browser.msgBox(`The start cell to input events is ${start_cell}. ; The column for ID is ${id_column}. ;  The column for "Update/Delete" is ${validate_column}.`,Browser.Buttons.YES_NO_CANCEL);
      while (check_cell_column !== 'yes'){
          if(check_cell_column === 'cancel'){
            Browser.msgBox(`Designating the start cell is cancelled.`);
            return;
          } else if (check_cell_column === 'no'){
            start_cell = Browser.inputBox('Set the start cell to input events manually.',Browser.Buttons.OK_CANCEL);
            id_column = Browser.inputBox('Set the column for id manually.',Browser.Buttons.OK_CANCEL);
            validate_column = Browser.inputBox('Set the column for validate manually.',Browser.Buttons.OK_CANCEL);
            if(start_cell === 'cancel' || id_column === 'cancel' || validate_column === 'cancel'){
              Browser.msgBox(`Inputting cell and column info manulally is cancelled.`);
              return;
            }
            check_cell_column = Browser.msgBox(`The start cell to input events is ${start_cell}. ; The column for ID is ${id_column}. ;  The column for "Update/Delete" is ${validate_column}.`,Browser.Buttons.YES_NO_CANCEL);
          }
      }

      head_row = start_head_cell.match(/\d+/)[0];

      let search_type = Browser.inputBox('Choose search type from either "Keyword" or "Period"',Browser.Buttons.OK_CANCEL);
      if (search_type === 'cancel'){
        Browser.msgBox('Inputting search type was cancelled.');
        return;
      } else {
        while (search_type !== 'Keyword' && search_type !== 'Period'){
          Browser.msgBox('Input search type was invalid. Try again.');
          search_type = Browser.inputBox('Choose search type from either "Keyword" or "Period"',Browser.Buttons.OK_CANCEL);
          if (search_type === 'cancel'){
            Browser.msgBox('Inputting search type was cancelled.');
            return;
          }
        }
      }

      if (search_type === 'Keyword'){
        let keyword = Browser.inputBox('Input a keyword',Browser.Buttons.OK_CANCEL);
        let period = Browser.inputBox('Choose period from either "365", "30", or "7',Browser.Buttons.OK_CANCEL);
        while (period!== '365' && period!== '30' && period!== '7'){
          Browser.msgBox('Input period was invalid. Try again.');
          period = Browser.inputBox('Choose period from either "365", "30", or "7',Browser.Buttons.OK_CANCEL);
          if (period === 'cancel'){
            Browser.msgBox('Inputting period was cancelled.');
            return;
          }
        }
        return {
          'head_row':head_row,
          'start_cell':start_cell,
          'id_column':id_column,
          'validate_column':validate_column,
          'keyword':keyword,
          'period':period
        }
      } else {
        let start_date = Browser.inputBox('Input a start date (YYYY/MM/DD)',Browser.Buttons.OK_CANCEL);
        let end_date = Browser.inputBox('Input a end date (YYYY/MM/DD)',Browser.Buttons.OK_CANCEL);

        let date_regular_ex = /^(19|20)\d\d\/(0[1-9]|1[0-2])\/(0[1-9]|[12][0-9]|3[01])$/;
        while(start_date === '' || !date_regular_ex.test(start_date)) {
          if (start_date === '') {
            Browser.msgBox(`You must enter a start date. Please try again.`);
          } else {
            Browser.msgBox(`You entered an invalid start date. Please try again.`);
          }
          start_date = Browser.inputBox('Input a start date (YYYY/MM/DD)',Browser.Buttons.OK_CANCEL);
          if (start_date === 'cancel'){
            Browser.msgBox('Inputting start date was cancelled.');
            return;
          }
        }

        while(end_date === '' || !date_regular_ex.test(end_date)) {
          if (end_date === '') {
            Browser.msgBox(`You must enter an end date. Please try again.`);
          } else {
            Browser.msgBox(`You entered an invalid end date. Please try again.`);
          }
          end_date = Browser.inputBox('Input an end date (YYYY/MM/DD)',Browser.Buttons.OK_CANCEL);
          if (end_date === 'cancel'){
            Browser.msgBox('Inputting end date was cancelled.');
            return;
          }
        }
        return {
          'head_row':head_row,
          'start_cell':start_cell,
          'id_column':id_column,
          'validate_column':validate_column,
          'start':start_date,
          'end':end_date
        }
      }
  }
}

/**
 * Displays events during the period designated by the user in Spreadsheet.
 * 
 * This is a private function and used in the `display_events()`.
 * 
 * @param {Object} pre_values The values that the user input.
 */
function display_events_period_(pre_values) {
      
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Update-Delete');
  const start = pre_values.start;
  const end = pre_values.end;
  const start_date = new Date(start);
  const end_date = new Date(end);
  start_date.setHours(0,0);
  end_date.setHours(23,59);
  console.log(start_date,end_date);

  //get events during target period and reflect them in the spreadsheet
  const events = CalendarApp.getEvents(start_date,end_date);
  const event_length = events.length;

  //range of the cells where event information is reflected.
  const start_cell = pre_values.start_cell;
  const id_column = pre_values.id_column;
  const head_row = pre_values.head_row;
  const end_cell = `${id_column}${parseInt(head_row) + parseInt(event_length)}`;

  //clear the content and border of the table
    data_clear_(pre_values);

  try {
    if(event_length == 0){
      Browser.msgBox("No event is found during the designated period.");
    } else {
      const contents = [];
      for (i=0; i < event_length; i++) {
        let start = events[i].getStartTime();
        let start_date = (start.getMonth() + 1).toString().padStart(2, '0') + '/' + start.getDate().toString().padStart(2, '0');
        let start_time = start.getHours().toString().padStart(2, '0') + ':' + start.getMinutes().toString().padStart(2, '0');    
        let end = events[i].getEndTime();
        let end_date = (end.getMonth() + 1).toString().padStart(2, '0') + '/' + end.getDate().toString().padStart(2, '0');
        let end_time = end.getHours().toString().padStart(2, '0') + ':' + end.getMinutes().toString().padStart(2, '0'); 
        let title = events[i].getTitle();
        let location = events[i].getLocation();
        let description = events[i].getDescription();
        let id = events[i].getId();
        contents.push([start_date,start_time,end_date,end_time,title,location,description,id]);
      }
      
      const range = sheet.getRange(`${start_cell}:${end_cell}`);

      range.setValues(contents);

      const validate_column = pre_values.validate_column;
      set_border_validation_(start_cell,validate_column,head_row,event_length);

      Browser.msgBox("Events during the designated period are displayed on the sheet.");
    }
  } catch (e) {
    console.log('Error in display_events: ' + e.message);
  }
}

/**
 * Displays events that match the keyword designated by the user in Spreadsheet.
 * 
 * This is a private function and used in the `display_events()`.
 * 
 * @param {Object} pre_values The values that the user input.
 */
function display_events_keyword_(pre_values) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Update-Delete');
  // convert period to milliseconds (1 day = 1*24*60*60*1000 milliseconds)
  const period = pre_values.period;
  const period_ms = parseInt(period) * 24 * 60 * 60 * 1000;
  const now = new Date();
  const start_time = new Date(now.getTime() - period_ms); // Subtract period from current date
  // console.log(start_time,period_ms,now);
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(start_time, now);
  const keyword = pre_values.keyword;
  const events_by_keyword = events.filter(function(event) {
    return event.getTitle().toLowerCase().includes(keyword.toLowerCase());
  });
  const event_length = events_by_keyword.length;

  //range of the cells where event information is reflected.
  const start_cell = pre_values.start_cell;
  const id_column = pre_values.id_column;
  const head_row = pre_values.head_row;
  const end_cell = `${id_column}${parseInt(head_row) + parseInt(event_length)}`;

  //clear the content and border of the table
  data_clear_(pre_values);

  try{
    if(event_length == 0){
      Browser.msgBox("No event is matched to the keyword.");
    } else {
      const contents = [];
      for(i = 0; i < events_by_keyword.length; i++) {
        let start = events[i].getStartTime();
        let start_date = start.getMonth()+1 + '/' + start.getDate();
        let start_time;
        if(start.getMinutes() === 0 ){
          start_time = start.getHours() + ':' + '00';
        } else {
          start_time = start.getHours() + ':' + start.getMinutes();
        }        
        let end = events[i].getEndTime();
        let end_date = end.getMonth()+1 + '/' + end.getDate();
        let end_time;
        if(end.getMinutes() === 0){
          end_time = end.getHours() + ':' + '00';
        } else {
          end_time = end.getHours() + ':' + end.getMinutes();
        }
        let title = events[i].getTitle();
        let location = events[i].getLocation();
        let description = events[i].getDescription();
        let id = events[i].getId();
        contents.push([start_date,start_time,end_date,end_time,title,location,description,id]);
      }
      // console.log(contents);      
      const range = sheet.getRange(`${start_cell}:${end_cell}`);

      range.setValues(contents);

      const validate_column = pre_values.validate_column;
      set_border_validation_(start_cell,validate_column, head_row, event_length);

      Browser.msgBox('Events that matched the keyword are displayed on the sheet.');
    }
  } catch(e){
    console.log('Error in display_events_by_keyword: ' + e.message);
  }
}

/**
 * Displays events by calling either `display_events_keyword_(pre_values)` or `display_events_period_(pre_values)`
 * 
 * Must set the `manage_event_library` in the target project in advance.
 *
 * Examples:
 * ```
 * library_name.display_events();
 * ```
 * 
 */
function display_events(){
  const pre_values = pre_update_delete_();
  if (!pre_values || !pre_values.start_cell) {
   return;
  } else {
    if(pre_values.hasOwnProperty('keyword')){
      display_events_keyword_(pre_values);
    } else if (pre_values.hasOwnProperty('start')) {
      display_events_period_(pre_values);
    }
  }
}

/**
 * Clears the originally input data out.
 * 
 * This is a private function and used in the `display_events(pre_values)`.
 */
function data_clear_(pre_values) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Update-Delete');
  const start_cell = pre_values.start_cell;
  const validate_column = pre_values.validate_column;
  const last_row = sheet.getRange(start_cell).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
  const end_cell = `${validate_column}${last_row}`;
  const range = sheet.getRange(`${start_cell}:${end_cell}`); 

  if(last_row > 0){
    range.clearContent();
    range.clearDataValidations();
    range.setBorder(true,false,false,false,false,false);
  }
}

/**
 * Set borders for the data range and data validation for a specific column.
 * 
 * This is a private function and used in the `display_events(pre_values)`.
 */
function set_border_validation_(start_cell, validate_column, head_row, event_length){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Update-Delete');
  const end_cell = `${validate_column}${parseInt(head_row) + parseInt(event_length)}`;
  

  const border_range = sheet.getRange(`${start_cell}:${end_cell}`);
  border_range.setBorder(true,true,true,true,true,true);

  const validate_first_cell = `${validate_column}${parseInt(head_row) + 1}`;
  const validate_range = sheet.getRange(`${validate_first_cell}:${end_cell}`);
  const values = ["Update", "Delete"];
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
  validate_range.setDataValidation(rule);
}

/**
 * Update or Delete designated events in Google Calendar.
 */
function update_delete_events(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Update-Delete');

  const search_start = sheet.getDataRange().createTextFinder("開始日時").findNext();
  const start_cell = search_start.getA1Notation();
  const last_row = sheet.getRange(start_cell).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
  const search_end = sheet.getDataRange().createTextFinder("Update/Delete").findNext();
  const head_end_cell = search_end.getA1Notation();
  const head_end_column = head_end_cell.match(/[A-Z]+/)[0];
  const end_cell = `${head_end_column}${last_row}`;
  const data = sheet.getRange(`${start_cell}:${end_cell}`).getValues(); 
  
  try{
      for(each_data of data){
          if(each_data[8] === 'Update'){
            event = CalendarApp.getDefaultCalendar().getEventById(each_data[7]);
            let start_date = new Date(each_data[0]);
            let s_hours = each_data[1].getHours();
            let s_minutes = each_data[1].getMinutes();
            start_date.setHours(s_hours,s_minutes);
            let end_date = new Date(each_data[2]);
            let e_hours = each_data[3].getHours();
            let e_minutes = each_data[3].getMinutes();
            end_date.setHours(e_hours,e_minutes);
            event.setTime(start_date,end_date);
            event.setTitle(each_data[4]);
            event.setLocation(each_data[5]);
            event.setDescription(each_data[6]);
          } else if (each_data[8] == 'Delete') {
            CalendarApp.getDefaultCalendar().getEventById(each_data[7]).deleteEvent();
          }
      }
  }catch(e){
      Browser.msgBox(`Error updating/deleting: ${e.message}`,Browser.Buttons.OK_CANCEL);
  }
  Browser.msgBox("Successfully updated or deleted designated events from Google Calendar.");
}