// NUMERIC CONSTANTS
var MILLIS_PER_MINUTE = 60 * 1000;
var MILLIS_PER_HOUR = 60 * MILLIS_PER_MINUTE;
var MILLIS_PER_DAY = 24 * MILLIS_PER_HOUR;
var MILLIS_PER_WEEK = 7 * MILLIS_PER_DAY;

// OTHER CONSTANTS
var MAX_HOURS_PER_WEEK = 40;
var WARNING_HOURS = 8;
var USER_EMAIL = "liam.connelly@uconn.edu";
var FIRST_PAY_PERIOD = "01/18/2018";
var FIRST_LISTED_PAY_PERIOD = "07/19/2019";

// ADD CUSTOM MENU
function onOpen(e) {
  
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Custom Functions")
  .addItem("Update Hours", 'onEveryHour')
  .addToUi();
  
}

// TRIGGERED HOURLY
function onEveryHour() {
  
  updateHoursCount(new Date());
  
}

// UPDATES HOURS COUNT FOR CURRENT 1/2 PERIOD AND CHECK HOURS WORKED PER WEEK
function updateHoursCount(currDate) {
  
  // GET SPREADSHEET CELL REFERENCES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Hours Counter");
  
  // GET TIMEZONE
  var timezone = Session.getScriptTimeZone();
  
  // GET LAST ROW
  var lastRow = sheet.getDataRange().getLastRow();
  
  // GET CURRENT TIME
  var currDateTime = currDate.getTime();
  var currDateDay = currDate.getDay();
  var currDateHour = currDate.getUTCHours() - currDate.getTimezoneOffset()/60;
  
  // GET MOST RECENT FRIDAY @ TIME = MIDNIGHT
  var beginPeriod = new Date(currDateTime);
  beginPeriod.setDate(currDate.getDate() - mod((currDateDay + 2),7));
  beginPeriod.setHours(0,0,0,0);
  
  // GET NEXT THURSDAY @ TIME = MIDNIGHT - 1 ms
  var endPeriod = new Date(beginPeriod.getTime() - 1);
  endPeriod.setDate(endPeriod.getDate() + 7);
  
  // GET LAST SUNDAY @ TIME = MIDNIGHT
  var beginWeek = new Date(currDateTime);
  beginWeek.setDate(currDate.getDate() - currDateDay);
  beginWeek.setHours(0,0,0,0);
  
  // GET NEXT SATURDAY @ TIME = MIDNIGHT - 1 ms
  var endWeek = new Date(beginWeek.getTime() - 1);
  endWeek.setDate(beginWeek.getDate() + 6);
  
  // FORMAT 1/2 PAY PERIOD RANGE STRINGS
  var periodString = Utilities.formatDate(beginPeriod, timezone, "MM/dd/yy")
  .concat(" - ")
  .concat(Utilities.formatDate(endPeriod, timezone, "MM/dd/yy"));
  
  // IF ROW IS INCORRECT, STOP FUNCTION
  if (sheet.getRange(lastRow,1).getValue() != periodString) return;
  
  // GET SHIFTS INFO FROM LAST THURSDAY - TODAY
  var shiftsInfoHalfPeriod = getShiftsInfo(beginPeriod,currDate);
  var EIRCTimeSheet = shiftsInfoHalfPeriod["EIRCTimeSheet"];
  var EIRCHoursHalfPeriod = shiftsInfoHalfPeriod["EIRCHours"];
  var ITSHoursHalfPeriod = shiftsInfoHalfPeriod["ITSHours"];
  
  // GET ADDITIONAL SHIFTS INFO FROM LAST THURSDAY - TODAY
  var addShiftsInfoHalfPeriod = getAddShiftsInfo(beginPeriod,currDate);
  var addEIRCTimeSheet = addShiftsInfoHalfPeriod["EIRCTimeSheet"];
  var addEIRCHoursHalfPeriod = addShiftsInfoHalfPeriod["EIRCHours"];
  var addITSHoursHalfPeriod = addShiftsInfoHalfPeriod["ITSHours"];
  
  // FORMAT ROW FOR CURRENT 1/2 PAY PERIOD
  var sheetRow = [periodString].concat(addElements(EIRCTimeSheet,addEIRCTimeSheet))
  .concat(EIRCHoursHalfPeriod+addEIRCHoursHalfPeriod)
  .concat(ITSHoursHalfPeriod+addITSHoursHalfPeriod)
  .concat(EIRCHoursHalfPeriod+addEIRCHoursHalfPeriod+ITSHoursHalfPeriod+addITSHoursHalfPeriod);
  
  // UPDATE ROW ON SHEET
  sheet.getRange(lastRow,1,1,sheetRow.length).setValues([sheetRow]);
  
  // GET SHIFTS INFO FROM LAST SUNDAY - NEXT SATURDAY -- UNUSED
  var shiftsInfoWeek = getShiftsInfo(beginWeek,currDate);
  var EIRCHoursWeek = shiftsInfoWeek["EIRCHours"];
  var ITSHoursWeek = shiftsInfoWeek["ITSHours"];
  
  // FIND REMAINING HOURS LEFT TO WORK THIS HALF PAY PERIOD
  var hoursRemaing = Math.floor(4*(MAX_HOURS_PER_WEEK-EIRCHoursHalfPeriod-addEIRCHoursHalfPeriod-ITSHoursHalfPeriod-addITSHoursHalfPeriod))/4;
  
  // IF NEGATIVE REMAINING HOURS, SEND OVERTIME EMAIL
  if (currDateHour==5 && hoursRemaing<0) {
    
    MailApp.sendEmail({
      to: USER_EMAIL,
      subject: "OVERTIME HOURS ALERT",
      htmlBody: ("As of ").concat(Utilities.formatDate(currDate, timezone, "MM/dd/yy hh:mm a"))
      .concat(" you have exceeded your maximum work hours of ")
      .concat(MAX_HOURS_PER_WEEK.toString())
      .concat(" hours this week. You have worked <b>")
      .concat((-1*hoursRemaing).toString())
      .concat(" hours</b> overtime.")
    });
  
  // IF <=WARNING_HOURS HOURS, SEND WARNING EMAIL
  } else if (currDateHour==5 && hoursRemaing<=WARNING_HOURS) {
    
    MailApp.sendEmail({
      to: USER_EMAIL,
      subject: "Approaching Max Hours Per Week",
      htmlBody: ("As of ").concat(Utilities.formatDate(currDate, timezone, "MM/dd/yy hh:mm a"))
      .concat(" you are close to working your maximum work time of ")
      .concat(MAX_HOURS_PER_WEEK.toString())
      .concat(" hours this week. You may only work <b>")
      .concat(hoursRemaing.toString())
      .concat(" more hour(s)</b> before ")
      .concat(Utilities.formatDate(new Date(endPeriod.getTime()+1), timezone, "MM/dd/yy"))
      .concat(".")
    });
    
  }
}

// TRIGGERED @3:00-4:00 AM FRIDAY MORNINGS
function onFridayMorning() {
  
  startNewRow(new Date());
  
}

// CREATES NEW ROW FOR WEEK AND FINALIZES PREVIOUS ROW INFORMATION
function startNewRow(currDate) {
  
  // GET SPREADSHEET CELL REFERENCES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Hours Counter");
  
  // GET TIMEZONE
  var timezone = Session.getScriptTimeZone();
  
  // GET SPREADSHEET CELL REFERENCES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Hours Counter");
  
  // GET LAST ROW
  var lastRow = sheet.getDataRange().getLastRow();
  
  // GET CURRENT TIME
  var currDateTime = currDate.getTime();
  var currDateDay = currDate.getDay();
  var currDateHour = currDate.getUTCHours() - currDate.getTimezoneOffset()/60;
  
  // IF NOT FRIDAY,STOP FUNCTION
  if (currDateDay != 5) return;
  
  // GET LAST FRIDAY (I.E. TODAY) @ TIME = MIDNIGHT
  var beginPeriod = new Date(currDateTime);
  beginPeriod.setDate(currDate.getDate() - mod((currDateDay + 2),7))
  
  // GET NEXT THURSDAY @ TIME = MIDNIGHT - 1 ms
  var endPeriod = new Date(beginPeriod.getTime() - 1);
  endPeriod.setDate(endPeriod.getDate() + 7);
  
  // GET LAST FRIDAY (I.E. A WEEK AGO) @ TIME = MIDNIGHT
  var prevBeginPeriod = new Date(beginPeriod.getTime());
  prevBeginPeriod.setDate(prevBeginPeriod.getDate() - 7);
  
  // GET LAST LAST THURSDAY (I.E. 8 DAYS AGO) @ TIME = MIDNIGHT - 1 ms
  var prevEndPeriod = new Date(endPeriod.getTime());
  prevEndPeriod.setDate(endPeriod.getDate() - 7);
  
  // FORMAT 1/2 PAY PERIOD RANGE STRINGS
  var periodString = Utilities.formatDate(beginPeriod, timezone, "MM/dd/yy")
  .concat(" - ")
  .concat(Utilities.formatDate(endPeriod, timezone, "MM/dd/yy"));
  
  var prevPeriodString = Utilities.formatDate(prevBeginPeriod, timezone, "MM/dd/yy")
  .concat(" - ")
  .concat(Utilities.formatDate(prevEndPeriod, timezone, "MM/dd/yy"));
  
  // IF LAST ROW IS NOT PAST 1/2 PAY PERIOD, STOP FUNCTION
  if (sheet.getRange(lastRow,1).getValue() != prevPeriodString) return;
  
  // GET SHIFTS INFO FROM LAST 1/2 PAY PERIOD
  var shiftsInfoHalfPeriod = getShiftsInfo(prevBeginPeriod,prevEndPeriod);
  var EIRCTimeSheet = shiftsInfoHalfPeriod["EIRCTimeSheet"];
  var EIRCHoursHalfPeriod = shiftsInfoHalfPeriod["EIRCHours"];
  var ITSHoursHalfPeriod = shiftsInfoHalfPeriod["ITSHours"];
  
  // GET ADDITIONAL SHIFTS INFO FROM LAST THURSDAY - TODAY
  var addShiftsInfoHalfPeriod = getAddShiftsInfo(prevBeginPeriod,prevEndPeriod);
  var addEIRCTimeSheet = addShiftsInfoHalfPeriod["EIRCTimeSheet"];
  var addEIRCHoursHalfPeriod = addShiftsInfoHalfPeriod["EIRCHours"];
  var addITSHoursHalfPeriod = addShiftsInfoHalfPeriod["ITSHours"];
  
  // FORMAT ROW FOR CURRENT 1/2 PAY PERIOD
  var sheetRow = [prevPeriodString].concat(addElements(EIRCTimeSheet,addEIRCTimeSheet))
  .concat(EIRCHoursHalfPeriod+addEIRCHoursHalfPeriod)
  .concat(ITSHoursHalfPeriod+addITSHoursHalfPeriod)
  .concat(EIRCHoursHalfPeriod+addEIRCHoursHalfPeriod+ITSHoursHalfPeriod+addITSHoursHalfPeriod);
  
  // UPDATE ROW ON SHEET AND FORMAT COLOR
  sheet.getRange(lastRow,1,1,sheetRow.length).setValues([sheetRow]);
  sheet.getRange(lastRow,1,1,sheetRow.length).setBackground(getRowAltColor(prevPeriodString) ? "#e8e7fc" : "#ffffff");
  
  // ADD NEW DATE STRING FOR CURRENT 1/2 PAY PERIOD AND FORMAT COLOR
  sheet.getRange(lastRow+1,1).setValue(periodString);
  sheet.getRange(lastRow+1,1,1,sheetRow.length).setBackground(getRowAltColor(periodString) ? "#e8e7fc" : "#ffffff");
  
}

// TRIGGERED @3:00-4:00 AM SUNDAY MORNINGS
function onSundayMorning() {
  
  sendSummaryEmail(new Date());
  
}

// SEND SUMMARY EMAIL WITH HOURS WORKED INFO
function sendSummaryEmail(currDate) {
  
  if (!currDate) {
    var currDate = new Date();
  }
  
  // GET SPREADSHEET CELL REFERENCES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Hours Counter");
  
  // GET TIMEZONE
  var timezone = Session.getScriptTimeZone();
  
  // GET LAST ROW
  var lastRow = sheet.getDataRange().getLastRow();
  
  // GET CURRENT TIME
  var currDateTime = currDate.getTime();
  var currDateDay = currDate.getDay();
  var currDateHour = currDate.getUTCHours() - currDate.getTimezoneOffset()/60;
  
  // GET MOST RECENT SUNDAY @ TIME = MIDNIGHT
  var beginWeek = new Date(currDateTime);
  beginWeek.setDate(currDate.getDate() - currDateDay);
  
  // GET LAST SUNDAY (I.E OVER A WEEK AGO) @ TIME = MIDNIGHT
  var beginPrevWeek = new Date(beginWeek.getTime());
  beginPrevWeek.setDate(beginPrevWeek.getDate() - 7);
  
  // GET LAST SATURDAY @ TIME = MIDNIGHT - 1 ms
  var endPrevWeek = new Date(beginPrevWeek.getTime() - 1);
  endPrevWeek.setDate(endPrevWeek.getDate() + 6)
  
  // GET SHIFTS INFO FROM LAST SUNDAY - LAST SATURDAY
  var shiftsInfoWeek = getShiftsInfo(beginPrevWeek,endPrevWeek);
  var EIRCHoursWeek = shiftsInfoWeek["EIRCHours"];
  var ITSHoursWeek = shiftsInfoWeek["ITSHours"];
  
  // GET ADDITIONAL SHIFTS INFO FROM LAST SUNDAY - LAST SATURDAY
  var addShiftsInfoWeek = getAddShiftsInfo(beginPrevWeek,endPrevWeek);
  var addEIRCHoursWeek = addShiftsInfoWeek["EIRCHours"];
  var addITSHoursWeek = addShiftsInfoWeek["ITSHours"];
  
  // ADD SHIFTS INFO AND ADDITIONAL SHIFTS INFO
  var totalEIRCHoursWeek = EIRCHoursWeek + addEIRCHoursWeek;
  var totalITSHoursWeek = ITSHoursWeek + addITSHoursWeek;
  
  // FORMAT WEEK STRING
  var prevWeekString = Utilities.formatDate(beginPrevWeek, timezone, "MM/dd/yy")
  .concat(" - ")
  .concat(Utilities.formatDate(endPrevWeek, timezone, "MM/dd/yy"));
  
  // SEND MAIL WITH HOURS WORKED INFO
  MailApp.sendEmail({
    to: "liam.connelly@uconn.edu",
    subject: "Weekly Hours Report",
    htmlBody: ("During the previous week of ")
    .concat(prevWeekString)
    .concat(", you worked a total of <b>")
    .concat((totalEIRCHoursWeek+totalITSHoursWeek).toString())
    .concat(" hours:</b><ul><li><b>")
    .concat(totalEIRCHoursWeek.toString())
    .concat(" hours</b> at the EIRC lab</li><li><b>")
    .concat(totalITSHoursWeek.toString())
    .concat(" hours</b> at ITS </li></ul>")
  });
  
}

// CREATE AND REBUILD SPREADSHEETS, AND POPULATE WITH ALL DATA
function rebuildSpreadsheet() {
 
  // GET SPREADSHEET CELL REFERENCES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var all_sheets = ss.getSheets();
  var hours_sheet = ss.getSheetByName("Hours Counter");
  var add_shifts_sheet = ss.getSheetByName("Additional Entries");
  
  // GET TIMEZONE
  var timezone = Session.getScriptTimeZone();
  
  // GET CURRENT TIME
  var currDate = new Date();
  var currDateTime = currDate.getTime();
  var currDateDay = currDate.getDay();
  var currDateHour = currDate.getUTCHours() - currDate.getTimezoneOffset()/60;
  
  // GET MOST RECENT FRIDAY @ TIME = MIDNIGHT
  var beginPeriod = new Date(currDateTime);
  beginPeriod.setDate(currDate.getDate() - mod((currDateDay + 2),7));
  beginPeriod.setHours(0,0,0,0);
  
  // GET NEXT THURSDAY @ TIME = MIDNIGHT - 1 ms
  var endPeriod = new Date(beginPeriod.getTime() - 1);
  endPeriod.setDate(beginPeriod.getDate() + 6);
  
  // GET LAST SUNDAY @ TIME = MIDNIGHT
  var beginWeek = new Date(currDateTime);
  beginWeek.setDate(currDate.getDate() - currDateDay);
  beginWeek.setHours(0,0,0,0);
  
  // GET NEXT SATURDAY @ TIME = MIDNIGHT - 1 ms
  var endWeek = new Date(beginWeek.getTime() - 1);
  endWeek.setDate(currDate.getDate() + 7);
  
  // HEADERS FOR SHEETS
  var hours_sheet_header = ["Week", "EIRC - F", "EIRC - Sa", "EIRC - Su", "EIRC - M", "EIRC - Tu", "EIRC - W", "EIRC - Th", "EIRC Hours", "ITS Hours", "Total Hours"];
  var add_shifts_sheet_header = ["Date", "Job", "Reason", "Hours"];
  
  // IF HOURS SHEET IS MISSING OR HAS AN INCORRECT HEADER, REBUILD IT
  if (!hours_sheet || !arrayEqual(hours_sheet.getDataRange().getValues()[0],hours_sheet_header)) {
    
    // IF HOURS HEADER IS INCORRECT, REMOVE ALL FORMATTING
    if (hours_sheet && !arrayEqual(hours_sheet.getDataRange().getValues()[0],hours_sheet_header)) {
      hours_sheet.clear()
    }
    
    // IF HOURS SHEET DOES NOT EXIST, ADD IT IN FIRST POSITION
    if (!hours_sheet) {
      hours_sheet = ss.insertSheet("Hours Counter",0);
    }
  
    // SET HOURS HEADER
    hours_sheet.getRange(1,1,1,11).setValues([hours_sheet_header])
    
    // MOVE FREEZE BAR
    hours_sheet.setFrozenRows(1);
  
    // SET COLUMNS WIDTHS
    hours_sheet.setColumnWidth(1,160);
    hours_sheet.setColumnWidths(2,10,80);
    
    // SET ROW HEIGHTS
    hours_sheet.setRowHeights(1,hours_sheet.getMaxRows(),21);
  
    // ALIGN TEXT
    hours_sheet.getRange(1,1,hours_sheet.getMaxRows(),11).setHorizontalAlignment("center").setVerticalAlignment("middle");
    
    // IF EXTRA COLUMNS EXIST, TRIM THEM
    if (hours_sheet.getMaxColumns()>12) {
      hours_sheet.deleteColumns(12, hours_sheet.getMaxColumns()-11);
    }
  
    // SET HOURS HEADER ROW AND TOTAL COLUMNS TO BE BOLDED TEXT
    hours_sheet.getRange(1,1,1,11).setFontWeight("bold");
    hours_sheet.getRange(1,9,hours_sheet.getMaxRows(),3).setFontWeight("bold");
  
    // SET HOURS HEADER COLOR
    hours_sheet.getRange(1,1,1,11).setBackground("#8989eb");
    
  }
  
  // IF ROWS ARE ALREADY ON SHEET, CLEAR THEM
  if (hours_sheet.getLastRow()>1) {
    hours_sheet.deleteRows(2,hours_sheet.getLastRow()-1);
  }
  
  // IF ADD SHIFTS SHEET IS MISSING, REBUILD IT
  if (!add_shifts_sheet) {
    
    add_shifts_sheet = ss.insertSheet("Additional Entries",1);
    
    // SET ADD SHFITS HEADER
    add_shifts_sheet.getRange(1,1,1,4).setValues([add_shifts_sheet_header])
    
    // SET COLUMNS WIDTHS
    add_shifts_sheet.setColumnWidth(1,150);
    add_shifts_sheet.setColumnWidths(2,3,100);
    add_shifts_sheet.setColumnWidth(3,300);
    
    // SET ROW HEIGHTS
    add_shifts_sheet.setRowHeights(1,add_shifts_sheet.getMaxRows(),21);
  
    // ALIGN TEXT
    add_shifts_sheet.getRange(1,1,add_shifts_sheet.getMaxRows(),11).setHorizontalAlignment("center").setVerticalAlignment("middle");
    
    // IF EXTRA COLUMNS EXIST, TRIM THEM
    if (add_shifts_sheet.getMaxColumns()>4) {
      add_shifts_sheet.deleteColumns(5, add_shifts_sheet.getMaxColumns()-4);
    }
  
    // SET ADD SHIFTS HEADER ROW TO BE BOLDED TEXT
    add_shifts_sheet.getRange(1,1,1,4).setFontWeight("bold");
    
    // SET ALTERNATING COLOR
    add_shifts_sheet.getRange(1,1,add_shifts_sheet.getMaxRows(),add_shifts_sheet.getMaxColumns())
    .applyRowBanding(SpreadsheetApp.BandingTheme.INDIGO);
    
    // CREATE NEW DATA VALIDATION FOR 'JOB' COLUMN
    var dataValidRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['EIRC','ITS'])
    .setAllowInvalid(false).build();
    
    // ADD DATA VALIDATION RULE
    add_shifts_sheet.getRange(2,2,add_shifts_sheet.getMaxRows()-1).setDataValidation(dataValidRule);
    
  }
  
  // GET LIST OF BEGIN AND END WEEKS FOR EACH 1/2 PAY PERIOD
  var dateList = getListofDates(new Date(FIRST_LISTED_PAY_PERIOD), beginPeriod);
  
  // CREATE FIRST 1/2 PAY PERIOD STRING
  var firstPeriodString = Utilities.formatDate(dateList[0][0], timezone, "MM/dd/yy")
  .concat(" - ")
  .concat(Utilities.formatDate(dateList[0][1], timezone, "MM/dd/yy"));
  
  // ADD 1/2 PERIOD STRING TO SHEET
  hours_sheet.getRange(2,1).setValue(firstPeriodString);
  
  // NEEDED TO VERIFY THAT PERIOD STRING WAS ADDED TO SHEET
  SpreadsheetApp.flush();
  
  // ITERATE THROUGH EACH 1/2 PAY PERIOD AND CREATE ROW WITH HOURS INFO
  for (var i=1; i<dateList.length; i++) {
    
    startNewRow(dateList[i][0]);
    
  }
  
  // UPDATE HOURS OF FINAL 1/2 PAY PERIOD TO CURRENT DATE
  updateHoursCount(currDate);
  
  // REMOVE ANY OTHER SHEETS
  for (var i=0; i<all_sheets.length; i++) {
    
    if (!(all_sheets[i].getName() == "Hours Counter" || all_sheets[i].getName() == "Additional Entries")) {
    
      ss.deleteSheet(all_sheets[i]);
    
    }
  }
  
  // BRING HOURS SHEET TO FRONT
  ss.setActiveSheet(hours_sheet);
  ss.moveActiveSheet(1)
  
}

// GIVEN A PERIOD OF TIME, COLLECT TIMESHEET AND HOURS WORKED
function getShiftsInfo(beginPeriod,endPeriod) {
  
  // GET CALENDAR IDS
  try {
    
    var eirc_cal_id = CalendarApp.getCalendarsByName("Lab Work")[0].getId();
    var its_cal_id = CalendarApp.getCalendarsByName("ITS Help Center")[0].getId();
    
  } catch (e) {
    
    throw new Error("One or more calendars not found");
    
  }
  
  // GET NON-INTERLAPPING SHIFTS
  var shifts = getValidShifts(beginPeriod,endPeriod);
  
  var EIRCHours = 0;
  var ITSHours = 0;
  var EIRCTimeSheet = []; for (var i=0;i<7;i++) EIRCTimeSheet[i] = 0;
  
  // SUM EIRCHours and ITSHours THROUGH EACH SHIFT
  for (var i=0; i<shifts.length; i++) {
    
    // GET JOB TYPE AND PAYED TIME (15 MIN BLOCKS), UNLESS OTHERWISE NOTED
    var shiftJob = (shifts[i][2].getOriginalCalendarId() == eirc_cal_id) ? "EIRC" : "ITS";
    var payedLength = getPayedLength(shifts[i]);
    
    // ADDITIONAL INFO SAVED FOR FURTHER USE
    shifts[i][3] = payedLength;
    shifts[i][4] = shiftJob;
    shifts[i][5] = shifts[i][2].getTitle();
    
    if (shiftJob == "EIRC") {
      
      EIRCHours += payedLength;
      EIRCTimeSheet[mod((shifts[i][2].getEndTime().getDay() + 2),7)] += payedLength;
      
    } else {
      
      ITSHours += payedLength;
      
    }
  }
  
  var shiftsInfo = [];
  shiftsInfo["EIRCTimeSheet"] = EIRCTimeSheet;
  shiftsInfo["EIRCHours"] = EIRCHours;
  shiftsInfo["ITSHours"] = ITSHours;
  
  return shiftsInfo;
  
}

// GET NON-OVERLAPPING SHIFTS FROM CALENDARS AND PARSE
function getValidShifts(begin,end) {
  
  
  // GET CALENDAR IDS
  try {
    
    var eirc_cal_id = CalendarApp.getCalendarsByName("Lab Work")[0].getId();
    var its_cal_id = CalendarApp.getCalendarsByName("ITS Help Center")[0].getId();
    
  } catch (e) {
    
    throw new Error("One or more calendars not found");
    
  }
  
  // GET CALENDAR REFERENCES
  var EIRCCal = CalendarApp.getCalendarById(eirc_cal_id);
  var ITSCal = CalendarApp.getCalendarById(its_cal_id);
  
  var EIRCShifts = EIRCCal.getEvents(begin, end);
  var ITSShifts = ITSCal.getEvents(begin, end);
  var allShifts = EIRCShifts.concat(ITSShifts);
  
  var shiftsInfo = [];
  
  if (allShifts.length == 0) return shiftsInfo;
  
  for (var i=allShifts.length-1; i>=0; i--) {
     
    var currShift = allShifts[i];
    var currDesc = currShift.getDescription();
    
    // IF "UNPAYED" IN EVENT DESCRIPTION, DROP THE SHIFT FROM CALCULATIONS
    if (currDesc.length && currDesc.toUpperCase().indexOf("UNPAYED") != -1) allShifts.splice(i,1);
    
  }
  
  for (var i=0; i<allShifts.length; i++) {
    
    var currShift = allShifts[i];
    
    // ADJUST START AND STOP TIMES IF AN SHIFT EXCEEDS THE PERIOD OF CALCULATION
    var startTimeUTC = Math.max(currShift.getStartTime().getTime(),begin.getTime())
    var endTimeUTC = Math.min(currShift.getEndTime().getTime(),end.getTime());
    
    shiftsInfo[i] = [startTimeUTC,endTimeUTC,currShift];
    
  }
  
  // SORT SHIFTS BY START TIME
  shiftsInfo.sort(compareFirstCol);
  
  for (var i=shiftsInfo.length-1; i>=0; i--) {
    
    var currSh = shiftsInfo[i];
    
    // IF SHIFT IS COMPLETELY CONTAINED BY ANOTHER SHIFT, DROP IT
    if (i!=0 && shiftsInfo[i][0]>=shiftsInfo[i-1][0] && shiftsInfo[i][1]<=shiftsInfo[i-1][1]) {
      
      shiftsInfo.splice(i,1);
      
    // IF (FIRST) SHIFT IS COMPLETELY CONTAINED BY ANOTHER SHIFT, DROP IT  
    } else if (i==0 && shiftsInfo.length!=1 && shiftsInfo[i][0]>=shiftsInfo[i+1][0] && shiftsInfo[i][1]<=shiftsInfo[i+1][1]) {
      shiftsInfo.splice(i,1);
      
    // OTHERWISE, KEEP SHIFT AND TRUNCATE ITS START TIME AT THE PREVIOUS SHIFT'S END TIME, IF OVERLAPPING
    } else if (i!=0) {
      
      shiftsInfo[i][0] = Math.max(shiftsInfo[i][0],shiftsInfo[i-1][1])
      
    }
    
  }
  
  return shiftsInfo;
  
}

// FIND CORRECT ROW COLOR, FORMATTED SO COLOR ALTERNATES BETWEEN PAY PERIODS - RETURNS 0 OR 1
function getRowAltColor(row) {
  
  // GET SPREADSHEET CELL REFERENCES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Hours Counter");
  
  // IF INPUT IS INTEGER, FIND dateString ON THE SPREADSHEET
  if (!isNaN(row)) {
  
    var lastRow = sheet.getDataRange().getLastRow();
  
    if (row>lastRow) return -1;
  
    var dateString = sheet.getRange(row,1).getValue();
    
  // IF INPUT IS dateString, ASSIGN
  } else {
    
    var dateString = row;
    
  }
  
  // FIND START AND STOP DATES, AND STOP FUNCTION IF ERROR
  var beginEndStrings = dateString.split(" - ");
  
  if (beginEndStrings.length == 1) return -1;
  
  var beginPeriod = new Date(beginEndStrings[0]);
  var endPeriod = new Date(beginEndStrings[1]);
  
  // IF DATE IS ROUNDED DOWN TO 1900s, CORRECT BY ADDING 100 YEARS
  beginPeriod = fixYear(beginPeriod);
  endPeriod = fixYear(endPeriod);
  
  // FIND WEEKS SINCE FIRST PAY PERIOD
  var weeksSince = (endPeriod.getTime() - new Date(FIRST_PAY_PERIOD).getTime()) / MILLIS_PER_WEEK;
  
  // IF UNEXPECTED RESULT, STOP FUNCTION
  if (Math.abs(weeksSince - Math.round(weeksSince))>.2) return -1; else weeksSince = Math.round(weeksSince);
  
  // RETURN 0 OR 1 BASED ON PAY PERIOD NUMBER, ALTERNATING
  return mod(Math.ceil(weeksSince/2),2);
  
}

// FIND "PAYED AS:" TAG IN SHIFT DESCRIPTION OR ELSE CALCUALTE SHIFT LENGTH
function getPayedLength(shift) {
  
  // CONVERT DESCRIPTION TO UPPER CASE WITH NO SPACES
  var desc = shift[2].getDescription().toUpperCase().replace(/\s/g,"");
  
  // FIND LOCATION OF "PAYED AS:" TAGE
  var payedAsTagIndex = desc.indexOf("PAYEDAS:");
  
  // IF "PAYED AS:" TAG EXISTS, PARSE NUMBER
  if (payedAsTagIndex != -1) {
    
    // REMOVE DESCRIPTION BEFORE "PAYED AS:" TAG
    var payedLength = desc.substring(payedAsTagIndex + "PAYEDAS:".length);
    
    // REMOVE ANY NON-NUMERIC (DECIMAL EXCLUDED) CHARACTERS AFTER "PAYED AS:" TAG
    if (payedLength.search(/[^0-9\.]/) != -1) {
      
       payedLength = payedLength.substring(0,payedLength.search(/[^0-9\.]/));
      
    }
    
    // CONVERT TO FLOAT
    payedLength = parseFloat(payedLength);
    
  // OTHERWISE CREATE payedAs VALUE FROM SHIFT LENGTH
  } else {
    
    // FIND payedAs VALUE AS SHIFT LENGTH, ROUNDED TO 15 MINUTES
    payedLength = Math.round((shift[1]-shift[0])/(MILLIS_PER_HOUR/4))/4;
  
  }
  
  return payedLength;
  
}

// GIVEN A PERIOD OF TIME, COLLECT TIMESHEET AND HOURS WORKED FROM ADD SHIFTS SHEET
function getAddShiftsInfo(beginPeriod,endPeriod) {
  
  // GET SPREADSHEET CELL REFERENCES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Hours Counter");
  var altSheet = ss.getSheetByName("Additional Entries");
  
  var EIRCHours = 0;
  var ITSHours = 0;
  var EIRCTimeSheet = []; for (var i=0;i<7;i++) EIRCTimeSheet[i] = 0;
  
  // GET ADDITIONAL SHIFTS INFO
  var addEntries = altSheet.getDataRange().getValues();
  
  // GO THROUGH ADDITIONAL SHIFTS AND ADD IF IN TIME RANGE
  for (i=addEntries.length-1;i>0;i--) {
    
    var entryDate = new Date(addEntries[i][0]);
    var entryJob = addEntries[i][1];
    var entryLength = addEntries[i][3];
    
    // IF DATE IS ROUNDED DOWN TO 1900s, CORRECT BY ADDING 100 YEARS
    entryDate = fixYear(entryDate);
    
    // SORT BY JOB IF SHIFT IS IN TIME RANGE
    if (dateInRange(entryDate,beginPeriod,endPeriod)) {
      
      if (entryJob == 'EIRC') {
        
        EIRCTimeSheet[mod((entryDate.getDay() + 2),7)] += entryLength;
        EIRCHours += entryLength;
        
      } else if (entryJob == 'ITS') {
        
        ITSHours += entryLength;
        
      }
    } 
  }
  
  var addShiftsInfo = [];
  addShiftsInfo["EIRCTimeSheet"] = EIRCTimeSheet;
  addShiftsInfo["EIRCHours"] = EIRCHours;
  addShiftsInfo["ITSHours"] = ITSHours;
  
  return addShiftsInfo;
  
}

// CHECK IF DATE IS BETWEEN TWO OTHER DATES, INCLUSIVE
function dateInRange(entryDate,beginPeriod,endPeriod) {
  
  var entryDateTime = entryDate.getTime();
  var beginPeriodTime = beginPeriod.getTime();
  var endPeriodTime = endPeriod.getTime();
  
  return (entryDateTime>=beginPeriodTime) & (entryDateTime<=endPeriodTime);
  
}

// GET LIST OF BEGIN/END DATES FOR EACH WEEK BETWEEN TWO DATES
function getListofDates(beginDate,endDate) {
  
  // GET TIMEZONE
  var timezone = Session.getScriptTimeZone();
  
  // IF DATE IS ROUNDED DOWN TO 1900s, CORRECT BY ADDING 100 YEARS
  beginDate = fixYear(beginDate);
  endDate = fixYear(endDate);
  
  // IF DATES ARE NOT THE SAME DAY OF THE WEEK, RETURN
  if (beginDate.getDay() != endDate.getDay()) return;
  
  // LIST OF DATE STRINGS
  var dateList = [];
  
  // CREATE FIRST BEGIN AND END DATE
  var currBeginDate = new Date(beginDate.getTime());
  var currEndDate = new Date(currBeginDate.getTime())
  currEndDate.setDate(currBeginDate.getDate()+6)
  
  // ADD BEGIN AND END DATES to dateList
  dateList = dateList.concat([[new Date(currBeginDate.getTime()),new Date(currEndDate.getTime())]]);
  
  // ITERATE THROUGH ALL WEEKS, INCLUDING BEGIN AND END DATE
  while (currBeginDate.getTime() != endDate.getTime()) {
    
    // ADVANCE BEGIN AND END DATES
    currBeginDate.setDate(currBeginDate.getDate() + 7);
    currEndDate.setDate(currEndDate.getDate() + 7);
    
    // ADD BEGIN AND END DATES to dateList
    dateList = dateList.concat([[new Date(currBeginDate.getTime()),new Date(currEndDate.getTime())]]);
    
  }
  
  return dateList;
  
}

// USED FOR SORTING BASED ON FIRST COLUMN
function compareFirstCol(a,b) {
  
  if (a[0]==b[0]) return 0;
  else return (a[0]>b[0]) ? 1 : -1;
  
}

// MOD p OF n
function mod(n, p) {
  
    return n - p * Math.floor(n/p);
  
}

// ADD ARRAYS ELEMENT-WISE
function addElements(a,b) {
  
  if (a.length != b.length) return -1;
  var sum = []
  for (i=0;i<a.length;i++) sum[i] = a[i] + b[i];
  return sum;
  
}

// OBSOLETE - CORRECTS ERROR DUE TO DATE MATH THAT CROSS DAYLIGHT SAVINGS CHANGES
function savingsError(beginDate,endDate) {
  
  var beginDate = beginDate.getTimezoneOffset();
  var endDate = endDate.getTimezoneOffset();
  
  var timezoneDifference = (beginDate - endDate) / 60;
  
  return timezoneDifference
  
}

// TEST IF TWO ARRAYS ARE EQUAL
function arrayEqual(arr1,arr2) {
  
  if (arr1.length != arr2.length) return false;
  
  for (var i=0; i<arr1.length; i++) {
    
    if (arr1[i] != arr2[i]) return false;
    
  }
  
  return true;
  
}

// IF DATE IS ROUNDED DOWN TO 1900s, CORRECT BY ADDING 100 YEARS
function fixYear(date) {
  
  if (date.getFullYear()<1970) {
    date.setFullYear(date.getFullYear() + 100);
  }
  
  return date;
  
}