// NUMERIC CONSTANTS
var MILLIS_PER_MINUTE = 60 * 1000;
var MILLIS_PER_HOUR = 60 * MILLIS_PER_MINUTE;
var MILLIS_PER_DAY = 24 * MILLIS_PER_HOUR;
var MILLIS_PER_WEEK = 7 * MILLIS_PER_DAY;

// OTHER CONSTANTS
var MAX_HOURS_PER_WEEK = 40;
var WARNING_HOURS = 8;

// ADD CUSTOM MENU
function onOpen(e) {

  var ui = SpreadsheetApp.getUi();

  ui.createMenu("Custom Functions")
  .addItem("Update Hours", 'onEveryHour')
  .addToUi();

}

// TRIGGERED HOURLY, UPDATES HOURS COUNT FOR CURRENT 1/2 PERIOD AND CHECK HOURS WORKED PER WEEK
function onEveryHour() {

  // GET SPREADSHEET CELL REFERENCES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Hours Counter");

  // GET TIMEZONE
  var timezone = Session.getScriptTimeZone();

  // GET LAST ROW
  var lastRow = sheet.getDataRange().getLastRow();

  // GET CURRENT TIME
  var currDate = new Date();
  var currDateTime = currDate.getTime();
  var currDateDay = currDate.getDay();
  var currDateHour = currDate.getUTCHours() - currDate.getTimezoneOffset()/60;

  // GET MOST RECENT FRIDAY @ TIME = MIDNIGHT
  var beginPeriod = new Date(new Date(currDateTime).setHours(0,0,0,0) - mod((currDateDay + 2),7) * MILLIS_PER_DAY);
  var beginPeriodTime = beginPeriod.getTime();

  // GET NEXT THURSDAY @ TIME = MIDNIGHT - 1 ms
  var endPeriod = new Date(beginPeriodTime + 7 * MILLIS_PER_DAY - 1);

  // GET LAST SUNDAY @ TIME = MIDNIGHT
  var beginWeek = new Date(new Date(currDateTime).setHours(0,0,0,0) - currDateDay * MILLIS_PER_DAY);

  // GET NEXT SATURDAY @ TIME = MIDNIGHT - 1 ms
  var endWeek = new Date(beginWeek.getTime() + 7 * MILLIS_PER_DAY - 1);

  // CATCH DAYLIGHT SAVINGS TIME-RELATED ISSUES
  var timezoneDifference = savingsError(beginPeriod,endPeriod,timezone);
  var weekTimezoneDifference = savingsError(beginWeek,endWeek);

  if (timezoneDifference != 0) {

    beginPeriod = new Date(beginPeriod.getTime() + timezoneDifference * MILLIS_PER_HOUR)
    beginPeriodTime = beginPeriod.getTime()

  } if (weekTimezoneDifference != 0) {

    beginWeek = new Date(beginWeek.getTime() + timezoneDifference * MILLIS_PER_HOUR)

  }

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
      to: "liam.connelly@uconn.edu",
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
      to: "liam.connelly@uconn.edu",
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

// TRIGGERED @3:00-4:00 AM FRIDAY MORNINGS, CREATES NEW ROW FOR WEEK AND FINALIZES PREVIOUS ROW INFORMATION
function onFridayMorning() {

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

  var currDate = new Date();
  var currDateTime = currDate.getTime();
  var currDateDay = currDate.getDay();

  // IF NOT FRIDAY,STOP FUNCTION
  if (currDateDay != 5) return;

  // GET LAST FRIDAY (I.E. TODAY) @ TIME = MIDNIGHT
  var beginPeriod = new Date(new Date(currDateTime).setHours(0,0,0,0) - mod((currDateDay + 2),7) * MILLIS_PER_DAY);
  var beginPeriodTime = beginPeriod.getTime();

  // GET NEXT THURSDAY @ TIME = MIDNIGHT - 1 ms
  var endPeriod = new Date(beginPeriodTime + 7 * MILLIS_PER_DAY - 1);

  // GET LAST FRIDAY (I.E. A WEEK AGO) @ TIME = MIDNIGHT
  var prevBeginPeriod = new Date(beginPeriodTime - 7 * MILLIS_PER_DAY);

  // GET LAST LAST THURSDAY (I.E. 8 DAYS AGO) @ TIME = MIDNIGHT - 1 ms
  var prevEndPeriod = new Date(beginPeriodTime - 1);

  // CATCH DAYLIGHT SAVINGS TIME-RELATED ISSUES
  var timezoneDifference = savingsError(beginPeriod,endPeriod,timezone);
  var prevTimezoneDifference = savingsError(prevBeginPeriod,prevEndPeriod);

  if (timezoneDifference != 0) {

    endPeriod = new Date(endPeriod.getTime() - timezoneDifference * MILLIS_PER_HOUR)

  } else if (prevTimezoneDifference != 0) {

    prevBeginPeriod = new Date(prevBeginPeriod.getTime() + prevTimezoneDifference * MILLIS_PER_HOUR);

  }

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

// TRIGGERED @3:00-4:00 AM SUNDAY MORNINGS, SEND EMAIL WITH HOURS WORKED INFO
function onSundayMorning() {

  // GET SPREADSHEET CELL REFERENCES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Hours Counter");

  // GET TIMEZONE
  var timezone = Session.getScriptTimeZone();

  var lastRow = sheet.getDataRange().getLastRow();

  // GET CURRENT TIME
  var currDate = new Date("3/15/2020 3:00 AM");
  var currDateTime = currDate.getTime();
  var currDateDay = currDate.getDay();

  // GET LAST SUNDAY (I.E A WEEK AGO) @ TIME = MIDNIGHT
  var beginPrevWeek = new Date(new Date(currDateTime).setHours(0,0,0,0) - (currDateDay + 7) * MILLIS_PER_DAY);

  // GET LAST SATURDAY (I.E. YESTERDAY) @ TIME = MIDNIGHT - 1 ms
  var endPrevWeek = new Date(beginPrevWeek.getTime() + 7 * MILLIS_PER_DAY - 1);

  // CATCH DAYLIGHT SAVINGS TIME-RELATED ISSUES
  var prevWeekTimezoneDifference = savingsError(beginPrevWeek,endPrevWeek);

  if (prevWeekTimezoneDifference != 0) {

    beginPrevWeek = new Date(beginPrevWeek.getTime() + prevWeekTimezoneDifference * MILLIS_PER_HOUR)

  }

  // GET SHIFTS INFO FROM LAST SUNDAY - LAST SATURDAY
  var shiftsInfoWeek = getShiftsInfo(beginPrevWeek,endPrevWeek);
  var EIRCHoursWeek = shiftsInfoWeek["EIRCHours"];
  var ITSHoursWeek = shiftsInfoWeek["ITSHours"];

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
    .concat((EIRCHoursWeek+ITSHoursWeek).toString())
    .concat(" hours:</b><ul><li><b>")
    .concat(EIRCHoursWeek.toString())
    .concat(" hours</b> at the EIRC lab</li><li><b>")
    .concat(ITSHoursWeek.toString())
    .concat(" hours</b> at ITS </li></ul>")
  });

}

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
  if (beginPeriod.getFullYear()<1970) beginPeriod.setFullYear(beginPeriod.getFullYear() + 100);
  if (endPeriod.getFullYear()<1970) endPeriod.setFullYear(endPeriod.getFullYear() + 100);

  // FIND WEEKS SINCE FIRST PAY PERIOD
  var weeksSince = (endPeriod.getTime() - new Date("01/18/2018 00:00").getTime()) / MILLIS_PER_WEEK;

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

    // IF DATE IS ROUNDED DOWN TO 1900s, CORRECT BY ADDING 100 YEARS;
    if (entryDate.getFullYear()<1970) entryDate.setFullYear(entryDate.getFullYear() + 100);

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

function dateInRange(entryDate,beginPeriod,endPeriod) {

  var entryDateTime = entryDate.getTime();
  var beginPeriodTime = beginPeriod.getTime();
  var endPeriodTime = endPeriod.getTime();

  return (entryDateTime>=beginPeriodTime) & (entryDateTime<=endPeriodTime);

}

function compareFirstCol(a,b) {

  if (a[0]==b[0]) return 0;
  else return (a[0]>b[0]) ? 1 : -1;

}

function mod(n, p) {

    return n - p * Math.floor(n/p);

}

function addElements(a,b) {

  if (a.length != b.length) return -1;
  var sum = []
  for (i=0;i<a.length;i++) sum[i] = a[i] + b[i];
  return sum;

}

function savingsError(beginDate,endDate) {

  var beginDate = beginDate.getTimezoneOffset();
  var endDate = endDate.getTimezoneOffset();

  var timezoneDifference = (beginDate - endDate) / 60;

  return timezoneDifference

}