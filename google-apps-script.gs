const SHEET_URL = "https://docs.google.com/spreadsheets/d/1qdU2CTcuk4peAIDGCNGTJYKfT-rl_9wcOAJJttrbWyE/edit?gid=0#gid=0";
const SHEET_NAME = "Bookings";

function getSheet_() {
  const ss = SpreadsheetApp.openByUrl(SHEET_URL);

  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    sheet.appendRow([
      "Timestamp",
      "Date",
      "Station",
      "StationName",
      "StartHour",
      "Duration",
      "StartTime",
      "EndTime",
      "Name",
      "Phone",
      "People"
    ]);
  }

  return sheet;
}


function doGet(e) {
  try {
    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();

    if (values.length === 0) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const headers = values.shift();

    const dateFilter = e.parameter.date;
    const stationFilter = e.parameter.station;

    const idx = {};

    headers.forEach((h, i) => {
      idx[h] = i;
    });

    const bookings = values
      .map(row => ({
        Timestamp: row[idx.Timestamp],
        Date: formatDate_(row[idx.Date]),
        Station: row[idx.Station],
        StationName: row[idx.StationName],
        StartHour: row[idx.StartHour],
        Duration: row[idx.Duration],
        StartTime: row[idx.StartTime],
        EndTime: row[idx.EndTime],
        Name: row[idx.Name],
        Phone: row[idx.Phone],
        People: row[idx.People]
      }))
      .filter(b => !dateFilter || b.Date === dateFilter)
      .filter(b => !stationFilter || b.Station === stationFilter);

    return ContentService
      .createTextOutput(JSON.stringify(bookings))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {

    return ContentService
      .createTextOutput(JSON.stringify({
        result: "error",
        message: err.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function doPost(e) {
  try {

    const data = JSON.parse(e.postData.contents);

    const sheet = getSheet_();

    sheet.appendRow([
      new Date(),
      data.date,
      data.station,
      data.stationName,
      data.startHour,
      data.duration,
      data.startTime,
      data.endTime,
      data.name,
      data.phone,
      data.people
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({
        result: "success"
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {

    return ContentService
      .createTextOutput(JSON.stringify({
        result: "error",
        message: err.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function formatDate_(value) {

  if (Object.prototype.toString.call(value) === "[object Date]") {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

  }

  return String(value);
}