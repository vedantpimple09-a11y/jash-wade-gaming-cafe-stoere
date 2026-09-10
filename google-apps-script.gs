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


/**
 * GET
 * Returns bookings for a selected date and station.
 */
function doGet(e) {
  try {
    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) {
      return jsonOutput_([]);
    }

    const headers = values.shift();

    const dateFilter = e.parameter.date || "";
    const stationFilter = e.parameter.station || "";

    const idx = {};

    headers.forEach((header, i) => {
      idx[header] = i;
    });

    const bookings = values
      .filter(row => row.length > 0)
      .map(row => ({
        Timestamp: row[idx.Timestamp],
        Date: formatDate_(row[idx.Date]),
        Station: row[idx.Station],
        StationName: row[idx.StationName],
        StartHour: Number(row[idx.StartHour]),
        Duration: Number(row[idx.Duration]),
        StartTime: row[idx.StartTime],
        EndTime: row[idx.EndTime],
        Name: row[idx.Name],
        Phone: row[idx.Phone],
        People: row[idx.People]
      }))
      .filter(booking =>
        !dateFilter || booking.Date === dateFilter
      )
      .filter(booking =>
        !stationFilter || String(booking.Station) === String(stationFilter)
      );

    return jsonOutput_(bookings);

  } catch (err) {

    return jsonOutput_({
      result: "error",
      message: err.message
    });
  }
}


/**
 * POST
 * Creates a new booking.
 *
 * Uses LockService so simultaneous bookings
 * cannot book the same station/time together.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const data = JSON.parse(e.postData.contents);

    const requiredFields = [
      "date",
      "station",
      "stationName",
      "startHour",
      "duration",
      "startTime",
      "endTime",
      "name",
      "phone",
      "people"
    ];

    for (const field of requiredFields) {
      if (
        data[field] === undefined ||
        data[field] === null ||
        String(data[field]).trim() === ""
      ) {
        return jsonOutput_({
          result: "error",
          message: `Missing field: ${field}`
        });
      }
    }

    const date = String(data.date).trim();
    const station = String(data.station).trim();

    const startHour = Number(data.startHour);
    const duration = Number(data.duration);

    if (
      !Number.isInteger(startHour) ||
      !Number.isInteger(duration) ||
      duration < 1 ||
      startHour < 10 ||
      startHour + duration > 23
    ) {
      return jsonOutput_({
        result: "error",
        message: "Invalid booking time or duration."
      });
    }

    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();

    /*
     * Check existing bookings for overlap.
     *
     * Example:
     * Existing: 5 PM - 7 PM
     * New:      6 PM - 8 PM
     *
     * This will be rejected.
     */
    if (values.length > 1) {
      const headers = values[0];

      const idx = {};

      headers.forEach((header, i) => {
        idx[header] = i;
      });

      for (let i = 1; i < values.length; i++) {
        const row = values[i];

        const existingDate = formatDate_(row[idx.Date]);
        const existingStation = String(row[idx.Station]);

        if (
          existingDate === date &&
          existingStation === station
        ) {
          const existingStart = Number(row[idx.StartHour]);
          const existingDuration = Number(row[idx.Duration]);
          const existingEnd =
            existingStart + existingDuration;

          const newEnd =
            startHour + duration;

          const overlaps =
            startHour < existingEnd &&
            newEnd > existingStart;

          if (overlaps) {
            return jsonOutput_({
              result: "error",
              message:
                "This time slot has just been booked. Please choose another time."
            });
          }
        }
      }
    }

    /*
     * Add booking.
     */
    sheet.appendRow([
      new Date(),
      date,
      station,
      String(data.stationName),
      startHour,
      duration,
      String(data.startTime),
      String(data.endTime),
      String(data.name).trim(),
      String(data.phone).trim(),
      String(data.people)
    ]);

    SpreadsheetApp.flush();

    return jsonOutput_({
      result: "success",
      message: "Booking confirmed."
    });

  } catch (err) {

    return jsonOutput_({
      result: "error",
      message: err.message
    });

  } finally {

    try {
      lock.releaseLock();
    } catch (e) {
      // Ignore lock release errors.
    }
  }
}


/**
 * Converts response to JSON.
 */
function jsonOutput_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Converts Sheet date values into yyyy-MM-dd.
 */
function formatDate_(value) {

  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value.getTime())
  ) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
  }

  return String(value).trim();
}