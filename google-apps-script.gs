const SHEET_URL = "https://docs.google.com/spreadsheets/d/1qdU2CTcuk4peAIDGCNGTJYKfT-rl_9wcOAJJttrbWyE/edit";
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


/* ================================
   GET BOOKINGS
================================ */

function doGet(e) {
  try {

    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) {
      return jsonOutput_([]);
    }

    const headers = values[0];

    const dateFilter =
      e && e.parameter
        ? String(e.parameter.date || "")
        : "";

    const stationFilter =
      e && e.parameter
        ? String(e.parameter.station || "")
        : "";

    const bookings = [];

    for (let i = 1; i < values.length; i++) {

      const row = values[i];

      const booking = {

        Timestamp: row[0],

        Date: formatDate_(row[1]),

        Station: String(row[2] || ""),

        StationName: String(row[3] || ""),

        StartHour: Number(row[4]),

        Duration: Number(row[5]),

        StartTime: String(row[6] || ""),

        EndTime: String(row[7] || ""),

        Name: String(row[8] || ""),

        Phone: String(row[9] || ""),

        People: String(row[10] || "")
      };


      if (
        dateFilter &&
        booking.Date !== dateFilter
      ) {
        continue;
      }


      if (
        stationFilter &&
        booking.Station !== stationFilter
      ) {
        continue;
      }


      bookings.push(booking);
    }


    return jsonOutput_(bookings);


  } catch (error) {

    return jsonOutput_({
      result: "error",
      message: error.message
    });
  }
}


/* ================================
   CREATE BOOKING
================================ */

function doPost(e) {

  const lock =
    LockService.getScriptLock();

  try {

    /*
     * Prevent two people from booking
     * the same slot simultaneously.
     */
    lock.waitLock(15000);


    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {

      return jsonOutput_({
        result: "error",
        message: "No booking data received."
      });
    }


    const data =
      JSON.parse(
        e.postData.contents
      );


    /*
     * Required fields
     */
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


    for (
      let i = 0;
      i < requiredFields.length;
      i++
    ) {

      const field =
        requiredFields[i];


      if (
        data[field] === undefined ||
        data[field] === null ||
        String(data[field]).trim() === ""
      ) {

        return jsonOutput_({
          result: "error",
          message:
            "Missing booking information: " +
            field
        });
      }
    }


    const date =
      String(data.date).trim();

    const station =
      String(data.station).trim();

    const stationName =
      String(data.stationName).trim();

    const startHour =
      Number(data.startHour);

    const duration =
      Number(data.duration);

    const startTime =
      String(data.startTime).trim();

    const endTime =
      String(data.endTime).trim();

    const name =
      String(data.name).trim();

    const phone =
      String(data.phone).trim();

    const people =
      String(data.people).trim();


    /*
     * Validate time.
     *
     * Opening: 10 AM
     * Closing: 11 PM
     */
    if (
      !Number.isInteger(startHour) ||
      !Number.isInteger(duration) ||
      startHour < 10 ||
      duration < 1 ||
      startHour + duration > 23
    ) {

      return jsonOutput_({
        result: "error",
        message:
          "Invalid booking time."
      });
    }


    /*
     * Validate name.
     */
    if (name.length < 2) {

      return jsonOutput_({
        result: "error",
        message:
          "Please enter a valid name."
      });
    }


    /*
     * Validate phone.
     */
    const cleanPhone =
      phone.replace(/\D/g, "");


    if (cleanPhone.length !== 10) {

      return jsonOutput_({
        result: "error",
        message:
          "Please enter a valid 10-digit phone number."
      });
    }


    const sheet =
      getSheet_();

    const values =
      sheet.getDataRange().getValues();


    /*
     * Check existing bookings.
     *
     * Prevents overlapping bookings
     * on the same date + station.
     */
    if (values.length > 1) {

      for (
        let i = 1;
        i < values.length;
        i++
      ) {

        const row =
          values[i];


        const existingDate =
          formatDate_(row[1]);

        const existingStation =
          String(row[2] || "");

        const existingStart =
          Number(row[4]);

        const existingDuration =
          Number(row[5]);

        const existingEnd =
          existingStart +
          existingDuration;


        /*
         * Only compare bookings
         * for the same date and station.
         */
        if (
          existingDate === date &&
          existingStation === station
        ) {

          const newStart =
            startHour;

          const newEnd =
            startHour + duration;


          /*
           * Overlap formula:
           *
           * New start < old end
           * AND
           * New end > old start
           */
          const overlap =
            newStart < existingEnd &&
            newEnd > existingStart;


          if (overlap) {

            return jsonOutput_({

              result: "error",

              message:
                "Sorry, this time slot has just been booked. Please select another available time."
            });
          }
        }
      }
    }


    /*
     * Save booking.
     */
    sheet.appendRow([

      new Date(),

      date,

      station,

      stationName,

      startHour,

      duration,

      startTime,

      endTime,

      name,

      cleanPhone,

      people

    ]);


    SpreadsheetApp.flush();


    return jsonOutput_({

      result: "success",

      message:
        "Booking confirmed successfully."

    });


  } catch (error) {

    return jsonOutput_({

      result: "error",

      message: error.message

    });

  } finally {

    try {

      lock.releaseLock();

    } catch (e) {

      // Nothing to do.
    }
  }
}


/* ================================
   JSON RESPONSE
================================ */

function jsonOutput_(data) {

  return ContentService

    .createTextOutput(
      JSON.stringify(data)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );
}


/* ================================
   DATE FORMATTER
================================ */

function formatDate_(value) {

  if (
    Object.prototype.toString.call(value)
    === "[object Date]"
  ) {

    if (
      !isNaN(value.getTime())
    ) {

      return Utilities.formatDate(

        value,

        Session.getScriptTimeZone(),

        "yyyy-MM-dd"
      );
    }
  }


  return String(value || "").trim();
}