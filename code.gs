function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // create menu
  ui.createMenu('Strava')
    .addItem('Get Rides', 'getRides')
    .addItem('Initialize Ride Sheet', 'initializeRideSheet')
    .addToUi();
}

function getRides() {
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // all raw data will be written to sheet called "raw_data" - you must create this beforehand
  var sheet = spreadsheet.getSheetByName('raw_data');
  var data = initiateGetRides();
  var rideData = [];

  data.forEach(
    function(activity){
        var arr = [];
        // filter only on rides
        if(activity.type == "Ride"){
        arr.push(
          activity.start_date_local,
          activity.name,
          // convert to hours
          activity.moving_time/60,
          // convert to miles
          activity.distance*0.000621371,
          // convert to feet
          activity.total_elevation_gain*3.28084,
          // convert to mph
          activity.average_speed*0.000621371*60*60,
          // convert to mph
          activity.max_speed*0.000621371*60*60,
          activity.kilojoules,
          activity.average_heartrate,
          activity.max_heartrate,
          activity.gear_id,
          activity.athlete_count
        );
        rideData.push(arr);
      }
    }
  );
  if(rideData.length > 0){
    sheet.getRange(sheet.getLastRow() + 1, 1, rideData.length, rideData[0].length).setValues(rideData);
  }
}

function initializeRideSheet() {

  var rowHeaders = [
      ["Time", "Name", "Duration", "Distance", "Elevation", "Average Speed", "Max Speed", "Output (kj)", "Average HR", "Max HR", "Gear ID", "Athlete Count"]
  ]

  SpreadsheetApp.getActive().getRange("raw_data!A1:L1").setValues(rowHeaders);
  
}

function initiateGetRides() {
  
  var service = getStravaService();
  if (service.hasAccess()) {
    Logger.log('App has access.');
    var api = 'https://www.strava.com/api/v3/athlete/activities';
    // query parameter for input is to read activities after user provided epoch time
    // max # of activities you can fetch is 200
    // on initial data export, you'll need to read activities in increments of 200 and update the epoch time for last imported activity
    // best site to convert unix timestamp to epoch - http://www.unixtimestampconverter.com/
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('scratch_data');
    var epoch = sheet.getRange('H3').getValue() + 60;
    var query = '?after=' + epoch + '&per_page=200';
    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
    var response = JSON.parse(UrlFetchApp.fetch(api + query, options));
    return response;  
  }
  else {
    Logger.log("App has no access yet.");
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log("Open the following URL and re-run the script: %s",
        authorizationUrl);
  }
