function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Strava')
    .addItem('Get Rides', 'getRides')
    .addItem('Get Segment Efforts', 'getSegmentEfforts')
    .addToUi();
}

function initiateStrava() {
  
  var service = getStravaService();
  if (service.hasAccess()) {
    Logger.log('App has access.');
    if (context == 'rides') {
      var api = 'https://www.strava.com/api/v3/athlete/activities';
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = spreadsheet.getSheetByName('epoch');
      var epoch = sheet.getRange('A1').getValue() + 60;
      var query = '?after=' + epoch + '&per_page=200';
      var headers = {
        Authorization: 'Bearer ' + service.getAccessToken()
      };
      var options = {
        headers: headers,
        method : 'GET',
        muteHttpExceptions: true
      };
    } 
    if (context == 'segments') {
      var segmentID = '141491'
      var api = 'https://www.strava.com/api/v3/segment_efforts';
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var query = '?segment_id=' + segmentID + '&per_page=200';
      var headers = {
        Authorization: 'Bearer ' + service.getAccessToken()
      };
      var options = {
        headers: headers,
        method : 'GET',
        muteHttpExceptions: true
      };
    }
    var response = JSON.parse(UrlFetchApp.fetch(api + query, options));
    return response;  
  }
  else {
    Logger.log("App has no access yet.");
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log("Open the following URL and re-run the script: %s",
        authorizationUrl);
  }
}

function getRides() {
  
  var context = 'rides'
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('raw_data');
  var data = initiateStrava(context);
  var rideData = [];

  data.forEach(
    function(activity){
        var arr = [];
        if(activity.type == "Ride"){
        arr.push(
          activity.start_date_local,
          activity.name,
          activity.moving_time/60,
          activity.distance*0.000621371,
          activity.total_elevation_gain*3.28084,
          activity.average_speed*0.000621371*60*60,
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

function getSegmentEfforts() {
  
  var context = 'segments'
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('segment_effort_data');
  var data = initiateStrava(context);
  var segmentEffortData = [];

  data.forEach(
    function(segmentEffort){
        var arr = [];
        arr.push(
          segmentEffort.start_date_local,
          segmentEffort.elapsed_time
        );
        segmentEffortData.push(arr);
      }
  );
  if(segmentEffortData.length > 0){
    sheet.getRange(sheet.getLastRow() + 1, 1, segmentEffortData.length, segmentEffortData[0].length).setValues(segmentEffortData);
  }
}
