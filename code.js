const api = 'https://www.strava.com/api/v3';

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Strava')
    .addItem('Get Rides', 'getRides')
    .addItem('Get Segment Efforts', 'getSegmentEfforts')
    .addToUi();
}

function getRides() {
  
  var context = 'rides'
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('raw_data');
  var data = callStrava(context);
  var rideData = [];

  data.forEach(
    function(activity){
        var arr = [];
        // TODO: Implement handling multiple possible activity types
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
  
  var context = 'segmentEfforts'
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('segment_effort_data');
  var data = callStrava(context);
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

function callStrava(context) {
  
  var service = getStravaService();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var headers = {
    Authorization: 'Bearer ' + service.getAccessToken()
  };
  var options = {
    headers: headers,
    method : 'GET',
    muteHttpExceptions: true
  };
  if (service.hasAccess()) {
    Logger.log('App has access.');
    if (context == 'rides') {
      var path = '/athlete/activities';
      var lastRideEpoch = spreadsheet.getSheetByName('epoch').getRange('A1').getValue() + 60;
      var query = '?after=' + lastRideEpoch + '&per_page=200';
      
    } 
    if (context == 'segmentEfforts') {
      var path = '/segment_efforts';
      // TODO: Implement handling a dynamicly populated segment ID. 
      var segmentID = '141491'
      var query = '?segment_id=' + segmentID + '&per_page=200';
    }
    var response = JSON.parse(UrlFetchApp.fetch(api + path + query, options));
    return response;  
  }
  else {
    Logger.log("App has no access yet.");
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log("Open the following URL and re-run the script: %s",
        authorizationUrl);
  }
}
