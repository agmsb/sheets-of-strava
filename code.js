const API_BASE_URL = 'https://www.strava.com/api/v3';
const METERS_TO_MILES = 0.000621371;
const METERS_TO_FEET = 3.28084;
const M_PER_S_TO_MPH = 2.23694;

// Config to dictate behavior for rides and segment efforts.
const CONFIG = {
  rides: {
    sheetName: 'raw_data',
    fetcher: () => stravaApi.fetchAthleteActivities(),
    formatter: (activities) => activities
      .filter(act => act.type === "Ride")
      .map(formatRideData),
  },
  allRides: {
    sheetName: 'raw_data',
    fetcher: () => stravaApi.fetchAllAthleteActivities(),
    formatter: (activities) => activities
      .filter(act => act.type === "Ride")
      .map(formatRideData),
  },
  segmentEfforts: {
    sheetName: 'segment_effort_data',
    fetcher: () => {
      const segmentId = ui.promptForSegmentId();
      return segmentId ? stravaApi.fetchSegmentEfforts(segmentId) : null;
    },
    formatter: (response) => {
      if (!response || !response.efforts) return [];
      return formatSegmentEffortData(response.efforts, response.segmentId);
    },
  },
};

// Creates menu for users in Google Sheets.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Strava')
    .addItem('Setup Rides Sheet', 'setupRidesSheet')
    .addItem('Get All Rides', 'runAllRideProcessing')
    .addSeparator()
    .addItem('Get New Rides', 'runRideProcessing')
    .addSeparator()
    .addItem('Get Segment Efforts', 'runSegmentEffortProcessing')
    .addToUi();
}

// Create raw_data sheet if it doesn't exist.
function setupRidesSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'raw_data';

  if (spreadsheet.getSheetByName(sheetName)) {
    ui.alert(`The sheet "${sheetName}" already exists.`);
    return;
  }

  const headers = [[
    'Date', 'Name', 'Time (min)', 'Distance (mi)', 'Elevation (ft)', 'Avg Speed (mph)',
    'Max Speed (mph)', 'Kilojoules', 'Avg Heart Rate', 'Max Heart Rate',
    'Gear ID', 'Athlete Count'
  ]];
  
  const sheet = spreadsheet.insertSheet(sheetName);
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  sheet.setFrozenRows(1); // Freeze the header row for better usability

  ui.alert(`Successfully created the "${sheetName}" sheet with headers.`);
}

// Entry for processing and writing rides data.
function runRideProcessing() {
  processData(CONFIG.rides);
}

// Eentry for processing and writing all rides data.
function runAllRideProcessing() {
  processData(CONFIG.allRides);
}

// Entry for processing and writing segment efforts data.
function runSegmentEffortProcessing() {
  processData(CONFIG.segmentEfforts);
}

// Fetch, format, and write data to a sheet.
function processData(config) {
  const service = getStravaService();
  if (!service.hasAccess()) {
    return ui.logAuthorizationUrl(service.getAuthorizationUrl());
  }

  const rawData = config.fetcher();
  if (!rawData || (Array.isArray(rawData) && rawData.length === 0)) {
    Logger.log("No new data to process.");
    return;
  }

  const formattedData = config.formatter(rawData);
  if (formattedData.length > 0) {
    writeDataToSheet(config.sheetName, formattedData);
  } else {
    Logger.log("No data left after formatting.");
  }
}

// Converts rides data default metrics into desired format.
function formatRideData(activity) {
  return [
    activity.start_date_local,
    activity.name,
    activity.moving_time / 60, // to minutes
    activity.distance * METERS_TO_MILES,
    activity.total_elevation_gain * METERS_TO_FEET,
    activity.average_speed * M_PER_S_TO_MPH,
    activity.max_speed * M_PER_S_TO_MPH,
    activity.kilojoules,
    activity.average_heartrate,
    activity.max_heartrate,
    activity.gear_id,
    activity.athlete_count,
  ];
}

// Formats segment efforts data.
function formatSegmentEffortData(efforts, segmentId) {
  return efforts.map(effort => [
    segmentId,
    effort.start_date_local,
    effort.elapsed_time,
  ]);
}

// Writes data to sheet.
function writeDataToSheet(sheetName, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
}

// Fetches the epoch time from your last ride.
function getLastRideEpoch() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('raw_data');
  
  // If sheet doesn't exist or has only a header, fetch all activities.
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log('No activities found. Defaulting to fetch all activities.');
    return 0;
  }

  const lastRow = sheet.getLastRow();
  const lastDateString = sheet.getRange(lastRow, 1).getValue();
  const lastEpoch = convertToEpoch(lastDateString);

  // Add 25200s to account for Pacific time.
  return lastEpoch + 25200
;
}

// Converts a date string to epoch time.
function convertToEpoch(dateString) {
  if (!dateString) return 0;
  
  const milliseconds = Date.parse(dateString);

  // Return epoch in seconds, or 0 if parsing fails.
  return isNaN(milliseconds) ? 0 : milliseconds / 1000;
}

const stravaApi = {
  // Fetches any activities after the last recorded activity.
  fetchAthleteActivities() {
    const lastRideEpoch = getLastRideEpoch();
    const endpoint = `/athlete/activities?after=${lastRideEpoch}&per_page=200`;
    return this._fetch(endpoint);
  },

  // Fetches all activities from the user's Strava history using the paginated fetcher.
  fetchAllAthleteActivities() {
    const baseEndpoint = '/athlete/activities';
    const allActivities = this._fetchAllPages(baseEndpoint);
    Logger.log(`Finished fetching. Total activities found: ${allActivities.length}`);
    return allActivities;
  },

  // Fetches all efforts for a given segment using the paginated fetcher.
  fetchSegmentEfforts(segmentId) {
    const baseEndpoint = `/segments/${segmentId}/all_efforts`;
    const allEfforts = this._fetchAllPages(baseEndpoint);
    Logger.log(`Finished fetching. Total efforts found for segment ${segmentId}: ${allEfforts.length}`);

    // Attach segmentId to the response for later use in the formatter.
    return allEfforts.length > 0 ? {
      efforts: allEfforts,
      segmentId
    } : null;
  },

  // Utility for fetching all pages for an endpoint from the Strava API.
  _fetchAllPages(baseEndpoint) {
    let allItems = [];
    let page = 1;
    let moreData = true;

    while (moreData) {
      const endpoint = `${baseEndpoint}?page=${page}&per_page=200`;
      Logger.log(`Fetching page ${page} for ${baseEndpoint}...`);
      const items = this._fetch(endpoint);

      if (items && items.length > 0) {
        allItems = allItems.concat(items);
        page++;
      } else {
        moreData = false;
      }
    }
    return allItems;
  },

  // Utility for fetching a page from the Strava API.
  _fetch(endpoint) {
    const service = getStravaService();
    const options = {
      headers: {
        Authorization: `Bearer ${service.getAccessToken()}`
      },
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(`${API_BASE_URL}${endpoint}`, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      return JSON.parse(response.getContentText());
    }

    Logger.log(`API request failed for ${endpoint} with status ${responseCode}: ${response.getContentText()}`);
    return null;
  }
};
