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
    .addItem('Get New Rides', 'runRideProcessing')
    .addItem('Get All Rides', 'runAllRideProcessing')
    .addItem('Get Segment Efforts', 'runSegmentEffortProcessing')
    .addToUi();
}

// Entry for processing and writing  rides data.
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('epoch');
  // Add 5 minutes to prevent fetching the same last activity again.
  return sheet.getRange('A1').getValue() + 300;
}


const stravaApi = {
  // Fetches all activities after the last recorded one.
  fetchAthleteActivities() {
    const lastRideEpoch = getLastRideEpoch();
    const endpoint = `/athlete/activities?after=${lastRideEpoch}&per_page=200`;
    return this._fetch(endpoint);
  },
  
  // New function to fetch all activities using pagination.
  fetchAllAthleteActivities() {
    let allActivities = [];
    let page = 1;
    let moreData = true;

    while (moreData) {
      const endpoint = `/athlete/activities?page=${page}&per_page=200`;
      Logger.log(`Fetching page ${page} of activities...`);
      const activities = this._fetch(endpoint);

      if (activities && activities.length > 0) {
        allActivities = allActivities.concat(activities);
        page++;
      } else {
        moreData = false;
      }
    }
    Logger.log(`Finished fetching. Total activities found: ${allActivities.length}`);
    return allActivities;
  },

  // Fetches all efforts for a given segment.
  // TODO: Support iteration over pagination if fetching more than 200.
  fetchSegmentEfforts(segmentId) {
    const endpoint = `/segments/${segmentId}/all_efforts?per_page=200`;
    const efforts = this._fetch(endpoint);
    // Attach segmentId to the response for later use in the formatter.
    return efforts ? { efforts, segmentId } : null;
  },
  
  // Generic utility for using the Strava API.
  _fetch(endpoint) {
    const service = getStravaService();
    const options = {
      headers: { Authorization: `Bearer ${service.getAccessToken()}` },
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


const ui = {
  // Prompt users for a segment ID when getting segment efforts.
  promptForSegmentId() {
    const ui = SpreadsheetApp.getUi();
    const result = ui.prompt('Enter a Strava segment ID', 'ID:', ui.ButtonSet.OK_CANCEL);
    const button = result.getSelectedButton();
    const segmentId = result.getResponseText();

    return (button === ui.Button.OK && segmentId) ? segmentId : null;
  },

  // Log authorization URL.
  logAuthorizationUrl(url) {
    Logger.log(`App has no access. Please authorize here: ${url}`);
  }
};
