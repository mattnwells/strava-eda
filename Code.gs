// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
 
  ui.createMenu('Strava App')
    .addItem('Get data', 'getStravaActivityData')
    .addToUi();
}

// Get athlete activity data
function getStravaActivityData() {

// get the sheet
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('Sheet1');

// call the Strava API to retrieve data
var data = callStravaActivityAPI();

// empty array to hold activity data
var stravaData = [];

// loop over activity data and add to stravaData array for Sheet
Logger.log('Starting to write.');
data.forEach(function(activity) {
var arr = [];
arr.push(
activity.id,
activity.name,
activity.distance,
activity.moving_time,
activity.elapsed_time,
activity.total_elevation_gain,
activity.average_speed,
activity.max_speed,
activity.type,
activity.workout_type,
activity.start_date_local,
activity.timezone,
activity.location_city,
activity.location_state,
activity.location_country,
activity.achievement_count,
activity.average_temp,
activity.average_watts,
activity.weighted_average_watts,
activity.max_watts,
activity.average_heartrate,
activity.max_heartrate,
activity.elev_high,
activity.elev_low,
activity.pr_count,
activity.suffer_score,
activity.start_latitude,
activity.start_longitude,
);
stravaData.push(arr);
});

// paste the values into the Sheet
// The below will do a full data refresh
// Clear all data
sheet.getRange(2,1,sheet.getLastRow(), stravaData[0].length).clearContent();
Logger.log('Clear the sheet.');
// Populate new data
sheet.getRange(2,1,stravaData.length, stravaData[0].length).setValues(stravaData);

// Past the refresh time stamp into the sheet
// variables = ss.getSheetByName(‘Variables’);
// Get today’s date and time to display when the data was last refreshed
// const refresh_date = Utilities.formatDate(new Date(), “GMT”, “MM/dd/yyyy HH:mm”);
// variables.getRange(“B3”).setValue(refresh_date);
// variables.getRange(“E5”).setValue(“Full”);
}

// call the Strava Activity API
function callStravaActivityAPI() {

// set up the service
var service = getStravaService();

if (service.hasAccess()) {
Logger.log('App has access.');

var endpoint = 'https://www.strava.com/api/v3/athlete/activities';
var params = '?after=1682899201&per_page=100';
var page = 1;

var headers = {
Authorization: 'Bearer ' + service.getAccessToken()
};

var options = {
headers: headers,
method : 'GET',
muteHttpExceptions: true
};

// Initiate the response array
var response = [];

// Get current response (first page of data)
var current_response = JSON.parse(UrlFetchApp.fetch(endpoint + params + page, options));

Logger.log('Starting the loop of responses.');
while (current_response.length != 0) {
page ++;
response = response.concat(current_response);
current_response = JSON.parse(UrlFetchApp.fetch(endpoint + params + page, options));
}
Logger.log('Array built.');

return response;
}
else {
Logger.log("App has no access yet.");

// open this url to gain authorization from github
var authorizationUrl = service.getAuthorizationUrl();

Logger.log("Open the following URL and re-run the script: %s",
authorizationUrl);
}
}
