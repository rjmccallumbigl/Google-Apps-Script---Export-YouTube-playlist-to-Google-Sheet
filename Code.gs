/****************************************************************************************************************************************
*
* Export YouTube playlist to Google Sheet
*
* Instructions
* 1. Open Google Apps Script for your spreadsheet (Tools -> Script Editor).
* 2. Delete all text in the scripting window and paste all this code. 
* 3. Replace the string "replaceThisText!" in the variable playlistId with your playlist ID. It is the identifier string in the URL after https://www.youtube.com/playlist?list=
* 4. Click "Services" in Google Apps Script and add the "YouTube Data API v3" service.
* 5. Run onOpen(). Accept the permissions.
* 6. Then run "Export YouTube playlist to Google Sheet" from the spreadsheet (Functions menu at end of menu bar) or retrieveMyYouTubePlaylistItems from the script.
*
* Sources
* https://www.reddit.com/r/googlesheets/comments/nbr56j/trying_to_turn_my_youtube_music_playlist_into_a/
* https://developers.google.com/apps-script/advanced/youtube
* https://developers.google.com/youtube/v3/docs/playlistItems/list
*
****************************************************************************************************************************************/

function retrieveMyYouTubePlaylistItems() {

  // Declare variables
  var nextPageToken;
  var ytPlaylistItems = [];
  var playlistId = "replaceThisText!";

  // Call to YouTube API
  do {
    var playlistResponse = YouTube.PlaylistItems.list('snippet', {
      playlistId: playlistId,
      maxResults: 50,
      pageToken: nextPageToken
    });

    // Loop through returned playlist items and add them to the array
    for (var j = 0; j < playlistResponse.items.length; j++) {
      ytPlaylistItems.push(playlistResponse.items[j].snippet);
      nextPageToken = playlistResponse.nextPageToken;
    }
  }
  while (nextPageToken);

  // Print to spreadsheet
  setArraySheet(ytPlaylistItems, playlistId);
}

/******************************************************************************************************
* Convert 2D array into sheet
* 
* @param {Array} array The multidimensional array that we need to map to a sheet
* @param {String} sheetName The name of the sheet the array is being mapped to
* 
******************************************************************************************************/

function setArraySheet(array, sheetName) {

  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var keyArray = [];
  var memberArray = [];

  // Define an array of all the returned object's keys to act as the Header Row
  keyArray.length = 0;
  keyArray = Object.keys(array[0]);
  memberArray.length = 0;
  memberArray.push(keyArray);

  //  Capture players from returned data
  for (var x = 0; x < array.length; x++) {
    memberArray.push(keyArray.map(function (key) { return array[x][key] }));
  }

  // Select the sheet and set values  
  try {
    sheet = spreadsheet.insertSheet(sheetName);
  } catch (e) {
    sheet = spreadsheet.getSheetByName(sheetName).clear();
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, memberArray.length, memberArray[0].length).setValues(memberArray);
}

/*********************************************************************************************************
*
* Create Google Sheet menu allowing script to be run from the spreadsheet.
*
*********************************************************************************************************/

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Functions')
    .addItem('Export YouTube playlist to Google Sheet', 'retrieveMyYouTubePlaylistItems')
    .addToUi();
}
