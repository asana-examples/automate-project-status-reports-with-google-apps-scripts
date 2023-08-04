// If this is your first time using the application, you must run this script (setToken.gs) once before running anything else
// Instructions:
// 1. Set your the `pat` to your personal access token (documentation: https://developers.asana.com/docs/personal-access-token).
// 2. Run the script once via the "Run" button above.
// 3. Remove your token and save the file. Do not run the script again.

function setToken() {
  const pat = "";

  // Check if the API key is provided
  if (pat === "") {
    Logger.log("Error: Please set your personal access token (pat) before running the script.");
    return;
  }

  // Perform a basic profile check and log the result for validation that the api key works.
  let url = "https://app.asana.com/api/1.0/users/me";
  let options = {
    'method': 'get',
    'headers': { 'Authorization': `Bearer ${pat}` },
    'muteHttpExceptions': true
  };

  try {
    let response = UrlFetchApp.fetch(url, options);
    let body = JSON.parse(response.getContentText());
    Logger.log(body);
  } catch (error) {
    Logger.log("Error: Unable to fetch data. Please verify your API key and try again.");
    return;
  }

  // Set the API key in user properties
  let userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('pat', pat);

  Logger.log("API key has been successfully set.");
}
