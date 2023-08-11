// Instructions before running for the first time:
// 1. Be sure you have already set your personal access token in setToken.gs.
// 2. Add the spreadsheet ID of a new/empty to `spreadsheetID` below (i.e., if the spreadsheet URL is https://docs.google.com/spreadsheets/d/abc1234567/edit#gid=0, the spreadsheet ID is abc1234567).
// 3. Add your project ID to `projectID` below (i.e., if the project URL is https://app.asana.com/0/12345678/list, the project ID is 12345678).
// 4. Run the script.

const spreadsheetID = "";
const projectID = "";

// Get the personal access token initially set in setToken.gs
const userProperties = PropertiesService.getUserProperties();
const pat = userProperties.getProperty('pat');

// ================================================
// Main function (i.e., entry point)
// ================================================

async function runAutomation() {
  // Check for PAT
  if (!pat) {
    Logger.log("Please set your PAT in setToken.gs before continuing.");
    return
  }

  // Check for non-empty inputs
  if (spreadsheetID === "" || projectID === "") {
    Logger.log("Please provide both a spreadsheet ID and project ID to continue.");
    return;
  }

  // Perform basic validation of project ID
  if (!projectID || isNaN(projectID) || projectID.length < 3) {
    Logger.log("Your project GID is invalid. Please check it and try again");
    return;
  }

  // Get the user-provided spreadsheet
  const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  Logger.log(`Accessing spreadsheet "${spreadsheet.getName()}"...`);

  // Get the first sheet of the user-provided spreadsheet
  const sheet = spreadsheet.getSheets()[0];

  // Get the last (i.e., most bottom) row of the sheet
  const lastRow = sheet.getLastRow();

  // Write headers (i.e., column names) if the sheet is empty 
  if (lastRow === 0.0) {
    Logger.log("Empty sheet detected...");
    Logger.log("Writing sheet headers...");
    // If the sheet is empty, write sheet headers (i.e., column names)
    const headers = [
      "Status GID",
      "Link", 
      "Title",
      "Status type", 
      "Text",
      "Author", 
      "Created at", 
      ];
    sheet.appendRow(headers);
  }

  // If the sheet has project data, there will be at least two non-empty rows (i.e., the first row contains sheet headers)
  const hasProjectData = lastRow > 1;

  // If there is existing project data in the sheet...
  if (hasProjectData) {
    const latestStatusGID = sheet.getRange(lastRow, 1).getValue();
    const projectData = await getAsanaProject(projectID, options);
    if (!projectData) return;

    const { current_status } = projectData.data;

    // If the project's latest status update isn't in the sheet's last row, append that status update to the sheet
    if (latestStatusGID != current_status.gid) {
      const newRow = [
        current_status.gid,
        `https://app.asana.com/0/0/${current_status.gid}`,
        current_status.title,
        current_status.status_type,
        current_status.text,
        current_status.author.name,
        current_status.created_at,
      ];

      sheet.appendRow(newRow);

      Logger.log(`Latest project status successfully written for project ID: ${projectID}.`);
    } else {
      Logger.log("There are no new project statuses to write.");
    }
  } else { // If there is NO existing project data in the sheet...

    // Rename the sheet
    spreadsheet.rename(`Statuses for project ${projectID}`);

    // Get all existing status updates (past and current) from the project
    const response = await getAsanaStatuses(projectID, options);
    if (!response) return;

    // Write the statuses to the sheet, row by row
    const statusList = response.data;
    for (let i = statusList.length - 1; i >= 0; i--) {
      const newRow = [
        statusList[i].gid,
        `https://app.asana.com/0/0/${statusList[i].gid}`,
        statusList[i].title,
        statusList[i].status_type,
        statusList[i].text,
        statusList[i].created_by.name,
        statusList[i].created_at
      ];
      sheet.appendRow(newRow);
    }
    
    Logger.log(`Project statuses successfully written for project ID: ${projectID}.`);
  }

  Logger.log(`View the current spreadsheet here: https://docs.google.com/spreadsheets/d/${spreadsheetID}/edit#gid=0`);
}

// ================================================
// Helper functions
// ================================================

// Get a single project
// Documentation: https://developers.asana.com/reference/getproject
const getAsanaProject = (projectID, options) => {
  const url = `https://app.asana.com/api/1.0/projects/${projectID}?opt_fields=name,current_status.status_type,current_status.title,current_status.text,current_status.author.name,current_status.created_at,created_at,modified_at`;
  return getAsanaData(url, options);
};

// Get status updates from an object (i.e., project)
// Documentation: https://developers.asana.com/reference/getstatusesforobject
const getAsanaStatuses = (projectID, options) => {
  const url = `https://app.asana.com/api/1.0/status_updates?parent=${projectID}&opt_fields=title,status_type,text,created_by.name,created_at`;
  return getAsanaData(url, options);
};

// Make requests against Asana's API
const getAsanaData = async (url, options) => {
  try {
    const response = UrlFetchApp.fetch(url, options);
    const body = JSON.parse(response.getContentText());
    return body;
  } catch (error) {
    Logger.log("Error: Unable to fetch data. Please verify your API key and try again.");
    return null;
  }
};

// Create HTTP headers
const options = {
  'method': 'get',
  'headers': { 'Authorization': `Bearer ${pat}` },
  'muteHttpExceptions': true
};
