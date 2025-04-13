/***************** COMMON HELPER FUNCTIONS ********************/
// Retrieve the API key securely from the script's properties.
const ASANA_TOKEN = PropertiesService.getScriptProperties().getProperty('ASANA_API_KEY');

/**
 * Fetches the direct subtasks for a given task.
 * If the taskId is invalid or an error occurs, returns an empty array.
 * @param {string} taskId - The Asana task GID.
 * @return {Array} - An array of subtasks.
 */
function fetchSubtasks(taskId) {
  if (!taskId) {
    Logger.log("fetchSubtasks: taskId is undefined, returning empty array.");
    return [];
  }
  
  const headers = {
    "Authorization": "Bearer " + ASANA_TOKEN,
    "Content-Type": "application/json"
  };
  // URL to fetch direct subtasks with necessary fields.
  const subtasksUrl = `https://app.asana.com/api/1.0/tasks/${taskId}/subtasks?opt_fields=completed,completed_at&limit=100`;
  let subtasks = [];
  let nextPage = subtasksUrl;
  
  while (nextPage) {
    try {
      const response = UrlFetchApp.fetch(nextPage, { headers: headers });
      const result = JSON.parse(response.getContentText());
      Logger.log("Fetched subtasks for task " + taskId + ", received " + result.data.length + " items.");
      subtasks = subtasks.concat(result.data);
      nextPage = (result.next_page && result.next_page.uri) ? result.next_page.uri : null;
    } catch (e) {
      Logger.log("Error fetching subtasks for taskId " + taskId + ": " + e);
      nextPage = null; // Exit the loop on error.
    }
  }
  return subtasks;
}

/**
 * Recursively fetches all subtasks for a given task (both direct and nested).
 * @param {string} taskId - The Asana task GID.
 * @return {Array} - An array containing all subtasks at any depth.
 */
function fetchAllSubtasks(taskId) {
  if (!taskId) return [];
  
  let directSubtasks = fetchSubtasks(taskId);
  let allSubtasks = [];
  
  directSubtasks.forEach(function(subtask) {
    if (subtask.gid) {
      allSubtasks.push(subtask);
      let nested = fetchAllSubtasks(subtask.gid);
      allSubtasks = allSubtasks.concat(nested);
    } else {
      Logger.log("Subtask without gid encountered; skipping nested fetch.");
    }
  });
  return allSubtasks;
}

/***************** PROJECT-SPECIFIC FUNCTION TEMPLATE ********************/
/**
 * updateAsanaProject1
 *
 * This function fetches all tasks (top-level plus all nested subtasks) for a specific Asana project,
 * then calculates:
 *   - Total tasks,
 *   - Completed tasks.
 *
 * It writes these values into designated fixed columns on a specified row of a Google Sheet:
 *   - Column A: Timestamp of last update.
 *   - Column C: Total tasks.
 *   - Column D: Completed tasks.
 *
 * HOW TO USE:
 * 1. Verify that this function works for one Asana project.
 * 2. To add a new project:
 *    - Copy the entire updateAsanaProject1() function (from "function updateAsanaProject1() {" to its final "}").
 *    - Paste it below the original function.
 *    - Rename the new function (for example, updateAsanaProject2).
 *    - In the new function, update the following project-specific constants:
 *         • PROJECT_ID: The new Asana Project ID.
 *         • SHEET_NAME: The name of the sheet where data should be written.
 *         • ROW: The row number for the new project (e.g., 2 writes to A2, C2, D2).
 *    - Save your script.
 *    - Open the Triggers panel (click the clock icon in the left sidebar) and add a new time-driven trigger for the new function.
 */
function updateAsanaProject1() {
  // ----------------- PROJECT-SPECIFIC CONSTANTS -----------------
  const PROJECT_ID = "1209195244873699";   // Asana Project ID for Project 1 (BeamlineX)
  const SHEET_NAME = "Sheet1";             // Name of the Google Sheet tab
  const ROW = 2;                         // The row number for this project (e.g., 2 writes to A2, C2, D2)
  // ----------------- END OF PROJECT-SPECIFIC SETTINGS -----------------
  
  const headers = {
    "Authorization": "Bearer " + ASANA_TOKEN,
    "Content-Type": "application/json"
  };
  
  // Use a very early date to ensure both completed and incomplete tasks are returned.
  // (Adjust or remove the completed_since parameter if needed.)
  const baseUrl = `https://app.asana.com/api/1.0/projects/${PROJECT_ID}/tasks?completed_since=1970-01-01T00:00:00Z&opt_fields=completed,completed_at&limit=100`;
  
  let topLevelTasks = [];
  let nextPage = baseUrl;
  
  // Fetch top-level tasks with pagination.
  while (nextPage) {
    try {
      const response = UrlFetchApp.fetch(nextPage, { headers: headers });
      const result = JSON.parse(response.getContentText());
      Logger.log("Fetched " + result.data.length + " top-level tasks from: " + nextPage);
      topLevelTasks = topLevelTasks.concat(result.data);
      nextPage = (result.next_page && result.next_page.uri) ? result.next_page.uri : null;
    } catch (e) {
      Logger.log("Error fetching top-level tasks: " + e);
      nextPage = null;
    }
  }
  
  Logger.log("Total top-level tasks fetched: " + topLevelTasks.length);
  
  // For every top-level task, fetch its nested subtasks recursively.
  let allSubtasks = [];
  topLevelTasks.forEach(function(task) {
    if (task.gid) {
      let subtasks = fetchAllSubtasks(task.gid);
      allSubtasks = allSubtasks.concat(subtasks);
    }
  });
  
  // Combine top-level tasks and all subtasks.
  let combinedTasks = topLevelTasks.concat(allSubtasks);
  
  // Calculate counts.
  let totalTasks = combinedTasks.length;
  let completedTasks = 0;
  
  combinedTasks.forEach(function(task) {
    if (task.completed) {
      completedTasks++;
    }
  });
  
  Logger.log("Project ID: " + PROJECT_ID);
  Logger.log("Total tasks (top-level + subtasks): " + totalTasks);
  Logger.log("Completed tasks: " + completedTasks);
  
  //Get the current timezone and format it to be the user's timeset, manually adding 1 for daylight savings
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var now = new Date();
  now.setHours(now.getHours() + 1);  // Adds one hour
  const timestamp = Utilities.formatDate(now, tz, "MM/dd/yyyy HH:mm:ss");


  // Write results to the designated cells:
  // - Column A: Timestamp of last update.
  // - Column C: Total tasks.
  // - Column D: Completed tasks.
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.getRange("A" + ROW).setValue(timestamp);            // Timestamp in Column A
  sheet.getRange("C" + ROW).setValue(totalTasks);              // Total tasks in Column C
  sheet.getRange("D" + ROW).setValue(completedTasks);          // Completed tasks in Column D
}

/***************** COPY-AND-PASTE INSTRUCTIONS ********************/
/*
To create a function for a new Asana project (without tasks-per-person details):

1. Copy the entire updateAsanaProject1() function (from "function updateAsanaProject1() {" to its final "}" at the end).
2. Paste it below the original function.
3. Rename the new function (e.g., updateAsanaProject2()).
4. Within the new function, update these project-specific constants:
   - Change PROJECT_ID to the new project's Asana Project ID.
   - Change SHEET_NAME if you want the output on a different sheet.
   - Change ROW to the row number you wish the new project’s data to occupy. 
     (For example, if ROW is 3, then:
        • Timestamp will be written to A3,
        • Total tasks to C3,
        • Completed tasks to D3.)
5. Save your script.
6. Open the Triggers panel (click the clock icon in the left sidebar).
7. Click "+ Add Trigger" and select your new function (e.g., updateAsanaProject2) to run at the desired time.
*/
