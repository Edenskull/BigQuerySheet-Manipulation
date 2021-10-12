var ui = SpreadsheetApp.getUi();
var PROMPTTITLE = "[BigQuerySheets System]";
var ALERTMESSAGECANCEL = "You canceled the process. You need to provide all the requested information for the script to run correctly.";

/**
 * Function to trigger the menu creation when the spreadsheet is open by the user
 */
function onOpen() {
  ui.createMenu('[🔄] BigQuerySheets-Manipulation')
    .addItem('Append Data', 'appendData')
    .addItem('Update Data', 'updateData')
    .addSeparator()
    .addItem('Update Configuration', 'updateConfiguration')
    .addItem('Show Configuration', 'showConfiguration')
    .addToUi();
}

/**
 * Function to update the current document (Spreadsheet) configuration. With this all the collaborator using a same Spreadsheet can execute event with the same configuration.
 */
function updateConfiguration() {
  let pidResp = ui.prompt(PROMPTTITLE, "Please provide the GCP project id:", ui.ButtonSet.OK_CANCEL);
  if (pidResp.getSelectedButton() == ui.Button.CANCEL || pidResp.getResponseText().trim() == "") {
    ui.alert(PROMPTTITLE, ALERTMESSAGECANCEL, ui.ButtonSet.OK);
    return;
  }
  let datasetResp = ui.prompt(PROMPTTITLE, "Please provide the BigQuery Dataset name:", ui.ButtonSet.OK_CANCEL);
  if (datasetResp.getSelectedButton() == ui.Button.CANCEL || datasetResp.getResponseText().trim() == "") {
    ui.alert(PROMPTTITLE, ALERTMESSAGECANCEL, ui.ButtonSet.OK);
    return;
  }
  let tableResp = ui.prompt(PROMPTTITLE, "Please provide the BigQuery Table name:", ui.ButtonSet.OK_CANCEL);
  if (tableResp.getSelectedButton() == ui.Button.CANCEL || tableResp.getResponseText().trim() == "") {
    ui.alert(PROMPTTITLE, ALERTMESSAGECANCEL, ui.ButtonSet.OK);
    return;
  }
  let sidResp = ui.prompt(PROMPTTITLE, "Please provide the Sheet name:", ui.ButtonSet.OK_CANCEL);
  if (sidResp.getSelectedButton() == ui.Button.CANCEL || sidResp.getResponseText().trim() == "") {
    ui.alert(PROMPTTITLE, ALERTMESSAGECANCEL, ui.ButtonSet.OK);
    return;
  }
  let pidValue = pidResp.getResponseText();
  let datasetValue = datasetResp.getResponseText();
  let tableValue = tableResp.getResponseText();
  let sidValue = sidResp.getResponseText();
  PropertiesService.getDocumentProperties().setProperty("pid", pidValue.trim().toLowerCase());
  PropertiesService.getDocumentProperties().setProperty("dataset", datasetValue.trim());
  PropertiesService.getDocumentProperties().setProperty("table", tableValue.trim());
  PropertiesService.getDocumentProperties().setProperty("sheetName", sidValue.trim());
  ui.alert(PROMPTTITLE, "The configuration has been updated!", ui.ButtonSet.OK)
}

/**
 * Function to show the user the current document (Spreadsheet) configuration.
 */
function showConfiguration() {
  var props = PropertiesService.getDocumentProperties().getProperties();
  ui.alert(
    PROMPTTITLE,
    `The current configuration is as follow:\n\n` +
    `GCP Project Name : ${props["pid"]}\n` +
    `BigQuery Dataset Name : ${props["dataset"]}\n` +
    `Table name : ${props["table"]}\n` +
    `Sheet name : ${props["sheetName"]}`,
    ui.ButtonSet.OK
  );
}

/**
 * Function linked to the menu. Run the data manipulation with parameter append data.
 */
function appendData() {
  manipulateData(false);
}

/**
 * Function linked to the menu. Run the data manipulation with parameter update data.
 */
function updateData() {
  manipulateData(true);
}

/**
 * Function to gather the information, the data and the job that need to be done
 * @param {Bool} update - This parameter define if data need to be truncated (true) or simply append to the current data in BigQuery.
 */
function manipulateData(update) {
  var props = PropertiesService.getDocumentProperties().getProperties();
  if (!checkIfDatasetExists(props["pid"], props["dataset"])) {
    ui.alert(PROMPTTITLE, "The dataset doesn't exists. Please check your configuration or create the dataset in GCP console.", ui.ButtonSet.OK);
    return;
  }
  let sheet = checkIfSheetExists()
  if (sheet == null) {
    return;
  }
  if (!getAllTables(props["pid"], props["dataset"]).includes(props["table"])) {
    ui.alert(PROMPTTITLE, "The table doesnt't exists. Please check your configuration or create this table manually with GCP Console.", ui.ButtonSet.OK);
    return;
  }

  let rows = sheet.getDataRange().getValues();

  // Normalize the headers (first row) to valid BigQuery column names.
  // https://cloud.google.com/bigquery/docs/schemas#column_names
  rows[0] = rows[0].map((header) => {
    header = header.toLowerCase().replace(/[^\w]+/g, '_');
    if (header.match(/^\d/))
      header = '_' + header;
    return header;
  });

  // BigQuery load jobs can only load files, so we need to transform our
  // rows (matrix of values) into a blob (file contents as string).
  // For convenience, we convert the rows into a CSV data string.
  // https://cloud.google.com/bigquery/docs/loading-data-local
  let csvRows = rows.map(values =>
    values.map(value => JSON.stringify(value).replace(/\\"/g, '""'))
  );
  csvRows.shift();
  let csvData = csvRows.map(values => values.join(',')).join('\n');
  let blob = Utilities.newBlob(csvData, 'application/octet-stream');

  insertData(props["pid"], blob, props["dataset"], props["table"], update)
}

/**
 * Function that check if the Sheet exists in the current Spreadsheet. 
 * @returns {Sheet|null} - Return the sheet object if it's found or null if not.
 */
function checkIfSheetExists() {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PropertiesService.getDocumentProperties().getProperty('sheetName'));
    return sheet;
  } catch (err) {
    ui.alert(PROMPTTITLE, "The Sheet name in the configuration is invalid. Please check the configuration.", ui.ButtonSet.OK)
    return null;
  }
}

/**
 * Function that check if the dataset in the configuration exists in the GCP project defined by the Project ID
 * @param {String} projectId - Project name that should appear in the configuration.
 * @param {String} datasetId - Dataset name that should appear in the configuration.
 * @returns {Bool} - Return true if the process work else it return false.
 */
function checkIfDatasetExists(projectId, datasetId) {
  try {
    BigQuery.Datasets.get(projectId, datasetId);
    return true;
  } catch (err) {
    return false;
  }
}

/**
 * Function that create insert data job in BigQuery.
 * @param {String} projectId - Project name that should appear in the configuration.
 * @param {Blob} blob - The CSV blob to import in BigQuery with all fetched data.
 * @param {String} datasetId - Dataset name that should appear in the configuration.
 * @param {String} tableId - Table name that should appear in the configuration.
 * @param {Bool} update - This parameter define if data need to be truncated (true) or simply append to the current data in BigQuery.
 */
function insertData(projectId, blob, datasetId, tableId, update) {
  // Create the BigQuery load job config. For more information, see:
  // https://developers.google.com/apps-script/advanced/bigquery
  let loadJob = {
    configuration: {
      load: {
        destinationTable: {
          projectId: projectId,
          datasetId: datasetId,
          tableId: tableId
        },
        autodetect: false,  // Infer schema from contents.
        writeDisposition: (update) ? 'WRITE_TRUNCATE' : 'WRITE_APPEND',
      }
    }
  };

  try {
    BigQuery.Jobs.insert(loadJob, projectId, blob);
    ui.alert(PROMPTTITLE, `Load job started. Click here to check your jobs: https://console.cloud.google.com/bigquery?project=${projectId}&page=jobs`, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert(PROMPTTITLE, "An error occured: " + err + ". Please check the configuration and the permissions on the project", ui.ButtonSet.OK);
  }
}

function getAllTables(projectId, datasetId) {
  const options = {
    fields: "nextPageToken,tables(tableReference/tableId)"
  }
  var tables = [];
  do {
    var search = BigQuery.Tables.list(projectId, datasetId, options);
    options.pageToken = search.nextPageToken;
    if (search.tables && search.tables.length) {
      search.tables.forEach((table) => {
        tables.push(table.tableReference.tableId);
      })
    }
  } while (options.pageToken);
  return tables;
}