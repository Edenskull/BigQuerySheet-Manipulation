<div align="center">
  <img src="https://raw.githubusercontent.com/Edenskull/BigQuerySheet-Manipulation/main/.github/assets/Banner.png">
</div>

<div align="center">

[![GitHub license](https://img.shields.io/github/license/Edenskull/BigQuerySheet-Manipulation?color=blue&style=for-the-badge)](https://github.com/Edenskull/BigQuerySheet-Manipulation/blob/master/LICENSE)
![Google](https://img.shields.io/badge/Google-Appscript-yellow?style=for-the-badge)

</div>

# BigQuerySheet-Manipulation
Google App Scripts to link Google sheets to bigquery table

# Before you edit the source code
`Before you edit the source code note that every functions below are already implemented in the workflow and can stop working if you change anything. Be sure to check the whole code before editing.`

## Functions

<dl>
<dt><a href="#onOpen">onOpen()</a></dt>
<dd><p>Function to trigger the menu creation when the spreadsheet is open by the user</p>
</dd>
<dt><a href="#updateConfiguration">updateConfiguration()</a></dt>
<dd><p>Function to update the current document (Spreadsheet) configuration. With this all the collaborator using a same Spreadsheet can execute event with the same configuration.</p>
</dd>
<dt><a href="#showConfiguration">showConfiguration()</a></dt>
<dd><p>Function to show the user the current document (Spreadsheet) configuration.</p>
</dd>
<dt><a href="#appendData">appendData()</a></dt>
<dd><p>Function linked to the menu. Run the data manipulation with parameter append data.</p>
</dd>
<dt><a href="#updateData">updateData()</a></dt>
<dd><p>Function linked to the menu. Run the data manipulation with parameter update data.</p>
</dd>
<dt><a href="#manipulateData">manipulateData(update)</a></dt>
<dd><p>Function to gather the information, the data and the job that need to be done</p>
</dd>
<dt><a href="#checkIfSheetExists">checkIfSheetExists()</a> ⇒ <code>Sheet</code> | <code>null</code></dt>
<dd><p>Function that check if the Sheet exists in the current Spreadsheet.</p>
</dd>
<dt><a href="#checkIfDatasetExists">checkIfDatasetExists(projectId, datasetId)</a> ⇒ <code>Bool</code></dt>
<dd><p>Function that check if the dataset in the configuration exists in the GCP project defined by the Project ID</p>
</dd>
<dt><a href="#createDataset">createDataset(projectId, datasetId)</a> ⇒ <code>Bool</code></dt>
<dd><p>Function that will create the dataset in the GCP project defined by the Project ID.</p>
</dd>
<dt><a href="#insertData">insertData(projectId, blob, datasetId, tableId, update)</a></dt>
<dd><p>Function that create insert data job in BigQuery.</p>
</dd>
</dl>

<a name="onOpen"></a>

## onOpen()
Function to trigger the menu creation when the spreadsheet is open by the user

**Kind**: global function
<a name="updateConfiguration"></a>

## updateConfiguration()
Function to update the current document (Spreadsheet) configuration. With this all the collaborator using a same Spreadsheet can execute event with the same configuration.

**Kind**: global function
<a name="showConfiguration"></a>

## showConfiguration()
Function to show the user the current document (Spreadsheet) configuration.

**Kind**: global function
<a name="appendData"></a>

## appendData()
Function linked to the menu. Run the data manipulation with parameter append data.

**Kind**: global function
<a name="updateData"></a>

## updateData()
Function linked to the menu. Run the data manipulation with parameter update data.

**Kind**: global function
<a name="manipulateData"></a>

## manipulateData(update)
Function to gather the information, the data and the job that need to be done

**Kind**: global function

| Param | Type | Description |
| --- | --- | --- |
| update | <code>Bool</code> | This parameter define if data need to be truncated (true) or simply append to the current data in BigQuery. |

<a name="checkIfSheetExists"></a>

## checkIfSheetExists() ⇒ <code>Sheet</code> \| <code>null</code>
Function that check if the Sheet exists in the current Spreadsheet.

**Kind**: global function
**Returns**: <code>Sheet</code> \| <code>null</code> - - Return the sheet object if it's found or null if not.
<a name="checkIfDatasetExists"></a>

## checkIfDatasetExists(projectId, datasetId) ⇒ <code>Bool</code>
Function that check if the dataset in the configuration exists in the GCP project defined by the Project ID

**Kind**: global function
**Returns**: <code>Bool</code> - - Return true if the process work else it return false.

| Param | Type | Description |
| --- | --- | --- |
| projectId | <code>String</code> | Project name that should appear in the configuration. |
| datasetId | <code>String</code> | Dataset name that should appear in the configuration. |

<a name="createDataset"></a>

## createDataset(projectId, datasetId) ⇒ <code>Bool</code>
Function that will create the dataset in the GCP project defined by the Project ID.

**Kind**: global function
**Returns**: <code>Bool</code> - - Return true if the process work else it return false.

| Param | Type | Description |
| --- | --- | --- |
| projectId | <code>String</code> | Project name that should appear in the configuration. |
| datasetId | <code>String</code> | Dataset name that should appear in the configuration. |

<a name="insertData"></a>

## insertData(projectId, blob, datasetId, tableId, update)
Function that create insert data job in BigQuery.

**Kind**: global function

| Param | Type | Description |
| --- | --- | --- |
| projectId | <code>String</code> | Project name that should appear in the configuration. |
| blob | <code>Blob</code> | The CSV blob to import in BigQuery with all fetched data. |
| datasetId | <code>String</code> | Dataset name that should appear in the configuration. |
| tableId | <code>String</code> | Table name that should appear in the configuration. |
| update | <code>Bool</code> | This parameter define if data need to be truncated (true) or simply append to the current data in BigQuery. |