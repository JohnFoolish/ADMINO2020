const ss = SpreadsheetApp.getActiveSpreadsheet();
const ssData = ss.getSheetByName('Data');
const ssResponses = ss.getSheetByName('Responses');
const ssBattalion = ss.getSheetByName('Battalion Structure');
const ssOptions = ss.getSheetByName('Options');
const ssPending = ss.getSheetByName('Pending Paperwork');

/*
  const data = ssData.getRange(1,1,ssData.getLastRow(), ssData.getLastColumn()).getValues();   
  ssData.getRange(1,1,ssData.getLastRow(), ssData.getLastColumn()).setValues(data);
*/

function myOnSubmit() {
	if (ssData.getLastRow() > 0) {
		// Get newly inserted data
		const data = ssResponses.getRange(ssResponses.getLastRow(), 1, 1, ssResponses.getLastColumn()).getValues();

		// Manipulate data

		//Write to data sheet
		ssData.getRange(ssData.getLastRow(), 1, ssData.getLastRow() + 1, data[0].length).setValues(data);
	}
}

function myOnEdit() {
	// Not sure what is going to end up going here
}
