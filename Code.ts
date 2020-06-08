const ss = SpreadsheetApp.getActiveSpreadsheet();
const ssData = ss.getSheetByName('Data');
const ssResponses = ss.getSheetByName('Responses');
const ssBattalion = ss.getSheetByName('Battalion Structure');

/*
  const data = ssData.getRange(1,1,ssData.getLastRow(), ssData.getLastColumn()).getValues();   
  ssData.getRange(1,1,ssData.getLastRow(), ssData.getLastColumn()).setValues(data);
*/

function myOnSubmit() {
	const data = ssResponses.getRange(ssResponses.getLastRow(), 1, 1, ssResponses.getLastColumn()).getValues();
}

function myOnEdit() {
	const data = ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).getValues();
	ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).setValues(data);
}
