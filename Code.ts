const ss = SpreadsheetApp.getActiveSpreadsheet();
const ssData = ss.getSheetByName('Data');

/*
  const data = ssData.getRange(1,1,ssData.getLastRow(), ssData.getLastColumn()).getValues();   
  ssData.getRange(1,1,ssData.getLastRow(), ssData.getLastColumn()).setValues(data);

*/

function myOnSubmit() {
	console.log('Hello world');

	const doSomething = () => {
		console.log('Litterally anything');
	};

	doSomething();

	//THis is a comment htat got pushed automatically
	// Really cool comment entered here
	// Hey there bud what about this
	//This will be a cool thing to edit

	// Written using vs code
}

function myOnEdit() {
	const data = ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).getValues();
	ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).setValues(data);
}
