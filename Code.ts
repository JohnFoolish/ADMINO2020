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

	// Are you sure that you cant?
	//Nope I cannot :(
	//THis is a comment htat got pushed automatically
	// Really cool comment entered here
	// Hey there bud what about this
	//This will be a cool thing to edit
	//This is a new save!

	// Written using vs code

	//Added to github and also now in app scripts
}

function myOnEdit() {
	const data = ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).getValues();
	ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).setValues(data);
}
