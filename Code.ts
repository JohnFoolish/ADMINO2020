const ss = SpreadsheetApp.getActiveSpreadsheet();
const ssData = ss.getSheetByName('Data');
const ssResponses = ss.getSheetByName('Responses');
const ssBattalion = ss.getSheetByName('Battalion Structure');
const ssOptions = ss.getSheetByName('Options');
const ssPending = ss.getSheetByName('Pending Paperwork');

const form = FormApp.openByUrl('https://docs.google.com/forms/d/1l6lZZhsOWb5rcyTFDxyiFJln0tFBuVIiFRGK_hjnZ84/edit');

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
		ssData.getRange(ssData.getLastRow() + 1, 1, 1, data[0].length).setValues(data);
	}
}

function myOnEdit() {
	// Not sure what is going to end up going here
}

//Make the sheet pretty function--once a week?

function updateFormGroups() {
	const item = form.addListItem();
	item.setTitle('Receiever Name/Group');
	const groups = getGroups();
	const groupList = [];
	for (const groupData of groups) {
		groupList.push(item.createChoice(groupData));
	}
	item.setChoices(groupList);
	item.isRequired();
	item.setHelpText('The group or MIDN you want to assign the paperwork to');
	Logger.log(groupList);
}

function getGroups(): string[] {
	const groupData = ssBattalion.getRange(1, 1, ssBattalion.getLastRow(), ssBattalion.getLastColumn()).getValues();
	const out = [];

	for (let i = 3; i < groupData[0].length; i++) {
		const group = groupData[0][i];
		if (group !== '') {
			out.push(group);
		}
	}

	for (let i = 1; i < groupData.length; i++) {
		const person = groupData[i][0] + groupData[i][1];
		if (person !== '') {
			out.push(person);
		}
	}

	Logger.log(out);
	return out;
}
