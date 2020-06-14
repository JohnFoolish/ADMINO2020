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

function test() {
	getIndividualsInGroup('DHs');
	Logger.log(getIndividualEmail('Bowes, Timothy'));
}

function myOnSubmit() {
	if (ssData.getLastRow() > 0) {
		// Get newly inserted data
		const data = ssResponses.getRange(ssResponses.getLastRow(), 1, 1, ssResponses.getLastColumn()).getValues();

		// Manipulate data
		// Go from [[Timestamp	Assigner's Name	Receiever Name/Group	Paperwork	Reason for paperwork	Date Assigned	Date Due	Send Initial Email Notification]]
		// To this [[Timestamp	Assigner's Name	Group	Receiver's Name	Paperwork	Date Assigned	Date Due	Received	Reason for paperwork]]

		//Write to data sheet
		ssData.getRange(ssData.getLastRow() + 1, 1, 1, data[0].length).setValues(data);

		//Check to see if we need to send the email to the recipient
		if (data[7] === true)
	}
}

function myOnEdit() {
	if (
		ss.getActiveCell().getSheet().getName() === 'Battalion Structure' &&
		(ss.getActiveCell().getColumn() === 1 || ss.getActiveCell().getColumn() === 2 || ss.getActiveCell().getRow() === 1)
	) {
		updateFormGroups();
	}
}

//Make the sheet pretty function--once a week?

function updateFormGroups() {
	// Update Recieve name / group
	const FormItem = form.getItems();
	const item = FormItem[1].asListItem();
	item.setTitle('Receiever Name/Group');
	const groups = getGroups(false);
	const groupList = [];
	for (const groupData of groups) {
		groupList.push(item.createChoice(groupData));
	}
	item.setChoices(groupList);
	item.isRequired();
	item.setHelpText('The group or MIDN you want to assign the paperwork to');
	Logger.log(groupList);

	// Update assigner names list
	const item2 = FormItem[0].asListItem();
	const ind = getGroups(true);
	const indList = [];
	for (const individuals of ind) {
		indList.push(item2.createChoice(individuals));
	}
	item2.setChoices(indList);
	item2.isRequired();
	item2.setHelpText('Select your name from the list below.');
	Logger.log(indList);
}

function getGroups(justIndividuals: boolean): string[] {
	const groupData = ssBattalion.getRange(1, 1, ssBattalion.getLastRow(), ssBattalion.getLastColumn()).getValues();
	const out = [];

	if (!justIndividuals) {
		for (let i = 3; i < groupData[0].length; i++) {
			const group = groupData[0][i];
			if (group !== '') {
				out.push(group);
			}
		}
	}

	for (let i = 1; i < groupData.length; i++) {
		const person = groupData[i][0] + ', ' + groupData[i][1];
		if (groupData[i][0] !== '' && groupData[i][1] !== '') {
			out.push(person);
		}
	}

	return out;
}

function getIndividualEmail(name: string): string {
	const groupData = ssBattalion.getRange(1, 1, ssBattalion.getLastRow(), ssBattalion.getLastColumn()).getValues();
	let returnEmail = '';
	for (let i = 1; i < groupData.length; i++) {
		const person = groupData[i][0] + ', ' + groupData[i][1];
		if (name === person) {
			returnEmail = groupData[i][2];
		}
	}
	Logger.log(returnEmail, name);
	return returnEmail;
}

function getIndividualsInGroup(groupName: string): string[] {
	const groupData = ssBattalion.getRange(1, 1, ssBattalion.getLastRow(), ssBattalion.getLastColumn()).getValues();
	const out = [];

	const columnOfGroup = groupData[0].indexOf(groupName);

	if (columnOfGroup !== -1) {
		for (let i = 1; i < groupData.length; i++) {
			if (groupData[i][columnOfGroup] !== '') out.push(groupData[i][0] + ', ' + groupData[i][1]);
		}
	}

	Logger.log(out);
	return out === [] ? [groupName] : out;
}
