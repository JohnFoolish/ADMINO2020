const ss = SpreadsheetApp.getActiveSpreadsheet();
const ssData = ss.getSheetByName('Data');
const ssAssignment = ss.getSheetByName('Assignment Responses');
const ssTurnedIn = ss.getSheetByName('Turnin Responses');
const ssBattalion = ss.getSheetByName('Battalion Structure');
const ssOptions = ss.getSheetByName('Options');
const ssPending = ss.getSheetByName('Pending Paperwork');
const ssVariables = ss.getSheetByName('Variables');
const ssDigitalBox = ss.getSheetByName('Digital Turn In Box');
const ssBattalionStructure = ss.getSheetByName('New Battalion Structure');
const ui = SpreadsheetApp.getUi();

const form = FormApp.openByUrl('https://docs.google.com/forms/d/1l6lZZhsOWb5rcyTFDxyiFJln0tFBuVIiFRGK_hjnZ84/edit');
const subForm = FormApp.openByUrl('https://docs.google.com/forms/d/1x2HP45ygThm6MoYlKasVnaacgZUW_yKA7Cz9pxKKOJc/edit');

function test() {
	Logger.log('start');
	chainOfCommandStructureUpdater();
	Logger.log('end');
}

//Triggers when the submission form is submitted
function myOnSubmit() {
	if (ssVariables.getRange(1, 2).getValue().toString() !== ssAssignment.getLastRow().toString()) {
		myOnAssignmentSubmit();
		ssVariables.getRange(1, 2).setValue(ssAssignment.getLastRow());
	}
	if (ssVariables.getRange(2, 2).getValue().toString() !== ssTurnedIn.getLastRow().toString()) {
		myOnFormTurnedInSubmit();
		ssVariables.getRange(2, 2).setValue(ssTurnedIn.getLastRow());
	}
}

function myOnAssignmentSubmit() {
	if (ssData.getLastRow() > 0) {
		// Get newly inserted data
		const data = ssAssignment.getRange(ssAssignment.getLastRow(), 1, 1, ssAssignment.getLastColumn()).getValues();

		// Manipulate data
		const people = getIndividualsInGroup(data[0][2]);
		const outData = new Array(people.length);
		const emailList = [];
		for (let i = 0; i < people.length; i++) {
			outData[i] = new Array(9);
			outData[i][0] = new Date(data[0][0].toString());
			outData[i][0].setSeconds(outData[i][0].getSeconds() + i); //Timestamp - UUID
			outData[i][1] = data[0][1]; // Assigners Name
			outData[i][2] = people.length === 1 ? 'Individual' : data[0][2]; // Group
			outData[i][3] = people[i]; // Recievers name
			outData[i][4] = data[0][3]; // Paperwork
			outData[i][5] = data[0][5]; // Data assigned
			if (outData[i][4] === 'Chit' || outData[i][4] === 'Negative Counseling' || outData[i][4] === 'Merit') {
				//This function does not work. Placeholder for now.
				const date = increaseDate(outData[i][4], outData[i][5]);
				outData[i][6] = date;
			} else {
				outData[i][6] = data[0][6]; // Date Due
			}
			outData[i][7] = 'FALSE'; // Turned in
			outData[i][8] = data[0][4]; // Reason for paperwork
			outData[i][9] = data[0][8]; //Link to paperwork

			if (data[0][7] == 'Yes') {
				emailList.push(getIndividualEmail(people[i]));
			}
		}
		sendEmail(emailList, data);

		//Write to data sheet
		ssData.getRange(ssData.getLastRow() + 1, 1, outData.length, outData[0].length).setValues(outData);

		// Write to Pending Paperwork
		ssPending.getRange(ssPending.getLastRow() + 1, 1, outData.length, outData[0].length).setValues(outData);
	}
}

//This function runs whenever the new paperwork submission form is submitted.
function myOnFormTurnedInSubmit() {
	// Get data from linked sheet to use
	const data = ssTurnedIn.getRange(ssTurnedIn.getLastRow(), 1, 1, ssTurnedIn.getLastColumn()).getValues();

	// Manipulate Data / Rearrange Data
	const outData = data;
	outData[0].push('FALSE');

	// Write to Digital Admin box sheet
	ssDigitalBox.getRange(ssDigitalBox.getLastRow() + 1, 1, outData.length, outData[0].length).setValues(outData);
	sortDigitalBox();
}

function sortDigitalBox() {
	if (ssDigitalBox.getLastRow() > 1) {
		ssDigitalBox.getRange(2, 1, ssDigitalBox.getLastRow() - 1, ssDigitalBox.getLastColumn()).sort(1);
		ssDigitalBox.getRange(2, 1, ssDigitalBox.getLastRow() - 1, ssDigitalBox.getLastColumn()).sort(4);
	}
}

function myOnEdit() {
	if (
		ss.getActiveCell().getSheet().getName() === 'Battalion Structure' &&
		(ss.getActiveCell().getColumn() === 1 || ss.getActiveCell().getColumn() === 2 || ss.getActiveCell().getRow() === 1)
	) {
		updateFormGroups();
	} else if (ss.getActiveCell().getSheet().getName() === 'Pending Paperwork' && ss.getActiveCell().getColumn() === 8) {
		const pending = ssPending.getRange(1, 1, ssPending.getLastRow(), ssPending.getLastColumn()).getValues();
		const data = ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).getValues();
		for (let j = 1; j < pending.length; j++) {
			if (pending[j][7].toString() === 'true') {
				if (pending[j][9] === '') {
					ui.alert('You need to put either "Turned in Physically" or the link to their digitally turned in file');
					pending[j][7] = 'false';
				} else {
					const uuidDate = pending[j][0].toString();
					for (let i = 0; i < data.length; i++) {
						if (data[i][0].toString() === uuidDate) {
							data[i][7] = 'true';
							data[i][9] = pending[j][9];
						}
					}
					pending[j] = pending[j].map((item) => '');
				}
			}
		}
		ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).setValues(data);
		ssPending.getRange(1, 1, ssPending.getLastRow(), ssPending.getLastColumn()).setValues(pending);
		ssPending.sort(1);
	} else if (
		ss.getActiveCell().getSheet().getName() === 'Digital Turn In Box' &&
		ss.getActiveCell().getColumn() === 4
	) {
		sortDigitalBox();
	} else if (
		ss.getActiveCell().getSheet().getName() === 'New Battalion Structure' &&
		ss.getActiveCell().getColumn() > 2
	) {
		chainOfCommandStructureUpdater();
	}
}

// ssBattalionStructure.getRange(2, 3).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInRange(groupsRange).build());

function chainOfCommandStructureUpdater() {
	if (ssBattalionStructure.getLastRow() > 1) {
		// Create list of all groups remaining
		const groups = [];
		ssBattalionStructure
			.getRange(2, 2, ssBattalionStructure.getLastRow(), 1)
			.getValues()
			.forEach((row) => {
				row.forEach((node) => {
					if (node !== '') {
						groups.push(node);
					}
				});
			});
		// Read the chain to figure out what the structure is
		interface chain {
			value: string;
			children: chain[];
			pos: number[];
		}
		let chainOfCommand = {} as chain;
		let previousLevel = [] as chain[];
		const data = ssBattalionStructure
			.getRange(1, 1, ssBattalionStructure.getLastRow(), ssBattalionStructure.getLastColumn())
			.getValues();
		for (let row = 1; row < data.length; row++) {
			for (let col = 2; col < data[0].length; col++) {
				const gridValue = data[row][col];
				if (row === 1) {
					if (col === 2) {
						if (groups.indexOf(gridValue) > -1) {
							chainOfCommand.value = gridValue;
							groups.splice(groups.indexOf(gridValue), 1);
							chainOfCommand.pos = [row, col];
							chainOfCommand.children = [];
						} else {
							data[row][col] = '';
						}
					} else {
						data[row][col] = '';
					}
				} else {
					if (groups.indexOf(gridValue) > -1) {
						groups.splice(groups.indexOf(gridValue), 1);
						let CoCnode = {} as chain;
						CoCnode.pos = [row, col];
						CoCnode.value = gridValue;
						CoCnode.children = [];
						let parent;
						for (let i = 0; i < previousLevel.length; i++) {
							if (previousLevel[i].pos[1] <= col) parent = previousLevel[i];
						}
						parent.children.push(CoCnode);
					} else {
						data[row][col] = '';
					}
				}
			}
			if (row === 1) {
				previousLevel = [chainOfCommand];
			} else {
				const outPreviousLevel = [];
				previousLevel.forEach((node) => {
					node.children.forEach((child) => {
						outPreviousLevel.push(child);
					});
				});
				previousLevel = outPreviousLevel;
			}
		}
		ssBattalionStructure
			.getRange(1, 1, ssBattalionStructure.getLastRow(), ssBattalionStructure.getLastColumn())
			.setValues(data);

		// Update the interface so it can be added to if needed

		//Write json chainOfCOmmand to variables sheet
		ssVariables.getRange(3, 2).setValue(JSON.stringify(chainOfCommand));
	}
}

function createGoogleFiles() {
	const root = DriveApp.getFolderById('1vPucUC-lnMzCRWPZQ8FYkQHswNkB7Nv9');
	const battalionIndividuals = getGroups(true);
	var ssTemplate = SpreadsheetApp.openByUrl(
		'https://docs.google.com/spreadsheets/d/1QbC9z04dQWhDNz-Q4qm2urfWzZjtKxLmmCsFQ4J9uOU/edit#gid=0'
	);
	const templateID = ssTemplate.getId();
	const newFile = DriveApp.getFileById(templateID);
	for (var idx = 0; idx < battalionIndividuals.length; idx++) {
		const email = getIndividualEmail(battalionIndividuals[idx]);
		if (email === '') {
			continue;
		}
		const indFile = newFile.makeCopy(battalionIndividuals[idx] + ', GT NROTC', root);
		const indID = indFile.getId();
		initSheet(indID, battalionIndividuals[idx]);
		indFile.addViewer(email);
		indFile.addEditor('gtnrotc.ado@gmail.com');
	}
}

function findIndSheet(name) {
	var files = DriveApp.getFilesByName(name + ', GT NROTC');
	const fileList = [];
	while (files.hasNext()) {
		var sheet = files.next();
		fileList.push(sheet);
	}
	return files;
}

function wipeGoogleFiles() {
	const root = DriveApp.getFolderById('1vPucUC-lnMzCRWPZQ8FYkQHswNkB7Nv9');
	const battalionIndividuals = getGroups(true);
	for (var ind = 0; ind < battalionIndividuals.length; ind++) {
		const email = getIndividualEmail(battalionIndividuals[ind]);
		if (email === '') {
			continue;
		}
		const fileList = findIndSheet(battalionIndividuals[ind]);
		while (fileList.hasNext()) {
			var file = fileList.next();
			root.removeFile(file);
		}
	}
}

function initSheet(sheetID, name) {
	const userSpread = SpreadsheetApp.openById(sheetID);
	const userPaperwork = userSpread.getSheetByName('Total_Paperwork');
	const outData = userPaperwork.getRange(1, 1, userPaperwork.getLastRow(), userPaperwork.getLastColumn()).getValues();

	// Manipulate outData according to incomingData

	var chits = 0;
	var merits = 0;
	var negCounsel = 0;
	const pending = ssPending.getRange(1, 1, ssPending.getLastRow(), ssPending.getLastColumn()).getValues();
	for (var i = 1; i < pending.length; i++) {
		if (pending[i][4] === name) {
			if (pending[i][5] === 'Chit') {
				chits++;
			} else if (pending[i][5] === 'Negative Counseling') {
				negCounsel++;
			} else if (pending[i][5] === 'Merit') {
				merits++;
			}
			userPaperwork[userPaperwork.getLastRow()].setValues(pending[i]);
		}
	}
	const helpData = [];
	helpData.push(chits);
	helpData.push(negCounsel);
	helpData.push(merits);
	userPaperwork.getRange(1, 1, 2, 3).setValues(helpData);

	//userPaperwork.getRange(1, 1, outData.length, outData[0].length).setValues(outData);
}

function updateFormGroups() {
	// Update Recieve name / group
	const FormItem = form.getItems();
	const item = FormItem[1].asListItem();
	item.setTitle('Receiver Name/Group');
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

	//Update the form submission page
	const subFormItem = subForm.getItems();
	const subItem = subFormItem[0].asListItem();
	subItem.setTitle('Your name');
	const subInd = getGroups(true);
	const subIndList = [];
	for (const subIndividuals of subInd) {
		subIndList.push(item2.createChoice(subIndividuals));
	}
	subItem.setChoices(subIndList);
	subItem.isRequired();
	subItem.setHelpText('Select your name from the dropdown menu below');
	Logger.log(subIndList);
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
	let returnEmail = null;
	for (let i = 1; i < groupData.length; i++) {
		const person = groupData[i][0] + ', ' + groupData[i][1];
		if (name === person) {
			returnEmail = groupData[i][2];
		}
	}
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

	return out.length === 0 ? [groupName] : out;
}

function sendEmail(emailList, data) {
	const emailsActivated = ssOptions.getRange(1, 2).getValue().toString().toLowerCase() === 'true';
	if (!emailsActivated) return;

	const dateDemo = String(data[0][6]).split(' ', 4);

	const date = dateDemo[0] + ', ' + dateDemo[2] + dateDemo[1].toUpperCase() + dateDemo[3];

	const emailSender = getIndividualEmail(data[0][0]);

	const emailSubject = 'NROTC ADMIN Department: New ' + data[0][3] + ' due COB ' + date + '.';

	const emailBody =
		"<h2 'style=color: #5e9ca0;'> You have been assigned a " +
		data[0][3] +
		' from ' +
		data[0][1] +
		'.</h2>' +
		'<p> The reason is the following: ' +
		data[0][4] +
		'.</p> <p> You must turn this form in by COB on ' +
		date +
		'.</p>' +
		'<p> If you have any questions regarding the validity of the ' +
		data[0][3] +
		', please contact the assignee. </p>' +
		'<p> You can find the paperwork to complete here: ' +
		data[0][8] +
		'</p>';

	//emailList.filter((email) => email !== '');
	var correctedEmail = '';
	for (let i = 0; i < emailList.length; i++) {
		if (emailList[i] === null) {
			continue;
		} else {
			if (correctedEmail === '') {
				correctedEmail = emailList[i];
			} else {
				correctedEmail = emailList[i] + ',' + correctedEmail;
			}
		}
	}
	Logger.log(emailList, emailSender);
	MailApp.sendEmail({
		to: emailSender,
		bcc: correctedEmail,
		subject: emailSubject,
		htmlBody: emailBody,
	});
}

function increaseDate(paperworkType, rawDate) {
	//const dateDemo = String(rawDate).split(" ", 4);
	//const date = dateDemo
	return rawDate;
	//const date = new Date(this.valueOf());
	//date.setDate(date.getDate() + 3);
}
