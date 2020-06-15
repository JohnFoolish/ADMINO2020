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

function test() {}

function myOnSubmit() {
	if (ssData.getLastRow() > 0) {
		// Get newly inserted data
		const data = ssResponses.getRange(ssResponses.getLastRow(), 1, 1, ssResponses.getLastColumn()).getValues();
		

		// Manipulate data
		const people = getIndividualsInGroup(data[0][2]);
		const outData = new Array(people.length);
		const emailList = [];
		for (let i = 0; i < people.length; i++) {
            outData[i] = new Array(9)
			outData[i][0] = new Date();
			outData[i][0].setTime(data[0][0].getTime() + i); //Timestamp - UUID
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

function myOnEdit() {
	if (
		ss.getActiveCell().getSheet().getName() === 'Battalion Structure' &&
		(ss.getActiveCell().getColumn() === 1 || ss.getActiveCell().getColumn() === 2 || ss.getActiveCell().getRow() === 1)
	) {
		updateFormGroups();
	}
	if (ss.getActiveCell().getSheet().getName() === 'Pending Paperwork' && ss.getActiveCell().getColumn() === 8) {
		const uuidDate = ssPending.getRange(ss.getActiveCell().getRow(), 1).getValue();
		const data = ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).getValues();
		for (let i = 0; i < data.length; i++) {
			if (data[i][0] == uuidDate) {
				data[i][7] = 'TRUE';
			}
		}
		ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).setValues(data);
		ssPending.deleteRow(ss.getActiveCell().getRow());
		Logger.log(uuidDate, data)
	}
}

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


	const dateDemo = String(data[0][6]).split(" ", 4);

	const date = dateDemo[0] + ", " + dateDemo[2] + dateDemo[1].toUpperCase() + dateDemo[3]

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
		'.<p>' + "<p> If you have any questions regarding the validity of the " + data[0][3] + ", please contact the assignee.";

	//emailList.filter((email) => email !== '');
    var correctedEmail = "";
	for (let i = 0; i < emailList.length; i++) {
        if (emailList[i] === null) {
			continue;
		} else {
			if (correctedEmail === "") {
				correctedEmail = emailList[i];
			} else {
			    correctedEmail = emailList[i] + "," + correctedEmail;
			}
		}
	}
    Logger.log(emailList, emailSender)
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
	return rawDate
	//const date = new Date(this.valueOf());
	//date.setDate(date.getDate() + 3);
}
