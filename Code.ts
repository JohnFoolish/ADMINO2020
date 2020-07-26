// This code was complited from typescript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const ssData = ss.getSheetByName('Data');
let ssAssignment = ss.getSheetByName('Assignment Responses');
const ssTurnedIn = ss.getSheetByName('Turnin Responses');
const ssOptions = ss.getSheetByName('Options');
const ssPending = ss.getSheetByName('Pending Paperwork');
const ssVariables = ss.getSheetByName('Variables');
const ssDigitalBox = ss.getSheetByName('Digital Turn In Box');
const ssBattalionStructure = ss.getSheetByName('Battalion Structure');
const ssBattalionMembers = ss.getSheetByName('Battalion Members');

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
		const dataResponseFormat = ssAssignment.getRange(1, 1, 1, ssAssignment.getLastColumn()).getValues();
		const submittedData = ssAssignment
			.getRange(ssAssignment.getLastRow(), 1, 1, ssAssignment.getLastColumn())
			.getValues();
		const dataPairs = [dataResponseFormat[0], submittedData[0]];

		interface submittedData {
			timestamp: Date;
			assigner: string;
			paperwork: string;
			reason: string;
			dateAssigned: Date;
			dateDue: Date;
			sendEmail: boolean;
			pdfLink: string;
		}

		const submitData = {} as submittedData;
		const keyValuePairsRawGridCheckbox = [];
		for (let i = 0; i < dataPairs[0].length; i++) {
			if (dataPairs[0][i] === 'Timestamp') {
				submitData.timestamp = dataPairs[1][i];
			} else if (dataPairs[0][i] === "Assigner's Name") {
				submitData.assigner = dataPairs[1][i];
			} else if (dataPairs[0][i] === 'Paperwork') {
				submitData.paperwork = dataPairs[1][i];
			} else if (dataPairs[0][i] === 'Reason for paperwork') {
				submitData.reason = dataPairs[1][i];
			} else if (dataPairs[0][i] === 'Date Assigned') {
				submitData.dateAssigned = dataPairs[1][i];
			} else if (dataPairs[0][i] === 'Date Due') {
				submitData.dateDue = specificDueDateLengthCheck(submitData.paperwork, submitData.dateAssigned, dataPairs[1][i]);
			} else if (dataPairs[0][i] === 'Send Initial Email Notification') {
				submitData.sendEmail = dataPairs[1][i] === 'No' ? false : true;
			} else if (dataPairs[0][i] === 'Upload your form as a PDF here:') {
				submitData.pdfLink = dataPairs[1][i];
			} else if (dataPairs[0][i].substring(0, 22) === 'Receiving Individual/s') {
				if (dataPairs[1][i] !== '') {
					keyValuePairsRawGridCheckbox.push({
						role: dataPairs[0][i].substring(24, dataPairs[0][i].length - 1),
						groups: dataPairs[1][i].split(',').map((element) => element.trim()),
					});
				}
			} else if (dataPairs[0][i].substring(0, 18) === 'Receiving Groups/s') {
				if (dataPairs[1][i] !== '') {
					keyValuePairsRawGridCheckbox.push({
						role: dataPairs[0][i].substring(20, dataPairs[0][i].length - 1),
						groups: dataPairs[1][i].split(',').map((element) => element.trim()),
					});
				}
			}
		}

		// Make sure assigner is in system
		let assignerFullData;
		JSON.parse(ssVariables.getRange(4, 2).getValue()).forEach((member) => {
			if (member.name === submitData.assigner) {
				assignerFullData = member;
			}
		});

		// Manipulate data
		const people = getIndividualsFromCheckBoxGrid(keyValuePairsRawGridCheckbox, assignerFullData);
		const outData = [];
		const emailList = [];
		const noAuthority = [];
		for (let i = 0; i < people.length; i++) {
			if (people[i].canBeAssignedFromAssigner) {
				const tempOutData = new Array(9);
				tempOutData[0] = new Date(submitData.timestamp);
				tempOutData[0].setSeconds(tempOutData[0].getSeconds() + i); //Timestamp - UUID
				tempOutData[1] = submitData.assigner; // Assigners Name
				tempOutData[2] = people[i].group; // Group
				tempOutData[3] = people[i].name; // Recievers name
				tempOutData[4] = submitData.paperwork; // Paperwork
				tempOutData[5] = submitData.dateAssigned; // Date assigned
				tempOutData[6] = submitData.dateDue; // Date Due
				tempOutData[7] = 'Pending'; // Status
				tempOutData[8] = submitData.reason; // Reason for paperwork
				tempOutData[9] = submitData.pdfLink; //Link to paperwork
				outData.push(tempOutData);

				if (submitData.sendEmail) {
					emailList.push(getIndividualEmail(people[i].name));
				}
			} else {
				noAuthority.push(people[i]);
			}
		}
		//Send ouot email notifiying everyone that their paperwork was assigned
		sendEmail(emailList, submitData);
		//Email the assigner who was assigned it and who was not
		Logger.log(noAuthority);

		//Write to data sheet
		ssData.getRange(ssData.getLastRow() + 1, 1, outData.length, outData[0].length).setValues(outData);

		// Write to Pending Paperwork
		ssPending.getRange(ssPending.getLastRow() + 1, 1, outData.length, outData[0].length).setValues(outData);
	}
}

function specificDueDateLengthCheck(paperwork: string, assignDate: Date, specifiedDueDate): Date {
	let out = specifiedDueDate;
	if (paperwork === 'Chit') {
		out = new Date(assignDate.toString());
		let chitTime = ssOptions.getRange(2, 2).getValue();
		if (typeof parseInt(chitTime) != 'number' || chitTime === '') {
			chitTime = '3';
		}
		out.setDate(out.getDate() + adjustDateForWeekends(out, parseInt(chitTime)));
	} else if (paperwork === 'Negative Counseling') {
		out = new Date(assignDate.toString());
		let ncTime = ssOptions.getRange(3, 2).getValue();
		if (typeof parseInt(ncTime) != 'number' || ncTime === '') {
			ncTime = '3';
		}
		out.setDate(out.getDate() + adjustDateForWeekends(out, parseInt(ncTime)));
	} else if (out === '') {
		out = new Date();
		out.setDate(out.getDate() + 7);
	}
	return out;
}

function adjustDateForWeekends(currentDate, daysToAddToDate): number {
	let daysAdded = 0;
	const maniputateDate = new Date(currentDate.toString());

	for (let i = 0; i < daysToAddToDate; i++) {
		maniputateDate.setDate(maniputateDate.getDate() + 1);
		if (maniputateDate.getDay() === 0 || maniputateDate.getDay() === 6) {
			i--;
		}
		daysAdded++;
	}
	Logger.log('Days added to assignment' + daysAdded);

	return daysAdded;
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
	if (ss.getActiveCell().getSheet().getName() === 'Battalion Members') {
		if (
			ss.getActiveCell().getColumn() === 1 ||
			ss.getActiveCell().getColumn() === 2 ||
			ss.getActiveCell().getColumn() === 3
		) {
			updateFormGroups();
		}
		updateBattalionMembersJSON();
	} else if (ss.getActiveCell().getSheet().getName() === 'Pending Paperwork' && ss.getActiveCell().getColumn() === 8) {
		const pending = ssPending.getRange(1, 1, ssPending.getLastRow(), ssPending.getLastColumn()).getValues();
		const data = ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).getValues();
		let oneWasTrue = false;
		for (let j = 1; j < pending.length; j++) {
			if (pending[j][7].toString() !== 'Pending' && pending[j][7].toString() !== '') {
				oneWasTrue = true;
				if (pending[j][9] === '' && pending[j][7] === 'Approved') {
					const ui = SpreadsheetApp.getUi();
					ui.alert('You need to put either "Turned in Physically" or the link to their digitally turned in file');
					pending[j][7] = 'Pending';
				} else {
					const uuidDate = pending[j][0].toString();
					for (let i = 0; i < data.length; i++) {
						if (data[i][0].toString() === uuidDate) {
							data[i][7] = pending[j][7];
							data[i][9] = pending[j][9];
						}
					}
					pending[j] = pending[j].map((item) => '');
				}
			}
		}
		if (oneWasTrue) {
			ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).setValues(data);
			ssPending.getRange(1, 1, ssPending.getLastRow(), ssPending.getLastColumn()).setValues(pending);
			ssPending.sort(1);
		}
	} else if (
		ss.getActiveCell().getSheet().getName() === 'Digital Turn In Box' &&
		ss.getActiveCell().getColumn() === 4
	) {
		sortDigitalBox();
	} else if (ss.getActiveCell().getSheet().getName() === 'Battalion Structure' && ss.getActiveCell().getColumn() > 1) {
		if (ss.getActiveCell().getColumn() === 1 || ss.getActiveCell().getColumn() === 2) {
			updateFormGroups();
			checkForUniqueRolesAndGroups();
		}
		if (ss.getActiveCell().getColumn() > 1) {
			chainOfCommandStructureUpdater();
		}
	}
}

function chainOfCommandStructureUpdater() {
	if (ssBattalionStructure.getLastRow() > 1) {
		// Create list of all groups remaining
		let groups = [];
		let groupsCopy;
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
		groupsCopy = groups;
		// Read the chain to figure out what the structure is
		interface chain {
			value: string;
			children: chain[];
			pos: number[];
		}
		let chainOfCommand = { value: 'DropDownPlaceHolder12233', children: [], pos: [0, 0] } as chain;
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

		//Write json chainOfCOmmand to variables sheet // plz upload already
		ssVariables.getRange(3, 2).setValue(JSON.stringify(chainOfCommand));

		// Add dropdown children
		function recursiveDropDownChildAddition(chainNode: chain) {
			if (chainNode.children.length > 0) {
				chainNode.children.forEach((child) => {
					recursiveDropDownChildAddition(child);
				});
			}
			chainNode.children.push({ value: 'DropDownPlaceHolder12233', children: [], pos: [0, 0] });
		}
		if (chainOfCommand.value !== 'DropDownPlaceHolder12233') recursiveDropDownChildAddition(chainOfCommand);

		// Clear Data validations and normal values
		ssBattalionStructure
			.getRange(2, 3, ssBattalionStructure.getMaxRows() - 1, ssBattalionStructure.getMaxColumns() - 2)
			.clearContent()
			.clearDataValidations()
			.clearFormat()
			.clearNote();

		// Write value out array
		const outArr = [['']];
		let outArrCol = 0;
		function outArrCreator(chainNode: chain, outArrRow: number) {
			outArr[outArrRow][outArrCol] = chainNode.value;
			if (chainNode.children.length > 0) {
				outArr.push(Array(outArr[0].length).fill(''));
				outArrRow++;
				chainNode.children.forEach((child) => {
					outArrCreator(child, outArrRow);
				});
				outArr.forEach((row) => {
					row.push('');
				});
				outArrCol++;
			}
		}
		outArrCreator(chainOfCommand, 0);

		// Write dropdown menus
		const CoCArea = ssBattalionStructure.getRange(2, 3, outArr.length, outArr[0].length);
		groups = groupsCopy;
		let outDataValidations = CoCArea.getDataValidations();

		for (let i = 0; i < outArr.length; i++) {
			for (let j = 0; j < outArr[0].length; j++) {
				if (groups.indexOf(outArr[i][j]) > -1) {
					groups.splice(groups.indexOf(outArr[i][j]));
				} else if (outArr[i][j] === 'DropDownPlaceHolder12233') {
					outDataValidations[i][j] = SpreadsheetApp.newDataValidation()
						.setAllowInvalid(false)
						.requireValueInList(groups)
						.build();
					outArr[i][j] = '';
				}
			}
		}

		// Write values and data validation
		CoCArea.setValues(outArr);
		CoCArea.setDataValidations(outDataValidations);
	}
}

function checkForUniqueRolesAndGroups() {
	const data = ssBattalionStructure.getRange(2, 1, ssBattalionStructure.getLastRow(), 2).getValues();
	const alreadyRoles = [];
	const alreadyGroups = [];
	let somethingWasDeleted = false;
	data.forEach((row) => {
		if (row[0] !== '') {
			alreadyRoles.forEach((role) => {
				if (role === row[0]) {
					somethingWasDeleted = true;
					row[0] = '';
				}
			});
			alreadyRoles.push(row[0]);
		}
		if (row[1] !== '') {
			alreadyGroups.forEach((group) => {
				if (group === row[1]) {
					somethingWasDeleted = true;
					row[1] = '';
				}
			});
			alreadyGroups.push(row[1]);
		}
	});
	if (somethingWasDeleted) {
		ssBattalionStructure.getRange(2, 1, ssBattalionStructure.getLastRow(), 2).setValues(data);
	}
}

function createGoogleFiles() {
	const root = DriveApp.getFolderById('1vPucUC-lnMzCRWPZQ8FYkQHswNkB7Nv9');
	const battalionIndividuals = getGroups(true, false);
	var ssTemplate = SpreadsheetApp.openByUrl(
		'https://docs.google.com/spreadsheets/d/1QbC9z04dQWhDNz-Q4qm2urfWzZjtKxLmmCsFQ4J9uOU/edit#gid=0'
	);
	const templateID = ssTemplate.getId();
	const newFile = DriveApp.getFileById(templateID);
	for (var idx = 0; idx < battalionIndividuals.length; idx++) {
		const email = getIndividualEmail(battalionIndividuals[idx]);
		if (email === 'tnbowes@gatech.edu') {
			continue;
		}
		const indFile = newFile.makeCopy(battalionIndividuals[idx] + ', GT NROTC', root);
		const indID = indFile.getId();
		updateSheet(indID, battalionIndividuals[idx]);
		indFile.addViewer(email);
		Logger.log(email, battalionIndividuals[idx]);
		//indFile.addEditor('gtnrotc.ado@gmail.com');
	}
}

function findIndSheet(name) {
	var files = DriveApp.getFilesByName(name + ', GT NROTC');
	const fileList = [];
	while (files.hasNext()) {
		var sheet = files.next();
		fileList.push(sheet);
	}
	Logger.log(fileList);
	return files;
}

function wipeGoogleFiles() {
	const root = DriveApp.getFolderById('1vPucUC-lnMzCRWPZQ8FYkQHswNkB7Nv9');
	const battalionIndividuals = getGroups(true, false);
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

function updateSheet(sheetID, name) {
	const userSpread = SpreadsheetApp.openById(sheetID);
	const userPaperwork = userSpread.getSheetByName('Total_Paperwork');
	const header = userPaperwork.getRange(1, 1, 2, 3).getValues();
	const outData = userPaperwork.getRange(2, 3, userPaperwork.getLastRow(), userPaperwork.getLastColumn()).getValues();

	// Manipulate outData according to incomingData

	var chits = 0;
	var merits = 0;
	var negCounsel = 0;
	const pending = ssPending.getRange(1, 1, ssPending.getLastRow(), ssPending.getLastColumn()).getValues();
	for (var i = 1; i < pending.length; i++) {
		if (pending[i][3] === name) {
			Logger.log(pending[i]);
			if (pending[i][4] === 'Chit') {
				chits++;
			} else if (pending[i][4] === 'Negative Counseling') {
				negCounsel++;
			} else if (pending[i][4] === 'Merit') {
				merits++;
			}
			userPaperwork[userPaperwork.getLastRow()].setValues(pending[i]);
		}
	}
	const helpData = [];
	helpData.push(chits);
	helpData.push(negCounsel);
	helpData.push(merits);
	header[1][0] = chits;
	header[1][1] = negCounsel;
	header[1][2] = merits;
	userPaperwork.getRange(1, 1, 2, 3).setValues(header);
	Logger.log(header);

	//userPaperwork.getRange(1, 1, outData.length, outData[0].length).setValues(outData);
}

function updateFormGroups() {
	// Update assigner names list
	const FormItem = form.getItems();
	const item2 = FormItem[0].asListItem();
	const ind = getGroups(true, false);
	const indList = [];
	for (const individuals of ind) {
		indList.push(item2.createChoice(individuals));
	}
	item2.setChoices(indList);
	item2.isRequired();
	item2.setHelpText('Select your name from the list below.');

	// Update Recieve groups
	const item = FormItem[1].asCheckboxGridItem();
	item.setTitle('Receiving Groups/s');
	let roles = ssBattalionStructure.getRange(2, 1, ssBattalionStructure.getLastRow(), 1).getValues();
	const rowItems = [];
	roles.forEach((item) => {
		if (item[0] !== '') rowItems.push(item[0]);
	});
	const colItems = [];
	getGroups(false, true).forEach((group) => {
		colItems.push(group);
	});
	item.setRows(rowItems);
	item.setColumns(colItems);
	item.setHelpText(
		'Select the groups/s receiving the paperwork. This question has smart group selection and will assign the paperwork to anyone who qualifies for any of the groups selected.'
	);

	// Update Reciever individuals
	const item3 = FormItem[2].asCheckboxGridItem();
	item3.setTitle('Receiving Individual/s');
	const rowItems2 = [];
	getGroups(true, false).forEach((person) => {
		rowItems2.push(person);
	});
	const colItems2 = ['Individual'];
	item3.setRows(rowItems2);
	item3.setColumns(colItems2);
	item3.setHelpText('Select the individual/s receiving the paperwork.');

	// Reset form response
	if (ssAssignment.getLastColumn() > 150) {
		// 256 is max number of columns, I use 150 cuz why not
		const destID = form.getDestinationId();
		const destType = form.getDestinationType();
		form.removeDestination();
		form.deleteAllResponses();
		ss.deleteSheet(ssAssignment);
		form.setDestination(destType, destID);
	}

	// Find sheet and rename to assignmnet
	ss.getSheets().forEach((sheet) => {
		if (sheet.getName().substring(0, 14) === 'Form Responses') {
			sheet.setName('Assignment Responses');
			ssAssignment = sheet;
		}
	});

	//Update the form submission page
	const subFormItem = subForm.getItems();
	const subItem = subFormItem[0].asListItem();
	subItem.setTitle('Your name');
	const subInd = getGroups(true, false);
	const subIndList = [];
	for (const subIndividuals of subInd) {
		subIndList.push(item2.createChoice(subIndividuals));
	}
	subItem.setChoices(subIndList);
	subItem.isRequired();
	subItem.setHelpText('Select your name from the dropdown menu below');
}

function getGroups(individuals: boolean, groups: boolean): string[] {
	const groupData = ssBattalionStructure.getRange(2, 2, ssBattalionStructure.getLastRow(), 1).getValues();
	const individualData = ssBattalionMembers.getRange(2, 1, ssBattalionMembers.getLastRow(), 3).getValues();
	const out = [];

	if (groups) {
		for (let i = 0; i < groupData.length; i++) {
			const group = groupData[i][0];
			if (group !== '') {
				out.push(group);
			}
		}
	}

	if (individuals) {
		for (let i = 0; i < individualData.length; i++) {
			if (individualData[i][0] !== '' && individualData[i][1] !== '' && individualData[i][2] !== '') {
				const person = `MIDN ${individualData[i][0]}/C ${individualData[i][1]}, ${individualData[i][2]}`;
				out.push(person);
			}
		}
	}

	return out;
}

function getIndividualEmail(name: string): string {
	const individualData = ssBattalionMembers.getRange(2, 1, ssBattalionMembers.getLastRow(), 4).getValues();
	let returnEmail = null;
	for (let i = 0; i < individualData.length; i++) {
		const person = `MIDN ${individualData[i][0]}/C ${individualData[i][1]}, ${individualData[i][2]}`;
		if (name === person) {
			returnEmail = individualData[i][3];
		}
	}
	return returnEmail;
}

function createFullBattalionStructure() {
	const people = JSON.parse(ssVariables.getRange(4, 2).getValue());
	const chain = JSON.parse(ssVariables.getRange(3, 2).getValue());

	function fillChain(chainNode) {
		chainNode.members = [];
		people.forEach((person) => {
			if (person.group === chainNode.value) {
				chainNode.members.push(person);
			}
		});
		if (chainNode.children.length > 0) {
			chainNode.children.forEach((child) => {
				child.parent = chainNode;
				fillChain(child);
			});
		}
	}
	chain.parent = null;
	fillChain(chain);
	return chain;
}

// Still working on this function
function getIndividualsFromCheckBoxGrid(parsedCheckBoxData, assigner) {
	Logger.log(JSON.stringify(parsedCheckBoxData) + ' ' + JSON.stringify(assigner));

	let outList = [] as { name: string; group: string; canBeAssignedFromAssigner: boolean }[];
	const battalion = createFullBattalionStructure();
	const battaionMembers = JSON.parse(ssVariables.getRange(4, 2).getValue());

	parsedCheckBoxData.forEach((node) => {
		// node: role: role or individual, groups: [Individual or groups]
		let isIndividual = false;
		battaionMembers.forEach((member) => {
			if (node.role === member.name) {
				isIndividual = true;
				outList.push({ name: member.name, group: 'Individual', canBeAssignedFromAssigner: true });
			}
		});
		if (!isIndividual) {
			node.groups.forEach((selectedGroup) => {
				if (selectedGroup !== 'Individual') {
					//Find the group in the chain
					let groupChain;
					function findGroup(chainNode) {
						if (chainNode.value === selectedGroup) {
							groupChain = chainNode;
						}
						chainNode.children.forEach((child) => {
							findGroup(child);
						});
					}
					findGroup(battalion);

					//Search down chain for role
					function addPeopleDownChain(chainNode) {
						chainNode.members.forEach((member) => {
							if (member.role === node.role) {
								outList.push({
									name: member.name,
									group: selectedGroup + ':' + node.role,
									canBeAssignedFromAssigner: true,
								});
							}
						});
						chainNode.children.forEach((child) => {
							addPeopleDownChain(child);
						});
					}
					addPeopleDownChain(groupChain);

					//Search up chain for role
					function addPeopleUpChain(chainNode) {
						if (chainNode.parent !== null) {
							chainNode.parent.members.forEach((member) => {
								if (member.role === node.role) {
									outList.push({
										name: member.name,
										group: selectedGroup + ':' + node.role,
										canBeAssignedFromAssigner: true,
									});
								}
							});
							addPeopleUpChain(chainNode.parent);
						}
					}
					addPeopleUpChain(groupChain);
				}
			});
		}
	});

	// check for duplicates and remove self
	const outListWithoutRepeats = [];
	outList.forEach((outPerson) => {
		let alreadyInArray = false;
		outListWithoutRepeats.forEach((withoutRepeats) => {
			if (withoutRepeats.name === outPerson.name) {
				alreadyInArray = true;
			}
		});
		if (!alreadyInArray && outPerson.name !== assigner.name) {
			outListWithoutRepeats.push(outPerson);
		}
	});
	outList = outListWithoutRepeats;

	// Check for assigning autority
	const canAssignToAnyone = ssOptions.getRange(4, 2).getValue();
	if (canAssignToAnyone !== 'Disabled' && canAssignToAnyone !== '') {
		const rolesList = [];
		ssBattalionStructure
			.getRange(2, 1, ssBattalionStructure.getLastRow(), 1)
			.getValues()
			.forEach((row) => {
				if (row[0] !== '') {
					rolesList.push(row[0]);
				}
			});

		//IF the assigner cannot assign to anyone check to make sure everyone assigning to is corrects
		if (
			rolesList.indexOf(canAssignToAnyone) < rolesList.indexOf(assigner.role) &&
			rolesList.indexOf(canAssignToAnyone) !== -1
		) {
			const subordinates = getSubordinates(assigner.name);
			outList.forEach((outPerson) => {
				let isSubordinate = false;
				subordinates.forEach((suboord) => {
					if (outPerson.name === suboord) {
						isSubordinate = true;
					}
				});
				if (!isSubordinate) {
					outPerson.canBeAssignedFromAssigner = false;
				}
			});
		}
	}

	Logger.log(outList);
	return outList; // [{name:string,group:string}]
}

function updateBattalionMembersJSON() {
	const data = ssBattalionMembers
		.getRange(2, 1, ssBattalionMembers.getLastRow(), ssBattalionMembers.getLastColumn())
		.getValues();
	const peopleList = [];
	if (data[0].length === 6) {
		data.forEach((row) => {
			if (row[0] !== '' && row[1] !== '' && row[2] !== '' && row[3] !== '' && row[4] !== '' && row[5] !== '') {
				peopleList.push({ name: `MIDN ${row[0]}/C ${row[1]}, ${row[2]}`, email: row[3], role: row[4], group: row[5] });
			}
		});
	}
	ssVariables.getRange(4, 2).setValue(JSON.stringify(peopleList));
}

function getSubordinates(name: string): string[] {
	const outPeople = [];
	const battalion = createFullBattalionStructure();
	const rolesList = [];
	ssBattalionStructure
		.getRange(2, 1, ssBattalionStructure.getLastRow(), 1)
		.getValues()
		.forEach((row) => {
			if (row[0] !== '') {
				rolesList.push(row[0]);
			}
		});
	let highestChainOfIndividual;
	let personFullData;
	let foundPerson = false;

	function searchChain(chainNode) {
		chainNode.members.forEach((member) => {
			if (member.name === name) {
				highestChainOfIndividual = chainNode;
				personFullData = member;
				foundPerson = true;
			}
		});
		chainNode.children.forEach((child) => {
			searchChain(child);
		});
	}
	searchChain(battalion);

	if (foundPerson) {
		highestChainOfIndividual.members.forEach((member) => {
			if (rolesList.indexOf(member.role) > rolesList.indexOf(personFullData.role)) {
				outPeople.push(member.name);
			}
		});
		function addSubGroupMembers(chainNode) {
			chainNode.members.forEach((member) => {
				outPeople.push(member.name);
			});
			chainNode.children.forEach((child) => {
				addSubGroupMembers(child);
			});
		}
		highestChainOfIndividual.children.forEach((child) => {
			addSubGroupMembers(child);
		});
	}

	return outPeople;
}

function sendEmail(emailList, data) {
	const emailsActivated = ssOptions.getRange(1, 2).getValue().toString().toLowerCase() === 'true';
	if (!emailsActivated) return;

	const dateDemo = data.dueDate.toString().split(' ', 4);

	const date = dateDemo[0] + ', ' + dateDemo[2] + dateDemo[1].toUpperCase() + dateDemo[3];

	const emailSender = getIndividualEmail(data.assigner);

	const emailSubject = 'NROTC ADMIN Department: New ' + data.paperwork + ' due COB ' + date + '.';

	const emailBody =
		"<h2 'style=color: #5e9ca0;'> You have been assigned a " +
		data.paperwork +
		' from ' +
		data.assigner +
		'.</h2>' +
		'<p> The reason is the following: ' +
		data.reason +
		'.</p> <p> You must turn this form in by COB on ' +
		date +
		'.</p>' +
		'<p> If you have any questions regarding the validity of the ' +
		data.paperwork +
		', please contact the assignee. </p>' +
		'<p> You can find the paperwork to complete here: ' +
		data.pdfLink +
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
