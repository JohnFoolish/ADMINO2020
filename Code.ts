/* To Do
Link sheets / names by semester
Sort data sheet in reverse order by date
Write out proper documentation/docstrings

*/

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
const ssPendingCache = ss.getSheetByName('PendingChangedCache');

const form = FormApp.openByUrl('https://docs.google.com/forms/d/1l6lZZhsOWb5rcyTFDxyiFJln0tFBuVIiFRGK_hjnZ84/edit');
const subForm = FormApp.openByUrl('https://docs.google.com/forms/d/1x2HP45ygThm6MoYlKasVnaacgZUW_yKA7Cz9pxKKOJc/edit');
const root = DriveApp.getFolderById('1vPucUC-lnMzCRWPZQ8FYkQHswNkB7Nv9');
const ssTemplate = SpreadsheetApp.openByUrl(
	'https://docs.google.com/spreadsheets/d/1QbC9z04dQWhDNz-Q4qm2urfWzZjtKxLmmCsFQ4J9uOU/edit#gid=0'
);
/**
 *
 */
function myOnOpen(e) {
	var ui = SpreadsheetApp.getUi();
	ui.createMenu('DB Functions')
		.addItem('Initialize', 'initForSemester')
		.addSeparator()
		.addItem('Clear Pending Cache', 'updateSheetsFromPendingCache')
		.addToUi();
}

/**
 *
 */
function initForSemester() {
	ssVariables.getRange(6, 2).setValue('true');
	ssOptions.getRange(6, 2).setValue('true');
	ssVariables.getRange(8, 2).setValue('true');
}

/**
 *
 */
function initSheetReminder() {
	// Reset reminders and disable sheet if semester end has been reached
	const now = new Date();
	if (ssVariables.getRange(8, 2).getValue().toString() == 'true') {
		const resetSheet = function () {
			wipeGoogleFiles();
			ssVariables.getRange(8, 2).setValue('false');
			ssOptions.getRange(6, 2).setValue('false');
		};

		// Reset for fall
		if (now.getMonth() === 7) {
			resetSheet();
		}

		// Reset for spring
		if (now.getMonth() === 11 && now.getDate() > 20) {
			resetSheet();
		}
	}

	// Send reminder if not disabled
	if (ssOptions.getRange(6, 2).getValue().toString() == 'false') {
		sendInitReminderEmail();
	}
}

/**
 *
 */
//Triggers when the submission form is submitted
function myOnSubmit() {
	if (ssVariables.getRange(8, 2).getValue().toString() == 'true') {
		if (ssVariables.getRange(1, 2).getValue().toString() !== ssAssignment.getLastRow().toString()) {
			myOnAssignmentSubmit();
			ssVariables.getRange(1, 2).setValue(ssAssignment.getLastRow());
		}
		if (ssVariables.getRange(2, 2).getValue().toString() !== ssTurnedIn.getLastRow().toString()) {
			myOnFormTurnedInSubmit();
			ssVariables.getRange(2, 2).setValue(ssTurnedIn.getLastRow());
		}
	} else {
		let submitterName = '';
		if (ssVariables.getRange(1, 2).getValue().toString() !== ssAssignment.getLastRow().toString()) {
			submitterName = ssAssignment.getRange(ssAssignment.getLastRow(), 2).getValue();
			ssVariables.getRange(1, 2).setValue(ssAssignment.getLastRow());
		}
		if (ssVariables.getRange(2, 2).getValue().toString() !== ssTurnedIn.getLastRow().toString()) {
			submitterName = ssTurnedIn.getRange(ssTurnedIn.getLastRow(), 2).getValue();
			ssVariables.getRange(2, 2).setValue(ssTurnedIn.getLastRow());
		}
		sendSheetNotEnabledEmail(submitterName);
	}
}

/**
 *
 */
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
			} else if (dataPairs[0][i] === 'Send Assignment Email Notification') {
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

		const people = getIndividualsFromCheckBoxGrid(keyValuePairsRawGridCheckbox, assignerFullData);

		// Check to make sure inputs are valid
		if (submitData.dateDue.getFullYear() !== 1945 && people.length !== 0) {
			// Manipulate data
			const outData = [];
			const emailNameList = [];
			const noAuthority = [];
			const Authority = [];
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
						emailNameList.push(people[i].name);
					}
					Authority.push(people[i].name);
				} else {
					noAuthority.push(people[i].name);
				}
			}
			//Send ouot email notifiying everyone that their paperwork was assigned
			sendAssigneesEmail(emailNameList, submitData);
			//Email the assigner who was assigned it and who was not
			sendAssignerSuccessEmail(assignerFullData, submitData, noAuthority, Authority);

			//Write to data sheet
			ssData.getRange(ssData.getLastRow() + 1, 1, outData.length, outData[0].length).setValues(outData);
			ssData.getRange(2, 1, ssData.getLastRow() - 1, ssData.getLastColumn()).sort({ column: 1, ascending: false });

			// Write to Pending Paperwork
			ssPending.getRange(ssPending.getLastRow() + 1, 1, outData.length, outData[0].length).setValues(outData);
			ssPending.getRange(2, 1, ssPendingCache.getLastRow() - 1, ssPendingCache.getLastColumn()).sort(7);

			outData.forEach((row) => {
				dynamicSheetUpdate(row);
			});
		} else {
			sendAssignerFailEmail(
				assignerFullData,
				submitData,
				submitData.dateDue.getFullYear() === 1945,
				people.length === 0
			);
		}
	}
}

/**
 *
 */
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
		const handleEmpty = ssOptions.getRange(5, 2).getValue();
		if (handleEmpty === 'Reject Submission') {
			out = new Date();
			out.setFullYear('1945');
		} else {
			out = new Date(assignDate.toString());
			out.setDate(out.getDate() + parseInt(handleEmpty));
		}
	}
	return out;
}

/**
 *
 */
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

/**
 *
 */
//This function runs whenever the new paperwork submission form is submitted.
function myOnFormTurnedInSubmit() {
	if (ssTurnedIn.getLastRow() > 1) {
		// Get data from linked sheet to use
		const data = ssTurnedIn.getRange(ssTurnedIn.getLastRow(), 1, 1, ssTurnedIn.getLastColumn()).getValues();

		// Manipulate Data / Rearrange Data
		const outData = data;
		outData[0].push('FALSE');

		updateTurnedInPaperworkTab(outData[0]);

		// Write to Digital Admin box sheet
		ssDigitalBox.getRange(ssDigitalBox.getLastRow() + 1, 1, outData.length, outData[0].length).setValues(outData);
		sortDigitalBox();
	}
}

/**
 *
 */
function sortDigitalBox() {
	if (ssDigitalBox.getLastRow() > 1) {
		ssDigitalBox.getRange(2, 1, ssDigitalBox.getLastRow() - 1, ssDigitalBox.getLastColumn()).sort(1);
		ssDigitalBox.getRange(2, 1, ssDigitalBox.getLastRow() - 1, ssDigitalBox.getLastColumn()).sort(4);
	}
}

/**
 *
 */
function myOnEdit() {
	if (ss.getActiveCell().getSheet().getName() === 'Battalion Members') {
		updateBattalionMembersJSON();
	} else if (ss.getActiveCell().getSheet().getName() === 'Pending Paperwork' && ss.getActiveCell().getColumn() === 8) {
		const pending = ssPending.getRange(1, 1, ssPending.getLastRow(), ssPending.getLastColumn()).getValues();
		let oneWasTrue = false;
		let alertUserToAddContent = false;
		for (let j = 1; j < pending.length; j++) {
			if (pending[j][7].toString() !== 'Pending' && pending[j][7].toString() !== '') {
				if (pending[j][9] === '' && pending[j][7] === 'Approved') {
					alertUserToAddContent = true;
					pending[j][7] = 'Pending';
				} else {
					oneWasTrue = true;
					ssPendingCache.getRange(ssPendingCache.getLastRow() + 1, 1, 1, pending[j].length).setValues([pending[j]]);
					pending[j] = pending[j].map((item) => '');
				}
			}
		}
		if (oneWasTrue || alertUserToAddContent) {
			ssPending.getRange(1, 1, pending.length, pending[0].length).setValues(pending);
			ssPending.sort(1);
		}
		if (alertUserToAddContent) {
			const ui = SpreadsheetApp.getUi();
			ui.alert('You need to put either "Turned in Physically" or the link to their digitally turned in file');
		}
	} else if (
		ss.getActiveCell().getSheet().getName() === 'Digital Turn In Box' &&
		ss.getActiveCell().getColumn() === 4
	) {
		sortDigitalBox();
		updateTurnedInPaperworkTab(
			ssDigitalBox.getRange(ss.getActiveCell().getRow(), 1, 1, ssDigitalBox.getLastColumn()).getValues()[0]
		);
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

/**
 *
 */
function updateSheetsFromPendingCache() {
	if (ssPendingCache.getLastRow() > 0) {
		const data = ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).getValues();
		const pending = ssPendingCache.getRange(1, 1, 1, ssPendingCache.getLastColumn()).getValues();

		const uuidDate = pending[0][0].toString();
		for (let i = 0; i < data.length; i++) {
			if (data[i][0].toString() === uuidDate) {
				data[i][7] = pending[0][7];
				data[i][9] = pending[0][9];
			}
		}
		ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).setValues(data);
		dynamicSheetUpdate(pending[0]);

		pending[0] = pending[0].map((item) => '');
		ssPendingCache.getRange(1, 1, 1, pending[0].length).setValues(pending);
		ssPendingCache.sort(1);
		Logger.log('Successfully updated: ', pending[0]);
		updateSheetsFromPendingCache();
	}
}

/**
 *
 */
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

/**
 *
 */
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

/**
 *
 */
function autoRunCreateGoogleFiles() {
	if (ssVariables.getRange(6, 2).getValue().toString().toLowerCase() === 'true') {
		createGoogleFiles();
	}
}

/**
 *
 */
function createGoogleFiles() {
	if (ssVariables.getRange(8, 2).getValue().toString().toLowerCase() == 'false') {
		return;
	}

	const battalionIndividuals = getGroups(true, false);

	const templateID = ssTemplate.getId();
	const newFile = DriveApp.getFileById(templateID);
	for (var idx = 0; idx < battalionIndividuals.length; idx++) {
		if (root.getFilesByName(battalionIndividuals[idx] + ', GT NROTC').hasNext()) {
			Logger.log('The form for user ' + battalionIndividuals[idx] + ' exists.');
			continue;
		}
		const email = getIndividualEmail(battalionIndividuals[idx]);
		const indFile = newFile.makeCopy(battalionIndividuals[idx] + ', GT NROTC', root);
		const indID = indFile.getId();
		initSheet(indID, battalionIndividuals[idx]);
		indFile.addViewer(email);
		Logger.log(email, battalionIndividuals[idx]);

		// For editting the sheet in the testing version
		if (Session.getEffectiveUser().getEmail() !== 'gtnrotc.ado@gmail.com') {
			indFile.addEditor('johnlcorker88@gmail.com');
			indFile.addEditor('tnbowes@gmail.com');
		}

		ssVariables.getRange(7, 2).setValue(ssVariables.getRange(7, 2).getValue() + '|' + battalionIndividuals[idx] + '|');
	}
	let finishInitMem = ssVariables
		.getRange(7, 2)
		.getValue()
		.split('|')
		.filter((a) => {
			return a === '' ? false : true;
		});
	for (let i = finishInitMem.length - 1; i >= 0; i--) {
		updateSubordinateTab(finishInitMem[i]);
		finishInitMem.splice(i, 1);
		ssVariables.getRange(7, 2).setValue(finishInitMem.join('|'));
	}
	if (finishInitMem.length === 0) {
		ssVariables.getRange(6, 2).setValue('false');
		MailApp.sendEmail({
			to: Session.getEffectiveUser().getEmail(),
			subject: 'The Paperwork Database was Successfully Enabled',
			htmlBody: `ADMINO,<br><br>The paperwork database was successfully initialized. Have a good semester!<br><br>Very respectfully,<br>The ADMIN Department`,
		});
	}
}

/**
 *
 */
function updateAllSubordinates() {
	const finishInitMem = getGroups(true, false);
	for (let i = finishInitMem.length - 1; i >= 0; i--) {
		updateSubordinateTab(finishInitMem[i]);
		finishInitMem.splice(i, 1);
		ssVariables.getRange(7, 2).setValue(finishInitMem.join('|'));
	}
}

/**
 *
 */
function findIndSheet(name) {
	var files = root.getFilesByName(name + ', GT NROTC');
	const fileList = [];
	while (files.hasNext()) {
		var sheet = files.next();
		fileList.push(sheet);
	}
	Logger.log('Name: ', name);
	Logger.log('The current sheets returned: ', fileList);
	const tup = [files, fileList];
	return tup;
}

/**
 *
 */
function wipeGoogleFiles() {
	const battalionIndividuals = getGroups(true, false);
	for (var ind = 0; ind < battalionIndividuals.length; ind++) {
		const email = getIndividualEmail(battalionIndividuals[ind]);
		if (email === '') {
			continue;
		}
		const [fileIterator, _] = findIndSheet(battalionIndividuals[ind]);
		const fileLinkedList = fileIterator as GoogleAppsScript.Drive.FileIterator;
		while (fileLinkedList.hasNext()) {
			var file = fileLinkedList.next();
			root.removeFile(file);
		}
	}
}
/**
 *
 */
function updateTurnedInPaperworkTab(tempData) {
	let [fileIterator, fileList] = findIndSheet(tempData[1]);
	let fileArray = fileList as Array<GoogleAppsScript.Drive.File>;
	let fileLinkedList = fileIterator as GoogleAppsScript.Drive.FileIterator;

	if (fileArray.length > 1) {
		Logger.log('Error, multiple sheets for ' + tempData[1]);
	} else if (fileArray.length == 0) {
		createGoogleFiles();
		Logger.log('Attempted to created google file for ', tempData[1]);
		[fileIterator, fileList] = findIndSheet(tempData[1]);
		fileArray = fileList as Array<GoogleAppsScript.Drive.File>;
		fileLinkedList = fileIterator as GoogleAppsScript.Drive.FileIterator;
	}
	const name = tempData[1];
	fileArray.forEach((file) => {
		const userSpread = SpreadsheetApp.open(file);

		const userPaperwork = userSpread.getSheetByName('Submitted Paperwork');
		const outData = userPaperwork.getRange(1, 1, userPaperwork.getLastRow(), userPaperwork.getLastColumn()).getValues();
		var lineAddition = userPaperwork.getLastRow() + 1;

		for (var i = 1; i < outData.length; i++) {
			if (tempData[0].toString() === outData[i][0].toString()) {
				//Duplicate file found!
				lineAddition = i + 1;
			}
		}
		Logger.log(lineAddition);
		Logger.log(tempData);

		userPaperwork.getRange(lineAddition, 1, 1, tempData.length).setValues([tempData]);

		if (userPaperwork.getLastRow() > 6) {
			userPaperwork.getRange(6, 1, userPaperwork.getLastRow() - 6, userPaperwork.getLastColumn()).sort(1);
		}
	});
}

/**
 *
 */
function updateSubordinateTab(name) {
	//go and get the data from each of the subordinate's paperwork sheets
	let [fileIterator, fileList] = findIndSheet(name);
	let fileArray = fileList as Array<GoogleAppsScript.Drive.File>;
	let fileLinkedList = fileIterator as GoogleAppsScript.Drive.FileIterator;
	if (fileArray.length > 1) {
		Logger.log('Error, multiple sheets for ' + name);
	} else if (fileArray.length == 0) {
		createGoogleFiles();
		Logger.log('Attempted to created google file for ', name);
		[fileIterator, fileList] = findIndSheet(name);
		fileArray = fileList as Array<GoogleAppsScript.Drive.File>;
		fileLinkedList = fileIterator as GoogleAppsScript.Drive.FileIterator;
	}
	fileArray.forEach((file) => {
		const userSpread = SpreadsheetApp.open(file);
		const subPaperwork = userSpread.getSheetByName('Subordinate Paperwork');

		const subList = descendingRankOrderOfSubordinateNames(name);
		var subordinateData = [];
		var blankLine;
		var indData;
		// here get each of the subordinates data arrays
		//for each something goes here fda

		var dict = {};
		let dataList = [];
		subList.forEach((subName) => {
			dict[subName] = { Merit: 0, Chit: 0, 'Negative Counseling': 0, Data: JSON.parse(JSON.stringify(dataList)) };
		});
		subordinateData = grabUsersData(dict);
		Logger.log('total subordinate data is: ', subordinateData.length);
		subPaperwork.getRange(2, 1, subPaperwork.getLastRow(), subPaperwork.getLastColumn()).clearContent();
		if (subordinateData.length > 0) {
			Logger.log(subordinateData[0]);
			subPaperwork.getRange(2, 1, subordinateData.length, subordinateData[0].length).setValues(subordinateData);
		}
	});
}

/**
 *
 */
function grabUsersData(dict) {
	Logger.log('Entering grabUsersData!');
	//Logger.log('dictionary is: ', dict);
	var finalSubData = [];

	const database = ssData.getRange(2, 1, ssData.getLastRow(), ssData.getLastColumn()).getValues();
	for (var idx = 0; idx < database.length; idx++) {
		if (database[idx][3] in dict) {
			if (database[idx][7] !== 'Cancelled' && database[idx][7] !== 'Rejected') {
				dict[database[idx][3]][database[idx][4]] += 1;
			}
			if (database[idx][7] !== 'Cancelled') {
				dict[database[idx][3]]['Data'].push(database[idx]);
			}
		}
	}

	//Now we should have a fully populated dictionary of values

	//Format them into a nice little array to look at and write!
	var indData;
	for (const key in dict) {
		indData = getFullMemberData(key);
		finalSubData.push([
			'Name:',
			key,
			'Rank:',
			Object.freeze(indData.role) + ', ' + Object.freeze(indData.group),
			'',
			'',
			'',
			'',
			'',
			'',
		]);
		finalSubData.push([
			'Chits: ' + dict[key]['Chit'].toString(),
			'Negative Counselings: ' + dict[key]['Negative Counseling'].toString(),
			'Merits: ' + dict[key]['Merit'].toString(),
			'',
			'',
			'',
			'',
			'',
			'',
			'',
		]);
		dict[key]['Data'].forEach((assignment) => {
			finalSubData.push(assignment);
		});
		finalSubData.push(['', '', '', '', '', '', '', '', '', '']);
		finalSubData.push(['', '', '', '', '', '', '', '', '', '']);
	}
	return finalSubData;
}

/**
 *
 */
function dynamicSheetUpdate(tempData) {
	let [fileIterator, fileList] = findIndSheet(tempData[3]);
	let fileArray = fileList as Array<GoogleAppsScript.Drive.File>;
	let fileLinkedList = fileIterator as GoogleAppsScript.Drive.FileIterator;

	if (fileArray.length > 1) {
		Logger.log('Error, multiple sheets for ' + tempData[3]);
	} else if (fileArray.length == 0) {
		createGoogleFiles();
		Logger.log('Attempted to created google file for ', tempData[3]);
		[fileIterator, fileList] = findIndSheet(tempData[3]);
		fileArray = fileList as Array<GoogleAppsScript.Drive.File>;
		fileLinkedList = fileIterator as GoogleAppsScript.Drive.FileIterator;
	}
	const name = tempData[3];
	fileArray.forEach((file) => {
		const userSpread = SpreadsheetApp.open(file);

		const userPaperwork = userSpread.getSheetByName('Total Paperwork');
		const totalPaperwork = userSpread.getSheetByName('All Semesters');
		const header = userPaperwork.getRange(1, 1, 3, 3).getValues();
		const outData = userPaperwork.getRange(1, 1, userPaperwork.getLastRow(), userPaperwork.getLastColumn()).getValues();
		const totalOutData = totalPaperwork
			.getRange(1, 1, totalPaperwork.getLastRow(), totalPaperwork.getLastColumn())
			.getValues();

		var chits = header[2][0];
		var merits = header[2][2];
		var negCounsel = header[2][1];
		var lineAddition = userPaperwork.getLastRow() + 1;
		var totalLineAddition = totalPaperwork.getLastRow() + 1;
		var found = false;

		for (var i = 5; i < outData.length; i++) {
			if (tempData[0].toString() === outData[i][0].toString()) {
				//Duplicate file found!
				lineAddition = i + 1;
				for (var j = 5; j < totalOutData.length; j++) {
					if (tempData[0].toString() === totalOutData[j][0].toString()) {
						totalLineAddition = j + 1;
						found = true;
					}
				}
				if (!found) {
					Logger.log("Did not find the data in the total paperwork, but it was present in this semester's... Error");
					throw Error;
				}
			}
		}
		Logger.log(lineAddition);
		Logger.log(tempData);
		var change = 1;
		if (tempData[7] === 'Cancelled' || tempData[7] === 'Rejected') {
			change = -1;
		} else if (tempData[7] === 'Approved') {
			change = 0;
		}
		if (tempData[4] === 'Chit') {
			chits += change;
		} else if (tempData[4] === 'Negative Counseling') {
			negCounsel += change;
		} else if (tempData[4] === 'Merit') {
			merits += change;
		}
		if (merits < 0) {
			merits = 0;
		} else if (negCounsel < 0) {
			negCounsel = 0;
		} else if (chits < 0) {
			chits = 0;
		}

		if (tempData[7] === 'Cancelled') {
			tempData = ['', '', '', '', '', '', '', '', ''];
		}
		userPaperwork.getRange(lineAddition, 1, 1, tempData.length).setValues([tempData]);
		totalPaperwork.getRange(totalLineAddition, 1, 1, tempData.length).setValues([tempData]);

		if (userPaperwork.getLastRow() > 6) {
			userPaperwork.getRange(6, 1, userPaperwork.getLastRow() - 6, userPaperwork.getLastColumn()).sort(1);
		}
		if (totalPaperwork.getLastRow() > 6) {
			totalPaperwork.getRange(6, 1, totalPaperwork.getLastRow() - 6, totalPaperwork.getLastColumn()).sort(1);
		}

		const helpData = [];
		helpData.push(chits);
		helpData.push(negCounsel);
		helpData.push(merits);
		header[2][0] = chits;
		header[2][1] = negCounsel;
		header[2][2] = merits;
		userPaperwork.getRange(1, 1, 3, 3).setValues(header);
		totalPaperwork.getRange(1, 1, 3, 3).setValues(header);
		Logger.log(header);
	});
	const superiorList = getSuperiors(name);
	superiorList.forEach((superior) => {
		updateSubordinateTab(superior);
	});
}

/**
 *
 */
function initSheet(sheetID, name) {
	const userSpread = SpreadsheetApp.openById(sheetID);
	const userPaperwork = userSpread.getSheetByName('Total Paperwork');
	const totalPaperwork = userSpread.getSheetByName('All Semesters');
	const header = userPaperwork.getRange(1, 1, 3, 3).getValues();
	const outData = userPaperwork.getRange(3, 3, userPaperwork.getLastRow(), userPaperwork.getLastColumn()).getValues();

	// Manipulate outData according to incomingData

	var chits = 0;
	var merits = 0;
	var negCounsel = 0;
	const data = ssData.getRange(1, 1, ssData.getLastRow(), ssData.getLastColumn()).getValues();
	for (var i = 1; i < data.length; i++) {
		if (data[i][3] === name) {
			Logger.log(data[i]);
			if (data[i][4] === 'Chit' && data[i][7] !== 'Cancelled' && data[i][7] !== 'Rejected') {
				chits++;
			} else if (data[i][4] === 'Negative Counseling' && data[i][7] !== 'Cancelled' && data[i][7] !== 'Rejected') {
				negCounsel++;
			} else if (data[i][4] === 'Merit' && data[i][7] !== 'Cancelled' && data[i][7] !== 'Rejected') {
				merits++;
			}
			if (data[i][7] !== 'Cancelled') {
				userPaperwork.getRange(userPaperwork.getLastRow() + 1, 1, 1, data[i].length).setValues([data[i]]);
				totalPaperwork.getRange(totalPaperwork.getLastRow() + 1, 1, 1, data[i].length).setValues([data[i]]);
			}
		}
	}
	const helpData = [];
	helpData.push(chits);
	helpData.push(negCounsel);
	helpData.push(merits);
	header[0][1] = name;
	header[2][0] = chits;
	header[2][1] = negCounsel;
	header[2][2] = merits;
	userPaperwork.getRange(1, 1, 3, 3).setValues(header);
	totalPaperwork.getRange(1, 1, 3, 3).setValues(header);
	Logger.log(header);
}

/**
 *
 */
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
		'Select any group/s that you would like to assign paperwork. Think: the _____(row)/s of the _______(column).'
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
	item3.setHelpText('Select the individual/s receiving the paperwork not already selected by group selection.');

	// Reset form response
	if (ssAssignment.getLastColumn() > 150) {
		// 256 is max number of columns, I use 150 cuz why not
		const destID = form.getDestinationId();
		const destType = form.getDestinationType();
		form.removeDestination();
		form.deleteAllResponses();
		ss.deleteSheet(ssAssignment);
		form.setDestination(destType, destID);

		// Find sheet and rename to assignmnet
		ss.getSheets().forEach((sheet) => {
			if (sheet.getName().substring(0, 14) === 'Form Responses') {
				sheet.setName('Assignment Responses');
				ssAssignment = sheet;
				ssAssignment.hideSheet();
			}
		});
		ssVariables.getRange(1, 2).setValue('1');
	}

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
	createGoogleFiles();
}

/**
 *
 */
function getGroups(individuals: boolean, groups: boolean): string[] {
	const groupData = ssBattalionStructure.getRange(2, 2, ssBattalionStructure.getLastRow(), 1).getValues();
	const individualData = JSON.parse(ssVariables.getRange(4, 2).getValue());
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
			out.push(individualData[i].name);
		}
	}

	return out;
}

/**
 *
 */
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

/**
 *
 */
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

/**
 *
 */
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
	if (canAssignToAnyone !== 'Anyone can assign to anyone' && canAssignToAnyone !== '') {
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

/**
 *
 */
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
	const oldMembers = JSON.parse(ssVariables.getRange(4, 2).getValue().toString());
	ssVariables.getRange(4, 2).setValue(JSON.stringify(peopleList));

	if (peopleList.length !== oldMembers.length) {
		updateFormGroups();
	} else {
		const dictOfPeople = {};
		function addToDict(name) {
			if (dictOfPeople[name] === undefined) {
				dictOfPeople[name] = 0;
			} else {
				dictOfPeople[name]++;
			}
		}
		oldMembers.forEach((member) => {
			addToDict(member.name);
		});
		peopleList.forEach((member) => {
			addToDict(member.name);
		});

		let runUpdateFormGroups = false;
		for (const name in dictOfPeople) {
			if (dictOfPeople[name] === 0) {
				runUpdateFormGroups = true;
			}
		}
		if (runUpdateFormGroups) {
			updateFormGroups();
		}
	}
}

/**
 *
 */
function getSuperiors(name: string): string[] {
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
			if (rolesList.indexOf(member.role) < rolesList.indexOf(personFullData.role)) {
				outPeople.push(member.name);
			}
		});
		function getMemebersFromChainAscending(chainNode) {
			chainNode.members.forEach((member) => {
				outPeople.push(member.name);
			});
			if (chainNode.parent !== null) {
				getMemebersFromChainAscending(chainNode.parent);
			}
		}
		if (highestChainOfIndividual.parent !== null) {
			getMemebersFromChainAscending(highestChainOfIndividual.parent);
		}
	}

	return outPeople;
}

/**
 *
 */
function descendingRankOrderOfSubordinateNames(name: string): string[] {
	const subordinates = getSubordinates(name);
	const fullSub = [];
	const rolesList = [];
	ssBattalionStructure
		.getRange(2, 1, ssBattalionStructure.getLastRow(), 1)
		.getValues()
		.forEach((row) => {
			if (row[0] !== '') {
				rolesList.push(row[0]);
			}
		});
	subordinates.forEach((suboord) => {
		fullSub.push(getFullMemberData(suboord));
	});

	fullSub.sort((a, b) => {
		return rolesList.indexOf(a.role) - rolesList.indexOf(b.role);
	});

	let outPeople = [];
	fullSub.forEach((sub) => outPeople.push(sub.name));

	Logger.log(outPeople);
	return outPeople;
}

/**
 *
 */
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

/**
 *
 */
function sendAssignerFailEmail(assigner, submitData, noDate: boolean, noPeople: boolean) {
	/*const emailsActivated = ssOptions.getRange(1, 2).getValue().toString().toLowerCase() === 'true';
	if (!emailsActivated) return;*/
	let emailBody = `${assigner.name},
	<br><br>
	Your ${submitData.paperwork} did not assign, because ${noDate ? 'you did not give a due date' : ''}${
		noDate && noPeople ? ' and ' : ''
	}${noPeople ? 'you did not select to assign it to anyone' : ''}.
	<br><br>
	Very respectfully,<br>
	The ADMIN Department`;

	MailApp.sendEmail({
		to: assigner.email,
		subject: `Failed to assign ${submitData.paperwork}`,
		htmlBody: emailBody,
	});
}

/**
 *
 */
function sendAssignerSuccessEmail(
	assignerData: { name: string; email: string; role: string; group: string },
	submitData,
	noAuthority: string[],
	authority: string[]
) {
	/*const emailsActivated = ssOptions.getRange(1, 2).getValue().toString().toLowerCase() === 'true';
	if (!emailsActivated) return;*/
	const namesToEmailFormat = function (names: string[]): string {
		const classSeperatedNames = [[], [], [], []];
		names.forEach((name) => {
			if (name.substring(5, 6) === '1') {
				classSeperatedNames[0].push(name.substring(9));
			} else if (name.substring(5, 6) === '2') {
				classSeperatedNames[1].push(name.substring(9));
			} else if (name.substring(5, 6) === '3') {
				classSeperatedNames[2].push(name.substring(9));
			} else if (name.substring(5, 6) === '4') {
				classSeperatedNames[3].push(name.substring(9));
			}
		});
		let out = '<ul>';
		Logger.log(classSeperatedNames);
		classSeperatedNames.forEach((classList, index) => {
			if (classList.length !== 0) {
				out += `<li>MIDN ${index + 1}/C ${classList.join('  <b>|</b>  ')}</li>`;
			}
		});
		out += '</ul>';

		return out;
	};

	let emailBody = `${assignerData.name},
	<br><br>
	You assigned a ${submitData.paperwork} on ${dateToROTCFormat(submitData.dateAssigned)} to:
	<br>
	${namesToEmailFormat(authority)}
	because, ${submitData.reason}. It will be due COB ${dateToROTCFormat(submitData.dateDue)}.`;

	if (noAuthority.length > 0) {
		emailBody += `
		<br><br><br>
		You attempted to assign a ${submitData.paperwork} to:<br>
		${namesToEmailFormat(noAuthority)}
		but you do not have the authority.`;
	}

	emailBody += `
	<br><br>
	Very respectfully, <br>
	The ADMIN Department`;

	MailApp.sendEmail({
		to: assignerData.email,
		subject: `${authority.length === 0 ? 'No' : authority.length} ${submitData.paperwork}${
			authority.length > 1 || authority.length === 0 ? 's were' : ' was'
		} successfully assigned`,
		htmlBody: emailBody,
	});
}

/**
 *
 */
function dateToROTCFormat(date: Date): string {
	let dayNum = date.getDate();
	let monthNum = date.getMonth();
	let year = date.getFullYear();
	let month = '';
	let day = '';
	switch (monthNum) {
		case 0:
			month = 'JAN';
			break;
		case 1:
			month = 'FEB';
			break;
		case 2:
			month = 'MAR';
			break;
		case 3:
			month = 'APR';
			break;
		case 4:
			month = 'MAY';
			break;
		case 5:
			month = 'JUN';
			break;
		case 6:
			month = 'JUL';
			break;
		case 7:
			month = 'AUG';
			break;
		case 8:
			month = 'SEP';
			break;
		case 9:
			month = 'OCT';
			break;
		case 10:
			month = 'NOV';
			break;
		case 11:
			month = 'DEC';
			break;
	}
	if (dayNum < 10) {
		day = '0' + dayNum;
	} else {
		day = dayNum.toString();
	}
	return day + month + year;
}

/**
 *
 */
function sendAssigneesEmail(emailNameList, data) {
	const emailsActivated = ssOptions.getRange(1, 2).getValue().toString().toLowerCase() === 'true';
	if (!emailsActivated) return;

	if (emailNameList.length > 49) {
		sendAssigneesEmail(emailNameList.slice(49), data);
		emailNameList = emailNameList.slice(0, 49);
	}

	const dateDemo = data.dateDue.toString().split(' ', 4);

	const date = dateDemo[0] + ', ' + dateDemo[2] + dateDemo[1].toUpperCase() + dateDemo[3];

	const emailSender = getIndividualEmail(data.assigner);

	const emailSubject = 'NROTC ADMIN Department: New ' + data.paperwork + ' due COB ' + date + '.';

	var correctedEmail = '';
	let numEmailsBcc = 0;
	let lastNameEntered = '';
	for (let i = 0; i < emailNameList.length; i++) {
		if (emailNameList[i] === null || emailNameList[i] === '' || correctedEmail.indexOf(emailNameList[i]) !== -1) {
			continue;
		} else {
			lastNameEntered = emailNameList[i];
			numEmailsBcc++;
			if (correctedEmail === '') {
				correctedEmail = getIndividualEmail(emailNameList[i]);
			} else {
				correctedEmail = getIndividualEmail(emailNameList[i]) + ',' + correctedEmail;
			}
		}
	}

	const emailBody = `${numEmailsBcc === 1 ? lastNameEntered : 'Team'},
	<br>
	<br>You have been assigned a ${data.paperwork}, because ${data.reason}. It is due COB ${date}. ${
		data.pdfLink === '' ? '' : 'You can find the paperwork to complete here: ' + data.pdfLink
	}
	<br>
	<br>Very respectfully,
	<br>${data.assigner}`;

	Logger.log(emailNameList, emailSender);
	MailApp.sendEmail({
		to: Session.getEffectiveUser().getEmail(),
		bcc: correctedEmail,
		subject: emailSubject,
		htmlBody: emailBody,
	});
}
/**
 *
 */
function sendSheetNotEnabledEmail(submitterName) {
	MailApp.sendEmail({
		to: getIndividualEmail(submitterName),
		cc: Session.getEffectiveUser().getEmail(),
		subject: 'The Paperwork Database is Currently Disabled',
		htmlBody: `${submitterName},<br><br>The Paperwork database has not yet been initialized for the semester. Please reach out to your ADMIN Department to initialize the database.<br><br>Very respectfully,<br>The ADMIN Department`,
	});
}
/**
 *
 */
function sendInitReminderEmail() {
	Logger.log('Email sent to:' + Session.getEffectiveUser().getEmail());
	MailApp.sendEmail({
		to: Session.getEffectiveUser().getEmail(),
		subject: 'The Paperwork Database is Currently Disabled',
		htmlBody: `ADMINO and AADMINO,
		<br>
		<br>The paperwork database has not yet been initialized for the semester.
		<br>To do this you will need to follow the following steps:
		<ol type="1">
			<li>Open up the Main Database file in the Paperwork Database folder that you can find in the google drive.</li>
			<li>Start by updating the battaion structure for the semester. To do this click on the "Battalion Structure" tab and update the roles and groups columns. Then recreate the chain of command area. The Chain of command area should update its structure as you fill out the chain of command. There are notes in the headers for each of the columns which will help you fill out those areas.</li>
			<li>Then you should update the "Battalion Members" tab to include all of this semester's members. Make sure to update the classes of each member. A member will not appear in the system unless all 6 columns are completed. Also, make sure to check that all of the dropdown selections are vaild. If there is an invalid entree there will be a red arrow in the top right hand corner of the cell.</li>
			<li>Next you should look at the "Options" tab. Make sure send emails is true. Update the policy on the number of buisness days to complete a chit and negative counseling, make sure it is a number. Then update your preferance for how the sheet will handle assignment permissions and assignment due dates when no due date is specified. Row 6 will update itself.</li>
			<li>Now you can run the initialization function by clicking on the "DB functions" dropdown menu in the user interface at the top of the google sheet and clicking the "Initialize" option. The database will be successfully set up once you recieve the success email!</li>
		</ol>
		To manage this system follow these general steps:
		<ul>
			<li>You should never edit any of the data in the "Data" tab. It is just there for your own referance.</li>
			<li>The pending paperwork tab will hold all of the paperwork that still needs to be processed. You can change the status of the paperwork here and it will update everywhere.</li>
			<li>The "Digital Turn In Box" will hold all of the paperwork that has been turned via the google form.</li>
		</ul>
		If you would like to turn off this notification email for the semester because you aren't using the database. Go the the "Options" tab and change the "Turn off reminder" option to "true".
		<br>
		<br>If you have any problems or questions please feel free to reach out to us.
		<br>
		<br>Very respectfully,
		<br>Timothy Bowes (tnbowes@gmail.com)
		<br>John Lewis Corker (johnlcorker88@gmail.com)`,
	});
}
/**
 *
 */
function getFullMemberData(name: string): { name: string; email: string; role: string; group: string } {
	let fullData;
	JSON.parse(ssVariables.getRange(4, 2).getValue()).forEach((member) => {
		if (member.name === name) {
			fullData = member;
		}
	});
	return fullData;
}
/**
 *
 */
function dailyRunFunctions() {
	dailyCheckToRemindPplOfPaperwork();
	approveDueGoogleFormsFromDatabase();
}
/**
 *
 */
function approveDueGoogleFormsFromDatabase() {
	const pendingData = ssPending.getRange(1, 1, ssPending.getLastRow(), ssPending.getLastColumn()).getValues();
	const time = new Date();
	for (var idx = 0; idx < pendingData.length; idx++) {
		const dueDate = new Date(pendingData[idx][6].toString());
		if (pendingData[idx][7] === 'Pending' && dueDate.getTime() - time.getTime() < 0) {
			pendingData[idx][7] = 'Approved';
		}
	}
	ssPending.getRange(1, 1, ssPending.getLastRow(), ssPending.getLastColumn()).setValues(pendingData);
}
/**
 *
 */
function dailyCheckToRemindPplOfPaperwork() {
	if (ssData.getLastRow() > 1) {
		const data = ssData.getRange(2, 1, ssData.getLastRow() - 1, ssData.getLastColumn()).getValues();
		const tomorrow = new Date();
		tomorrow.setDate(tomorrow.getDate() + 1);
		const paperworkTypes = {};
		for (let i = 0; i < data.length; i++) {
			const dueDate = new Date(data[i][6].toString());
			if (
				data[i][7] === 'Pending' &&
				tomorrow.getDate() === dueDate.getDate() &&
				tomorrow.getMonth() === dueDate.getMonth() &&
				dueDate.getFullYear() === tomorrow.getFullYear()
			) {
				paperworkTypes[data[i][4]] =
					paperworkTypes[data[i][4]] === undefined ? [data[i][3]] : paperworkTypes[data[i][4]].push(data[i][3]);
			}
		}
		sendPaperworkHeadsUpNotification(paperworkTypes);
	}
}

function sendPaperworkHeadsUpNotification(paperworkTypes) {
	for (const key in paperworkTypes) {
		sendEmail(key, paperworkTypes[key]);
	}
	function sendEmail(key, names) {
		if (names.length > 49) {
			sendEmail(key, names.slice(49));
			names = names.slice(0, 49);
		}
		const bccEmails = names.map((name) => getIndividualEmail(name));

		MailApp.sendEmail({
			to: Session.getEffectiveUser().getEmail(),
			bcc: bccEmails.join(),
			subject: `Your ${key} is due tomorrow`,
			htmlBody: `${
				names.length === 1 ? names[0] : 'Team'
			},<br><br>Your ${key} is due tomorrow.<br><br>Very respectfully,<br>The ADMIN Department`,
		});
	}
}
