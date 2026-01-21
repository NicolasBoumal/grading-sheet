/**
* @OnlyCurrentDoc
*/

function onOpen() {
  // Hide helper sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName("Template_").hideSheet();
  ss.getSheetByName("Emails").hideSheet();
  ss.getSheetByName("Totals").hideSheet();

  // Create a custom menu
  SpreadsheetApp.getUi()
    .createMenu('⭐ Grading Sheet admin tools ⭐')
    .addItem('Send / resume sending E-mails', 'sendAllEmails')
    .addSeparator()
    .addItem('Reset spreadsheet (cannot be undone!)', 'resetSpreadsheet')
    .addSeparator()
    .addItem('About', 'about')
    .addToUi();
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const cellAddress = range.getA1Notation();
  const newValue = e.value;
  const ui = SpreadsheetApp.getUi();

  // Upon checking the "Problems are done" checkbox
  if (sheet.getName() === "Problems" && cellAddress === "F7") {
    if (newValue === "TRUE") {

      // Reset the checkbox first, because onEdit triggers have a 30-second
      // runtime limit that does not pause while waiting for user input.
      range.setValue(false);

      if (!validateSheetNames()) {
        return;
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // Check if at least one problem sheet was already created
      const notTheFirstTime = ss.getSheets().some(s => {
        return s.getName() === sheet.getRange("A2").getDisplayValue().toString().trim();
      });

      if (notTheFirstTime || ui.alert("Once this is checked, the list of problems and their names should no longer be changed.\n\nProceed?\n\n(Allow up to 30 seconds for the script to run. If not all problem sheets have been created, check the box anew until completion.)", ui.ButtonSet.YES_NO) == "YES") {

        createSheetsFromList();

        // The script may run out of time before getting here.

        ss.getSheetByName("Totals").showSheet();

        // Prevent further edits to the list of problem names
        sheet.getRange("A:A").protect().setDescription("Problem names should no longer be changed.").setWarningOnly(true);

        // We are done doing all the things that need to happen after checking the box,
        // so it is safe to show it as "checked".
        range.setValue(true);
        range.protect().setDescription("This checkbox served its purpose.").setWarningOnly(true);

        ui.alert("Problem names are now protected (weakly) against accidental edits.");

      }
    }
  }

  // Upon checking the "students and groups are done" checkbox
  if (sheet.getName() === "Students" && cellAddress === "K8") {
    if (newValue === "TRUE") {

      // Reset the checkbox first, because onEdit triggers have a 30-second
      // runtime limit that does not pause while waiting for user input.
      range.setValue(false);

      if (ui.alert("Once this is checked, student names and their groups should no longer be changed.\n\nProceed?\n\n(If so, then click YES within 30 seconds.)", ui.ButtonSet.YES_NO) == "YES") {
        
        // Prevent further edits to student names and group names.
        lockStudentsGroups();

        // We are done doing all the things that need to happen after checking the box,
        // so it is safe to show it as "checked".
        range.setValue(true);
        range.protect().setDescription("This checkbox served its purpose.").setWarningOnly(true);

        ui.alert("Student names and group names are now protected (weakly) against accidental edits.");

      }
    }
  }

  // Upon selecting a group to preview e-mail
  if (sheet.getName() === "Emails" && range.getA1Notation() === "G29") {
    const groupName = e.value;
    previewEmail(groupName);
  }
}

function lockStudentsGroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName("Students");
  const emailSheet = ss.getSheetByName("Emails");

  // Do this first, as it is more important.
  emailSheet.showSheet();

  // Prevent further edits to the list of student names and groups, and to the checkbox itself.
  studentSheet.getRange("A:A").protect().setDescription("Student names should no longer be changed (though it's really their ordering that should be preserved; typos are OK to fix).").setWarningOnly(true);
  studentSheet.getRange("C:C").protect().setDescription("Student groups should no longer be changed (though their specific group name would be OK to change).").setWarningOnly(true);
}

function previewEmail(groupName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailSheet = ss.getSheetByName("Emails");
  const toCell = emailSheet.getRange("G31");
  const subjectCell = emailSheet.getRange("G30");
  const bodyRange = emailSheet.getRange("F32:G50");

  if (!groupName) {
    [toCell, subjectCell, bodyRange].forEach(r => r.clearContent());
    return;
  }

  try {
    const data = getEmailData(groupName);
    
    toCell.setValue(data.to);
    subjectCell.setValue(data.subject);
    bodyRange.setValue(data.body);
    
    bodyRange.setVerticalAlignment("top");
    bodyRange.setWrap(true);
  } catch (err) {
    bodyRange.setValue("Error generating preview: " + err.message);
  }
}

function about() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'About Grading Sheet',
    'Created by Nicolas Boumal with Gemini.\n\n' +
    '1. List the student names, and optionally group them.\n' +
    '2. Check the box once this is done.\n' +
    '3. List the problem names.\n' +
    '4. Check the box once this is done.\n\n' +
    'Then, in each problem sheet, create one column per type ' +
    'of mistake, choose the associated penalty, and indicate ' +
    'which groups made that mistake ' +
    '(e.g., with a 1, or .5 if less severe).\n\n' +
    'Optionally, when done, you can send an e-mail to each group.\n\n' + 
    'Color conventions:\n' + 
    ' * green cells: fill in once when setting up, then leave as is.\n' + 
    ' * blue cells: fill in whenever you want, and edit at will.\n' + 
    ' * yellow cells: these are computed automatically: do not edit.',
    ui.ButtonSet.OK
  );
}

function getLastRowInColumn(sheet, columnLetter) {
  const range = sheet.getRange(columnLetter + ":" + columnLetter);
  const values = range.getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0].toString().trim() !== "") {
      return i + 1;
    }
  }
  return 0;
}

function createSheetsFromList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const problemSheet = ss.getSheetByName("Problems");
  const templateSheet = ss.getSheetByName("Template_");
  
  const lastRow = getLastRowInColumn(problemSheet, "A");
  if (lastRow < 2) return;
  
  // Get the list of problem names
  const names = problemSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
  
  names.forEach((name, index) => {
    const sheetName = name.trim();
    
    // Calculate the actual row number in the 'Problems' sheet
    // index starts at 0, and our data starts at row 2, so:
    const currentRow = index + 2; 

    // If the sheet does not exist yet, create it
    if (sheetName !== "" && !ss.getSheetByName(sheetName)) {

      const newSheet = templateSheet.copyTo(ss);
      newSheet.setName(sheetName).showSheet();
      
      newSheet.getRange("A1").setFormula("='Problems'!A" + currentRow);
      newSheet.getRange("A2").setFormula("='Problems'!B" + currentRow);
      
      // ss.setActiveSheet(newSheet);
      // ss.moveActiveSheet(ss.getNumSheets());

    }
  });
  
  ss.setActiveSheet(problemSheet);
}

function validateSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const problemSheet = ss.getSheetByName("Problems");

  const lastRow = getLastRowInColumn(problemSheet, "A");
  if (lastRow < 2) {
    ui.alert("The list is empty. Please add problem names in column A, starting from row 2.");
    return false;
  }

  const names = problemSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
  const existingSheetNames = ss.getSheets().map(s => s.getName().toLowerCase());
  const invalidChars = /[\\\/\?\*\:\[\]]/;
  
  let errors = [];
  let seenInList = new Set();

  names.forEach((name, index) => {
    const rowNum = index + 2;
    const trimmedName = name.toString().trim();

    // 1. Check for empty names
    if (trimmedName === "") {
      errors.push(`- Cell A${rowNum}: Problem name is blank (problem names should be listed consecutively).`);
      return;
    }

    // 2. Check for length (Max 100 chars)
    if (trimmedName.length > 100) {
      errors.push(`- Cell A${rowNum}: Problem name is too long (max 100 characters).`);
    }

    // 3. Check for forbidden characters
    if (invalidChars.test(trimmedName)) {
      errors.push(`- Cell A${rowNum}: Problem name contains forbidden characters ( / \\ ? * : [ ] ).`);
    }

    // 4. Check for duplicates within the Problems list itself
    if (seenInList.has(trimmedName.toLowerCase())) {
      errors.push(`- Cell A${rowNum}: Duplicate problem name "${trimmedName}" found within your list.`);
    }
    seenInList.add(trimmedName.toLowerCase());

    // 5. Check against reserved sheet names (Case-Insensitive)
    const reserved = ["students", "problems", "emails", "template_", "totals"];
    if (reserved.includes(trimmedName.toLowerCase())) {
      errors.push(`- Cell A${rowNum}: "${trimmedName}" cannot be used as a problem name.`);
    }
  });

  if (errors.length > 0) {
    const message = "Please fix the following issues before creating the problem sheets:\n\n" + errors.join("\n");
    ui.alert("Problem names should be changed", message, ui.ButtonSet.OK);
    return false;
  }

  return true; // All checks passed
}

function resetSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'CAREFUL', 
    'This will delete all problem sheets and clear data. This cannot be undone.\n\nTo proceed, type "RESET" in the box below:', 
    ui.ButtonSet.OK_CANCEL
  );
  
  // Check if the user clicked OK and typed the word exactly
  if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText().toUpperCase() !== 'RESET') {
    ui.alert('Action cancelled. The spreadsheet was not reset.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studSheet = ss.getSheetByName("Students");
  const probSheet = ss.getSheetByName("Problems");
  const mailSheet = ss.getSheetByName("Emails");

  // 1. Loop through problem names in Problems!A2:A and delete sheets/data
  const lastRow = probSheet.getLastRow();
  if (lastRow >= 2) {
    const names = probSheet.getRange("A2:A" + lastRow).getValues();
    names.forEach(row => {
      let sheetToDelete = ss.getSheetByName(row[0]);
      if (sheetToDelete) ss.deleteSheet(sheetToDelete);
    });
    // Clear names in Column A and values in Column B
    probSheet.getRange("A2:B" + lastRow).clearContent();
  }

  // 2. Uncheck checkboxes
  studSheet.getRange("K8").setValue(false);
  probSheet.getRange("F7").setValue(false);

  // 3. Remove protections from specific ranges
  const protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const targets = ["Problems!A:A", "Problems!F7", "Students!A:A", "Students!C:C", "Students!K8"];
  
  protections.forEach(p => {
    let a1 = p.getRange().getA1Notation();
    let sName = p.getRange().getSheet().getName();
    if (targets.includes(sName + "!" + a1) || targets.includes(sName + "!" + a1.split(':')[0] + ":" + a1.split(':')[0])) {
      p.remove();
    }
  });

  // 4. Reset things on the E-mails sheet
  mailSheet.getRange("G29").setValue("");  // selector for e-mail preview
  mailSheet.getRange("G30").setValue("");
  mailSheet.getRange("G31").setValue("");
  mailSheet.getRange("F32:G50").setValue("");
  const mailLastRow = mailSheet.getLastRow();
  if (mailLastRow >= 2) {
    mailSheet.getRange("D2:D" + mailLastRow).clearContent(); 
  }

  // 5. Hide some sheets
  ss.getSheetByName("Totals").hideSheet();
  ss.getSheetByName("Emails").hideSheet();
}

/**
 * Assembles all data for a specific group and returns the parsed email data.
 */
function getEmailData(groupName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailSheet = ss.getSheetByName("Emails");
  const studentSheet = ss.getSheetByName("Students");
  const totalsSheet = ss.getSheetByName("Totals");

  // A. GET TEMPLATES & SETTINGS
  const subjectTemplate = emailSheet.getRange("G4").getValue(); // Updated to G4
  const ccAddress = emailSheet.getRange("G6").getValue();        // Added CC from G6
  const workName = emailSheet.getRange("G2").getValue();

  // B. GET RECIPIENT & LOCATION FROM Emails!A:C
  const emailSheetData = emailSheet.getRange("A:C").getValues();
  const emailRow = emailSheetData.find(r => r[0].toString().trim() === groupName.toString().trim());
  const recipientEmail = emailRow ? emailRow[2] : ""; 
  const locationValue = emailRow ? emailRow[1] : "";

  // C. GET TOTAL GRADE & MAX GRADE (Totals Sheet)
  const totalsData = totalsSheet.getRange("A:B").getValues();
  const maxGrade = Number(totalsSheet.getRange("B2").getValue()).toFixed(2);
  const groupRowInTotals = totalsData.find(r => r[0].toString().trim() === groupName.toString().trim());
  const groupTotalScore = groupRowInTotals ? Number(groupRowInTotals[1]).toFixed(2) : "0.00";
  const totalDisplay = `${groupTotalScore} / ${maxGrade}`;

  // D. GET STUDENT NAMES & BONUS/MALUS (Students Sheet)
  const studentsData = studentSheet.getRange("A2:E" + studentSheet.getLastRow()).getValues();
  const groupMembers = studentsData.filter(r => (r[2] === groupName || (r[2] === "" && r[0] === groupName)));
  
  const formattedNames = groupMembers.map(r => r[0]).join(", ").replace(/, ([^,]*)$/, ' and $1');

  let bonusLines = [];
  let hasAnyAdjustment = false;
  groupMembers.forEach(member => {
    if (member[3] !== "" && member[3] !== 0 && member[3] !== null) {
      hasAnyAdjustment = true;
      let line = `${member[0]}: ${member[3] > 0 ? "+" : ""}${Number(member[3]).toFixed(2)}`;
      if (member[4]) line += ` (${member[4]})`;
      bonusLines.push(line);
    }
  });

  // E. ASSEMBLE PLACEHOLDER OBJECT
  const groupData = {
    names: formattedNames,
    entityName: groupName,
    workName: workName,
    location: locationValue,
    gradesList: getProblemGradesList(groupName),
    total: totalDisplay,
    bonusText: bonusLines.join("\n"),
    bonusmalus: hasAnyAdjustment 
  };

  // F. FINAL STRINGS
  const finalSubject = subjectTemplate
    .replace("{{student or group name}}", groupName)
    .replace("{{work name}}", workName);
    
  const finalBody = generateEmail(groupData);

  return {
    to: recipientEmail,
    cc: ccAddress,
    subject: finalSubject,
    body: finalBody
  };
}

/**
 * Helper function to build the vertical list of problem grades.
 */
function getProblemGradesList(groupName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const totalsSheet = ss.getSheetByName("Totals");
  
  const fullData = totalsSheet.getDataRange().getValues();
  if (fullData.length < 2) return "No data found.";

  const headers = fullData[0];   // Row 1 (Problem Names)
  const maxPoints = fullData[1]; // Row 2 (Max Points)
  
  const groupRow = fullData.find(row => row[0] === groupName);
  if (!groupRow) return "Grade details missing.";

  let gradesArray = [];
  
  // Problems start in Column C (Index 2)
  for (let i = 2; i < groupRow.length; i++) {
    let problemName = headers[i];
    let score = Number(groupRow[i]).toFixed(2);
    let max = Number(maxPoints[i]).toFixed(2);
    
    if (problemName && problemName !== "") {
      gradesArray.push(`${problemName}: ${score} / ${max}`);
    }
  }
  
  return gradesArray.join("\n");
}

function generateEmail(groupData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailSheet = ss.getSheetByName("Emails");

  let template = emailSheet.getRange("F9:G27").getValues().flat().join("\n").trim();

  // A. Handle Conditional Blocks (if/endif)
  // This regex finds {{if key}} ... {{endif key}} and either keeps or deletes the content
  const conditionals = ["location", "bonusmalus"];
  conditionals.forEach(key => {
    const regex = new RegExp(`{{if ${key}}}([\\s\\S]*?){{endif ${key}}}`, "g");
    if (!groupData[key]) {
      template = template.replace(regex, ""); // Remove the whole block
    } else {
      template = template.replace(regex, "$1"); // Keep content inside, remove tags
    }
  });

  // B. Handle Simple Replacements
  const replacements = {
    "{{student names}}": groupData.names,
    "{{student or group name}}": groupData.entityName,
    "{{location}}": groupData.location,
    "{{work name}}": groupData.workName,
    "{{problem grades}}": groupData.gradesList,
    "{{total group grade}}": groupData.total,
    "{{bonusmalus}}": groupData.bonusText
  };

  for (let placeholder in replacements) {
    template = template.split(placeholder).join(replacements[placeholder] || "");
  }

  return template;
}

function sendAllEmails() {
  const ui = SpreadsheetApp.getUi();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailSheet = ss.getSheetByName("Emails");

  const lastRow = emailSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No groups found to email.");
    return;
  }

  const response = ui.prompt(
    'CAREFUL', 
    'This will send e-mails to all groups not already marked as SENT in column D.\n\nTo proceed, type "SEND" in the box below:',
    ui.ButtonSet.OK_CANCEL
  );
  
  // Check if the user clicked OK and typed the word exactly
  if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText().toUpperCase() !== 'SEND') {
    ui.alert('Action cancelled. The e-mails were not sent.');
    return;
  }
  
  // Get Group (Col A) through Status (Col D)
  const dataRange = emailSheet.getRange(2, 1, lastRow - 1, 4); 
  const data = dataRange.getValues();

  data.forEach((row, index) => {
    const groupName = row[0];
    const status = row[3]; // Column D

    if (groupName && status !== "SENT") {
      try {
        const emailOptions = getEmailData(groupName); // This returns {to, cc, subject, body} 

        if (emailOptions.to) {
          // Pass the object directly 
          MailApp.sendEmail(emailOptions);

          // Mark as sent
          emailSheet.getRange(index + 2, 4).setValue("SENT");
          
          Utilities.sleep(200); 
        }
      } catch (err) {
        console.error("Failed to send to " + groupName + ": " + err.message);
      }
    }
  });

  ui.alert("Emailing process complete.");
}
