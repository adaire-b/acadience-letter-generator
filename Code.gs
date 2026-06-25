/*
ACADIENCE LETTER GENERATOR

Primary user function:
runLetterWorkflow()

Main workflow:
1. Import latest CSV
2. Process data
3. Generate teacher letter packets
4. Update dashboard

Parked:
- PDF generation
- Dynamic logo support

*/

/*
FINAL CLEANUP BEFORE RELEASE:
- Protect Dashboard, Set-up Wizard, Help, Skipped Students, and Run Log.
- Allow editing only in intended input cells.
- Hide backend sheets: Settings, Raw Import, Processed Data, Benchmark Goals.
- Test all buttons from a fresh copied file.
- Create demo version with fake student data.
- Prepare portfolio/presentation summary.
*/

/*************************************************
 * WORKFLOW FUNCTIONS
 *************************************************/
function importLatestCsv() {
  const ss = getSpreadsheet_();

  const settings = getSettings_();
  const uploadFolderId = settings["Upload Folder ID"];

  if (!uploadFolderId) {
    throw new Error("Missing Upload Folder ID in Settings tab.");
  }

  const uploadFolder = DriveApp.getFolderById(uploadFolderId);
  const files = uploadFolder.getFilesByType(MimeType.CSV);

  let latestFile = null;
  let latestDate = null;

  while (files.hasNext()) {
    const file = files.next();
    const date = file.getLastUpdated();

    if (!latestDate || date > latestDate) {
      latestDate = date;
      latestFile = file;
    }
  }

 if (!latestFile) {
  throw new Error("No CSV files found in the Uploads folder.");
}

setDashboardStatus_("Acadience File Found", "✓", "Acadience CSV imported.");

safeSetNamedRange_("Status_File_Pill", "READY");

SpreadsheetApp.flush();

  const csvText = latestFile.getBlob().getDataAsString();
  const csvData = Utilities.parseCsv(csvText);

  const rawSheet = ss.getSheetByName("Raw Import");
  rawSheet.clearContents();

  rawSheet
    .getRange(1, 1, csvData.length, csvData[0].length)
    .setValues(csvData);

  logRun_("CSV Imported", `Imported ${latestFile.getName()} with ${csvData.length - 1} student rows.`);
}

/*************************************************
 * SETTINGS / LOGGING
 *************************************************/
function getSettings_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName("Settings");

  const values = sheet.getDataRange().getValues();
  const settings = {};

  for (let i = 1; i < values.length; i++) {
    const key = values[i][0];
    const value = values[i][1];

    if (key) {
      settings[key] = value;
    }
  }

  return settings;
}


function logRun_(status, message) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName("Run Log");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Status", "Message"]);
  }

  sheet.appendRow([new Date(), status, message]);
} 

/*************************************************
 * DATA PROCESSING
 *************************************************/
function processImportedData() {
  const ss = getSpreadsheet_();

  const rawSheet = ss.getSheetByName("Raw Import");
  const processedSheet = ss.getSheetByName("Processed Data");
  const skippedSheet = ss.getSheetByName("Skipped Students");
  


const selectedWindow = normalizeWindow_(
  ss.getSheetByName("Dashboard").getRange("L11").getValue()
);

  if (!selectedWindow) {
    throw new Error("Assessment Window is missing on the Dashboard.");
  }

  const rawData = rawSheet.getDataRange().getValues();
  const headers = rawData[0];
  const rows = rawData.slice(1);
  const headerMap = createHeaderMap_(headers);
  const goalMap = createBenchmarkGoalMap_();

  const processed = [];
  const skipped = [];

 rows.forEach(row => {
  const studentId = getValue_(row, headerMap, "Student Number");
  const schoolYear = getValue_(row, headerMap, "School Year");
  const firstName = getValue_(row, headerMap, "Student First Name");
  const lastName = getValue_(row, headerMap, "Student Last Name");
  const teacherFirst = getValue_(row, headerMap, "Teacher First Name");
  const teacherLast = getValue_(row, headerMap, "Teacher Last Name");
  const grade = normalizeGrade_(getValue_(row, headerMap, "Student Grade Level"));
  const window = normalizeWindow_(getValue_(row, headerMap, "Benchmark Period"));
  
if (!window && ["K", "1", "2", "3"].includes(String(grade).trim())) {
  skipped.push([
    `${firstName} ${lastName}`.trim(),
    `${teacherFirst} ${teacherLast}`.trim(),
    grade,
    selectedWindow,
    "",
    "",
    "",
    "",
    "No benchmark data found for selected window",
    new Date()
  ]);
  return;
}

  if (window !== selectedWindow) {
    return;
  }

  const readingScore = getValue_(row, headerMap, "Reading Composite Score");


    const readingStatus = getValue_(row, headerMap, "Reading Composite Status");
    const mathScore = getValue_(row, headerMap, "Math Composite Score");
    const mathStatus = getValue_(row, headerMap, "Math Composite Status");

    const studentName = `${firstName} ${lastName}`.trim();
    const teacherName = `${teacherFirst} ${teacherLast}`.trim();


    let skipReason = "";

    if (!studentName) skipReason = "Missing student name";
    else if (!teacherName) skipReason = "Missing teacher name";
    else if (!["K", "1", "2", "3"].includes(String(grade).trim())) return;
    else if (!readingScore || !readingStatus) skipReason = "Missing reading composite data";
    else if (!mathScore || !mathStatus) skipReason = "Missing math composite data";

    const goalKey = `${grade}|${window}`;
    const goals = goalMap[goalKey];

    if (!skipReason && !goals) {
      skipReason = `Missing benchmark goal range for ${grade} ${window}`;
    }

    if (skipReason) {
      skipped.push([
        studentName,
        teacherName,
        grade,
        window,
        readingScore,
        readingStatus,
        mathScore,
        mathStatus,
        skipReason,
        new Date()
      ]);
      return;
    }

    const readingConcern = statusToConcernLevel_(readingStatus);
    const mathConcern = statusToConcernLevel_(mathStatus);
    const overallConcern = getOverallConcernLevel_(readingConcern, mathConcern);
    const primaryConcern = getPrimaryConcernArea_(readingConcern, mathConcern);

    processed.push([
      studentId,
      studentName,
      teacherName,
      grade,
      window,
      schoolYear,
      readingScore,
      readingStatus,
      goals.reading,
      readingConcern,
      mathScore,
      mathStatus,
      goals.math,
      mathConcern,
      overallConcern,
      primaryConcern,
      true,
      false,
      ""
    ]);
  });
 
  clearDataKeepHeaders_(processedSheet);
  clearDataKeepHeaders_(skippedSheet);

  if (processed.length > 0) {
    processedSheet.getRange(2, 1, processed.length, processed[0].length).setValues(processed);
  }

  if (skipped.length > 0) {
    skippedSheet.getRange(2, 1, skipped.length, skipped[0].length).setValues(skipped);
  }

  logRun_("Data Processed", `Processed ${processed.length} students. Skipped ${skipped.length} students.`);
}

function runLetterWorkflow() {
  const ss = getSpreadsheet_();
  const processedSheet = ss.getSheetByName("Processed Data");

  const hasProcessedData = processedSheet.getLastRow() > 1;

  if (!hasProcessedData) {
    importLatestCsv();
    processImportedData();
  }

  generateTeacherPdfBatch7();
}

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  return SpreadsheetApp.openById(id);
}

/*************************************************
 * DATA PROCESSING HELPERS
 *************************************************/
function createHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[String(header).trim()] = index;
  });
  return map;
}

function getValue_(row, headerMap, headerName) {
  const index = headerMap[headerName];
  if (index === undefined) return "";
  return row[index];
}

function createBenchmarkGoalMap_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName("Benchmark Goals");
  const values = sheet.getDataRange().getValues();

  const map = {};

  for (let i = 1; i < values.length; i++) {
    const grade = normalizeGrade_(values[i][0]);
    const window = normalizeWindow_(values[i][1]);
    const reading = values[i][2];
    const math = values[i][3];

    if (grade && window) {
      map[`${grade}|${window}`] = { reading, math };
    }
  }

  return map;
}

function normalizeGrade_(grade) {
  const value = String(grade).trim();

  if (value.toLowerCase() === "kindergarten") return "K";
  if (value === "0") return "K";

  return value;
}

function normalizeWindow_(window) {
  const value = String(window).trim().toUpperCase();

  if (value.includes("BEGIN")) return "BOY";
  if (value.includes("MIDDLE")) return "MOY";
  if (value.includes("END")) return "EOY";

  if (value === "BOY" || value === "MOY" || value === "EOY") return value;

  return value;
}

function statusToConcernLevel_(status) {
  const value = String(status).trim().toLowerCase();

  if (value.includes("well below")) return "Intensive";
  if (value.includes("below")) return "Strategic";
  if (value.includes("benchmark")) return "Core";

  return "";
}

function getOverallConcernLevel_(readingConcern, mathConcern) {
  if (readingConcern === "Intensive" || mathConcern === "Intensive") return "Intensive";
  if (readingConcern === "Strategic" || mathConcern === "Strategic") return "Strategic";
  return "Core";
}

function getPrimaryConcernArea_(readingConcern, mathConcern) {
  const readingConcerned = readingConcern === "Strategic" || readingConcern === "Intensive";
  const mathConcerned = mathConcern === "Strategic" || mathConcern === "Intensive";

  if (readingConcerned && mathConcerned) return "Reading + Math";
  if (readingConcerned) return "Reading";
  if (mathConcerned) return "Math";
  return "None";
}

function clearDataKeepHeaders_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
  }
}

function storeIdNow() {
  PropertiesService.getScriptProperties()
    .setProperty("SPREADSHEET_ID", SpreadsheetApp.getActiveSpreadsheet().getId());
  Logger.log("Stored ID: " + SpreadsheetApp.getActiveSpreadsheet().getId());
}
/*************************************************
 * LETTER GENERATION
 *************************************************/
function generateTestLetter() {
  Logger.log("Starting generateTestLetter...");

  const ss = getSpreadsheet_();

  const settings = getSettings_();
  const templateDocId = settings["Template Doc ID"];
  const outputFolderId = settings["Output Folder ID"];

  if (!templateDocId) throw new Error("Missing Template Doc ID in Settings tab.");
  if (!outputFolderId) throw new Error("Missing Output Folder ID in Settings tab.");


 const schoolName = ss.getRangeByName("School_Name").getValue();
const signerName = ss.getRangeByName("Signer_Name").getValue();
const letterDate = ss.getRangeByName("Letter_Date").getDisplayValue();

  const processedSheet = ss.getSheetByName("Processed Data");
  const data = processedSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const headerMap = createHeaderMap_(headers);

  const row = rows.find(r => {
    const generateLetter = getValue_(r, headerMap, "Generate Letter");
    return generateLetter === true || String(generateLetter).toLowerCase() === "true";
  });

  if (!row) {
    throw new Error("No rows found with Generate Letter = TRUE.");
  }

  const studentName = getValue_(row, headerMap, "Student Name");
  Logger.log("Student selected: " + studentName);

  const outputFolder = DriveApp.getFolderById(outputFolderId);
  const templateFile = DriveApp.getFileById(templateDocId);

  Logger.log("Making template copy...");
  const copiedFile = templateFile.makeCopy(`TEST LETTER - ${studentName}`, outputFolder);

  Logger.log("Opening copied document...");
  const doc = DocumentApp.openById(copiedFile.getId());
  const body = doc.getBody();

  const replacements = {
    "{{SCHOOL_NAME}}": schoolName,
    "{{ASSESSMENT_WINDOW}}": windowToReportSeason_(getValue_(row, headerMap, "Assessment Window")),
    "{{WINDOW_CODE}}": getValue_(row, headerMap, "Assessment Window"),
   "{{SCHOOL_YEAR}}": getValue_(row, headerMap, "School Year"),
    "{{TEACHER_NAME}}": getValue_(row, headerMap, "Teacher Name"),
    "{{GRADE}}": getValue_(row, headerMap, "Grade"),
    "{{LETTER_DATE}}": letterDate,
    "{{STUDENT_FULL_NAME}}": studentName,
    "{{READING_GOAL_RANGE}}": getValue_(row, headerMap, "Reading Goal Range"),
    "{{READING_SCORE}}": getValue_(row, headerMap, "Reading Score"),
    "{{READING_STATUS}}": titleCase_(getValue_(row, headerMap, "Reading Status")),
    "{{MATH_GOAL_RANGE}}": getValue_(row, headerMap, "Math Goal Range"),
    "{{MATH_SCORE}}": getValue_(row, headerMap, "Math Score"),
    "{{MATH_STATUS}}": titleCase_(getValue_(row, headerMap, "Math Status")),
    "{{SIGNER_NAME}}": signerName
  };

  Logger.log("Replacing fields...");
  Object.keys(replacements).forEach(key => {
    body.replaceText(escapeRegExp_(key), String(replacements[key] ?? ""));
  });
  

  Logger.log("Saving document...");
  doc.saveAndClose();

  logRun_("Test Letter Processed", `Created test letter for ${studentName}.`);
  Logger.log("Test letter complete: " + copiedFile.getUrl());
}
function generateTeacherPdfTest() {
  generateTeacherPdfs_({
    limitTeachers: 1,
    markGenerated: false,
    prefix: "TEST - "
  });
}

function startTeacherPacketWorkflow_() {
  resetPacketStatus_();

  return generateTeacherPdfs_({
    limitTeachers: null,
    markGenerated: true,
    prefix: ""
  });
}

function testContinue() {
  continuePacketGeneration_();
}

function continuePacketGeneration_() {
  generateTeacherPdfs_({
    limitTeachers: null,
    markGenerated: true,
    prefix: ""
  });

  if (hasUnprocessedTeachers_()) {
    ScriptApp.newTrigger("continuePacketGeneration_")
      .timeBased()
      .after(60 * 1000)
      .create();
  }
}

function generateTeacherPdfs() {
  startTeacherPacketWorkflow_();
}

function generateTeacherPdfBatch7() {
  resetPacketStatus_();

  generateTeacherPdfs_({
    limitTeachers: 7,
    markGenerated: true,
    prefix: ""
  });
}

function generateTeacherPdfs_(options) {
  options = options || {};

  const ss = getSpreadsheet_();;
  




try {

const startTime = new Date().getTime();

const settings = getSettings_();
const dash = ss.getSheetByName("Dashboard");
const schoolName = dash.getRange("K17").getValue();
const signerName = dash.getRange("K20").getValue();
const letterDate = dash.getRange("L14").getDisplayValue();
  

  const templateDocId = settings["Template Doc ID"];
  const outputFolderId = settings["Output Folder ID"];

  const outputFolder = DriveApp.getFolderById(outputFolderId);
  const templateDoc = DocumentApp.openById(templateDocId);
  const templateBody = templateDoc.getBody();  

  const processedSheet = ss.getSheetByName("Processed Data");
  const data = processedSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const headerMap = createHeaderMap_(headers);

  const teacherGroups = {};

let packetsCreated = 0;
let studentsProcessed = 0;
let skippedCount = 0;

  

rows.forEach((row, index) => {
  const generateLetter = getValue_(row, headerMap, "Generate Letter");
  const letterProcessed = getValue_(row, headerMap, "Letter Processed");

  if (!(generateLetter === true || String(generateLetter).toLowerCase() === "true")) {
    skippedCount++;
    return;
  }

  if (letterProcessed === true || String(letterProcessed).toLowerCase() === "true") {
    skippedCount++;
    appendSkippedStudent_(row, headerMap, "Letter previously generated");
    return;
  }
    const teacherName = getValue_(row, headerMap, "Teacher Name");
    if (!teacherName) return;

    if (!teacherGroups[teacherName]) teacherGroups[teacherName] = [];
    teacherGroups[teacherName].push({
      row,
      sheetRow: index + 2
    });
  });

  let teacherNames = Object.keys(teacherGroups).sort();

if (options.startTeacherIndex !== undefined || options.limitTeachers) {
  const startTeacherIndex = Number(options.startTeacherIndex || 0);
  const limitTeachers = options.limitTeachers || teacherNames.length;

  teacherNames = teacherNames.slice(
    startTeacherIndex,
    startTeacherIndex + limitTeachers
  );
} else if (options.limitTeachers) {
  teacherNames = teacherNames.slice(0, options.limitTeachers);
}

  for (let t = 0; t < teacherNames.length; t++) {
    const teacherName = teacherNames[t];

    if (new Date().getTime() - startTime > 270000) {
      break;
    }

    const students = teacherGroups[teacherName];

    students.sort((a, b) => {
      return getLastName_(getValue_(a.row, headerMap, "Student Name"))
        .localeCompare(getLastName_(getValue_(b.row, headerMap, "Student Name")));
    });

    const firstRow = students[0].row;
    const season = windowToReportSeason_(getValue_(firstRow, headerMap, "Assessment Window"));
    const schoolYear = getValue_(firstRow, headerMap, "School Year");

    const teacherLastName = getLastName_(teacherName);
const docName = `${options.prefix || ""}${teacherLastName} ${season} Benchmark Letters ${schoolYear}`;

deleteExistingFilesByName_(outputFolder, docName);

const templateFile = DriveApp.getFileById(templateDocId);
const packetFileCopy = templateFile.makeCopy(docName, outputFolder);

const packetDoc = DocumentApp.openById(packetFileCopy.getId());
Logger.log("Packet Doc ID: " + packetDoc.getId());

const packetBody = packetDoc.getBody();
packetBody.clear();

    students.forEach((studentObj, studentIndex) => {
      const row = studentObj.row;

      const replacements = {
        "{{SCHOOL_NAME}}": schoolName,
        "{{ASSESSMENT_WINDOW}}": season,
        "{{SCHOOL_YEAR}}": getValue_(row, headerMap, "School Year"),
        "{{TEACHER_NAME}}": getValue_(row, headerMap, "Teacher Name"),
        "{{GRADE}}": getValue_(row, headerMap, "Grade"),
        "{{LETTER_DATE}}": letterDate,
        "{{STUDENT_FULL_NAME}}": getValue_(row, headerMap, "Student Name"),
        "{{READING_GOAL_RANGE}}": getValue_(row, headerMap, "Reading Goal Range"),
        "{{READING_SCORE}}": getValue_(row, headerMap, "Reading Score"),
        "{{READING_STATUS}}": titleCase_(getValue_(row, headerMap, "Reading Status")),
        "{{MATH_GOAL_RANGE}}": getValue_(row, headerMap, "Math Goal Range"),
        "{{MATH_SCORE}}": getValue_(row, headerMap, "Math Score"),
        "{{MATH_STATUS}}": titleCase_(getValue_(row, headerMap, "Math Status")),
        "{{SIGNER_NAME}}": signerName
      };

      appendTemplateLetter_(packetBody, templateBody, replacements, {
  logoFileId: settings["Logo File ID"]
});

studentsProcessed++;

if (studentIndex < students.length - 1) {
  packetBody.appendPageBreak();
}
});
    
packetDoc.saveAndClose();
packetsCreated++;

updatePacketProgress_(
  packetsCreated,
  Object.keys(teacherGroups).length,
  studentsProcessed,
  skippedCount
);
SpreadsheetApp.flush();




/*
V2 IDEA:
Generate teacher packets as PDFs.

Current implementation creates PDFs successfully,
but formatting/layout differs slightly from the
Google Doc version. Leaving disabled until a
reliable PDF workflow is finalized.

To re-enable later:
const packetFile = DriveApp.getFileById(packetDoc.getId());
const pdfBlob = packetFile.getBlob().getAs(MimeType.PDF).setName(docName + ".pdf");
outputFolder.createFile(pdfBlob);
*/

// const pdfBlob = packetFile.getBlob().getAs(MimeType.PDF).setName(pdfName);
// outputFolder.createFile(pdfBlob);

if (options.markGenerated) {
  const generatedCol = headerMap["Letter Processed"] + 1;
  students.forEach(studentObj => {
    processedSheet.getRange(studentObj.sheetRow, generatedCol).setValue(true);
  });
}

logRun_("Teacher Packet Generated", `Created ${docName} with ${students.length} letters.`);
}

if (hasUnprocessedTeachers_()) {
  updatePacketProgress_(
    packetsCreated,
    Object.keys(teacherGroups).length,
    studentsProcessed,
    skippedCount
  );

  const detailRange = ss.getRangeByName("Status_PacketsCreated_Detail");
  if (detailRange) {
    detailRange.setValue(
      "Working… " +
      studentsProcessed +
      " letters generated in this batch. More to come."
    );
  }
} else {
  updatePacketStatus_(packetsCreated, studentsProcessed, skippedCount);
}

return {
  packetsCreated,
  studentsProcessed,
  skippedCount,
  teacherNamesProcessed: teacherNames.length
};

} catch (error) {

  setPacketErrorStatus_(error);

  Logger.log(error);

  throw error;
}
}

function hasUnprocessedTeachers_() {
  const ss = getSpreadsheet_();
  const processedSheet = ss.getSheetByName("Processed Data");
  const data = processedSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const headerMap = createHeaderMap_(headers);

  return rows.some(row => {
    const generate = getValue_(row, headerMap, "Generate Letter");
    const processed = getValue_(row, headerMap, "Letter Processed");

    const shouldGenerate = (generate === true || String(generate).toLowerCase() === "true");
    const notYetProcessed = !(processed === true || String(processed).toLowerCase() === "true");

    return shouldGenerate && notYetProcessed;
  });
}

function testHasUnprocessed() {
  Logger.log("Unprocessed teachers remaining? " + hasUnprocessedTeachers_());
}

/*************************************************
 * TESTING / DIAGNOSTICS
 *************************************************/
 function testLogoAccess() {
  const settings = getSettings_();

  const logoFileId = settings["Logo File ID"];

  Logger.log("Logo File ID: " + logoFileId);

  const file = DriveApp.getFileById(logoFileId);

  Logger.log("Logo Name: " + file.getName());
  Logger.log("Mime Type: " + file.getMimeType());
}

function testDriveAccess() {
  const folders = DriveApp.getFolders();
  Logger.log(folders.hasNext());
}
function showSettings() {
  const settings = getSettings_();

  Logger.log(settings["Template Doc ID"]);
  Logger.log(settings["Upload Folder ID"]);
  Logger.log(settings["Output Folder ID"]);
}
function testTemplateAndOutputAccess() {
  const settings = getSettings_();

  const templateDocId = settings["Template Doc ID"];
  const outputFolderId = settings["Output Folder ID"];

  Logger.log("Template Doc ID: " + templateDocId);
  Logger.log("Output Folder ID: " + outputFolderId);

  const templateFile = DriveApp.getFileById(templateDocId);
  Logger.log("Template file name: " + templateFile.getName());

  const outputFolder = DriveApp.getFolderById(outputFolderId);
  Logger.log("Output folder name: " + outputFolder.getName());
}
function testLetterSetup() {
  const settings = getSettings_();

  Logger.log("Template Doc ID: " + settings["Template Doc ID"]);
  Logger.log("Output Folder ID: " + settings["Output Folder ID"]);

  const processedSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Processed Data");

  const data = processedSheet.getDataRange().getValues();

  Logger.log("Rows: " + data.length);
  Logger.log("Headers: " + data[0].join(", "));
}

function testLastRunLocation() {
  const ss = getSpreadsheet_();
  const range = ss.getRangeByName("Status_LastRun");

  Logger.log(range.getSheet().getName());
  Logger.log(range.getA1Notation());
}

function testPacketDetail() {
  const ss = getSpreadsheet_();

  ss.getRangeByName("Status_PacketsCreated_Detail")
    .setValue("TESTING PACKET DETAIL TEXT");
}

function testProgressBar() {
  updatePacketProgress_(8, 14, 173, 11);
}

function testProgressRanges() {
  const ss = getSpreadsheet_();

  [
    "Status_Packets_ProgressBar",
    "Status_Packets_ProgressPercent"
  ].forEach(name => {
    const range = ss.getRangeByName(name);
    Logger.log(name + ": " + (range ? range.getSheet().getName() + "!" + range.getA1Notation() : "MISSING"));
  });
}

function testOneClickNow() {
  generateParentLettersOneClick();
}

function diagnoseSchoolName() {
  const ss = getSpreadsheet_();
  Logger.log("Spreadsheet: " + (ss ? ss.getName() : "NULL"));
  const range = ss.getRangeByName("School_Name");
  Logger.log("School_Name range: " + (range ? range.getA1Notation() : "NULL"));
  if (range) {
    Logger.log("Value: " + range.getValue());
  }
}

function diagnoseRanges() {
  const ss = getSpreadsheet_();
  ["School_Name", "Signer_Name", "Letter_Date"].forEach(name => {
    const r = ss.getRangeByName(name);
    Logger.log(name + ": " + (r ? r.getA1Notation() : "NULL"));
  });
}

function diagnoseAssessment() {
  const ss = getSpreadsheet_();
  const r = ss.getRangeByName("Assessment_Window");
  Logger.log("Assessment_Window: " + (r ? r.getA1Notation() : "NULL"));
  if (r) Logger.log("Value: " + r.getValue());
}

/*************************************************
 * LETTER HELPERS
 *************************************************/

 function windowToReportSeason_(window) {
  if (window === "BOY") return "Fall";
  if (window === "MOY") return "Winter";
  if (window === "EOY") return "Spring";
  return window;
}
function titleCase_(text) {
  return String(text)
    .toLowerCase()
    .replace(/\b\w/g, char => char.toUpperCase());
}

function escapeRegExp_(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function appendTemplateLetter_(targetBody, templateBody, replacements, options) { 
   const numChildren = templateBody.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const child = templateBody.getChild(i).copy();
    replaceInElement_(child, replacements);
    if (options && options.logoFileId) {
  replaceLogoPlaceholder_(child, options.logoFileId);
}

    const type = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      targetBody.appendParagraph(child.asParagraph());
    } else if (type === DocumentApp.ElementType.TABLE) {
      targetBody.appendTable(child.asTable());
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      targetBody.appendListItem(child.asListItem());
    } else if (type === DocumentApp.ElementType.PAGE_BREAK) {
      targetBody.appendPageBreak();
    } else if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
      targetBody.appendHorizontalRule();
    }
  }
}

function replaceInElement_(element, replacements) {
  if (element.getType && element.getType() === DocumentApp.ElementType.TEXT) {
    const text = element.asText();
    Object.keys(replacements).forEach(key => {
      text.replaceText(escapeRegExp_(key), String(replacements[key] ?? ""));
    });
    return;
  }

  if (element.getNumChildren) {
    for (let i = 0; i < element.getNumChildren(); i++) {
      replaceInElement_(element.getChild(i), replacements);
    }
  }
}

function deleteExistingFilesByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

function getLastName_(fullName) {
  const parts = String(fullName).trim().split(/\s+/);
  return parts[parts.length - 1] || "";
}

function replaceLogoPlaceholder_(body, logoFileId) {
  if (!logoFileId) return;

  const found = body.findText(escapeRegExp_("{{SCHOOL_LOGO}}"));
  if (!found) return;

  const textElement = found.getElement().asText();
  const start = found.getStartOffset();
  const end = found.getEndOffsetInclusive();

  textElement.deleteText(start, end);

  const logoBlob = DriveApp.getFileById(logoFileId).getBlob();

  const parent = textElement.getParent();
  const image = parent.asParagraph().insertInlineImage(0, logoBlob);

  image.setWidth(55);
  image.setHeight(55);
}
/*************************************************
 * DASHBOARD HELPERS
 *************************************************/
function updateDashboardLinks() {
  const ss = getSpreadsheet_();
  const settings = getSettings_();

  const uploadCell = ss.getRangeByName("QuickLink_UploadFolder");
  const outputCell = ss.getRangeByName("QuickLink_OutputFolder");
  

  const uploadFolderId = settings["Upload Folder ID"];
  const outputFolderId = settings["Output Folder ID"];

  uploadCell.setFormula(
    '=HYPERLINK("https://drive.google.com/drive/folders/' + uploadFolderId + '","Open Acadience Uploads Folder")'
  );

  outputCell.setFormula(
    '=HYPERLINK("https://drive.google.com/drive/folders/' + outputFolderId + '","Open Generated Letters Folder")'
  
  );
}
function openAcadienceUploadsFolder() {
  const settings = getSettings_();

  const folderId = settings["Upload Folder ID"];

  const url = "https://drive.google.com/drive/folders/" + folderId;

  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '", "_blank");google.script.host.close();</script>'
  );

  SpreadsheetApp.getUi().showModalDialog(html, "Opening Folder...");
}

function setDashboardStatus_(statusName, icon, text) {
  const ss = getSpreadsheet_();

  const statusMap = {
    "Acadience File Found": {
      icon: "Status_FileFound_Icon",
      text: "Status_FileFound_Text"
    },
    "Teacher Packets Created": {
      icon: "Status_PacketsCreated_Icon"
    },
    "Last Run": {
      text: "Status_LastRun"
    }
  };

  const target = statusMap[statusName];
  if (!target) return;

  if (target.icon && icon !== null && icon !== undefined) {
    safeSetNamedRange_(target.icon, icon);
  }

  if (target.text && text !== undefined) {
    safeSetNamedRange_(target.text, text);
  }
}
function testDashboardNamedRanges() {
  const ss = getSpreadsheet_();

  Logger.log("Assessment Window: " + ss.getRangeByName("Assessment_Window").getValue());
  Logger.log("Letter Date: " + ss.getRangeByName("Letter_Date").getDisplayValue());
  Logger.log("School Name: " + ss.getRangeByName("School_Name").getValue());
  Logger.log("Signer Name: " + ss.getRangeByName("Signer_Name").getValue());
}

function testDashboardStatus() {
  setDashboardStatus_("Acadience File Found", "✓", "Ready");
  getSpreadsheet_().getRangeByName("Status_File_Pill").setValue("READY");

  setDashboardStatus_("Data Imported", "✓", "Imported");
  setDashboardStatus_("Data Processed", "✓", "Processed");

  setDashboardStatus_("Teacher Packets Created", "⏳", "Pending");
  getSpreadsheet_().getRangeByName("Status_Packets_Pill").setValue("IN PROGRESS");

  setDashboardStatus_("Last Run", null, new Date());
  getSpreadsheet_().getRangeByName("Status_LastRun_Icon").setValue("◷");
}

function testStatusNamedRanges() {
  const ss = getSpreadsheet_();

  const names = [
    "Status_FileFound_Icon",
    "Status_FileFound_Text",
    "Status_DataImported_Icon",
    "Status_DataImported_Text",
    "Status_DataProcessed_Icon",
    "Status_DataProcessed_Text",
    "Status_PacketsCreated_Icon",
    "Status_PacketsCreated_Detail",
    "Status_LastRun",
    "Status_File_Pill",
    "Status_Packets_Pill",
    "Status_LastRun_Icon"
  ];

  names.forEach(name => {
    const range = ss.getRangeByName(name);
    Logger.log(name + ": " + (range ? "FOUND" : "MISSING"));
  });
}

function updatePacketStatus_(packetsCreated, studentsProcessed, skippedCount) {
  const ss = getSpreadsheet_();

  const setVal = (name, val) => {
    const r = ss.getRangeByName(name);
    if (r) r.setValue(val);
  };
  const setForm = (name, formula) => {
    const r = ss.getRangeByName(name);
    if (r) r.setFormula(formula);
  };

  setVal("Status_PacketsCreated_Icon", "✓");
  setVal("Status_Packets_Pill", "COMPLETE");

  setVal("Status_PacketsCreated_Detail",
    packetsCreated + " teacher packet(s) generated • " +
    studentsProcessed + " student(s) processed" +
    (skippedCount ? " • " + skippedCount + " skipped" : ""));

  setForm("Status_Packets_ProgressBar",
    '=SPARKLINE(1, {"charttype","bar";"max",1;"color1","#1A73E8"})');

  setVal("Status_Packets_ProgressPercent", "100%");

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "M/d/yyyy h:mm a"
  );

  setVal("Status_LastRun", timestamp);
  setVal("Status_LastRun_Icon", "◷");
}

function updatePacketProgress_(packetsCreated, totalPackets, studentsProcessed, skippedCount) {
  const ss = getSpreadsheet_();

  const setVal = (name, val) => {
    const r = ss.getRangeByName(name);
    if (r) r.setValue(val);
  };
  const setForm = (name, formula) => {
    const r = ss.getRangeByName(name);
    if (r) r.setFormula(formula);
  };

  setVal('Status_PacketsCreated_Icon', "⏳");
  setVal("Status_Packets_Pill", "IN PROGRESS");
  setVal('Status_PacketsCreated_Detail',
    packetsCreated + ' of ' + totalPackets + ' teacher packet(s) created • ' +
    studentsProcessed + ' student(s) processed so far' +
    (skippedCount ? ' • ' + skippedCount + ' skipped' : ''));

  const percent = totalPackets ? packetsCreated / totalPackets : 0;

  setForm("Status_Packets_ProgressBar",
    '=SPARKLINE(' + percent + ', {"charttype","bar";"max",1;"color1","#1A73E8"})');
  setVal("Status_Packets_ProgressPercent", Math.round(percent * 100) + "%");
}

function resetPacketStatus_() {
  const ss = getSpreadsheet_();

  const setVal = (name, val) => {
    const r = ss.getRangeByName(name);
    if (r) r.setValue(val);
  };
  const clear = (name) => {
    const r = ss.getRangeByName(name);
    if (r) r.clearContent();
  };

  setVal("Status_PacketsCreated_Icon", "⏳");
  setVal("Status_Packets_Pill", "IN PROGRESS");

  setVal("Status_PacketsCreated_Detail", "");
  clear("Status_Packets_ProgressBar");
  setVal("Status_Packets_ProgressPercent", "");

  setVal("Status_LastRun", "");
  setVal("Status_LastRun_Icon", "");
}

function setPacketErrorStatus_(error) {
  trySetNamedRange_("Status_PacketsCreated_Icon", "❌");
  trySetNamedRange_("Status_Packets_Pill", "ERROR");
  trySetNamedRange_("Status_PacketsCreated_Detail", String(error));
  Logger.log("Packet error: " + error);
}

function appendSkippedStudent_(row, headerMap, reason) {
  const ss = getSpreadsheet_();
  const skippedSheet = ss.getSheetByName("Skipped Students");

  skippedSheet.appendRow([
    getValue_(row, headerMap, "Student Name"),
    getValue_(row, headerMap, "Teacher Name"),
    getValue_(row, headerMap, "Grade"),
    getValue_(row, headerMap, "Assessment Window"),
    getValue_(row, headerMap, "Reading Score"),
    getValue_(row, headerMap, "Reading Status"),
    getValue_(row, headerMap, "Math Score"),
    getValue_(row, headerMap, "Math Status"),
    reason,
    new Date()
  ]);
}

function openHelpSheet() {
  const ss = getSpreadsheet_();
  const helpSheet = ss.getSheetByName("Help");

  if (!helpSheet) {
    SpreadsheetApp.getUi().alert(
      'Help sheet not found. Please verify the sheet is named "Help".'
    );
    return;
  }

  helpSheet.activate();
}
function openDashboard() {
  const ss = getSpreadsheet_();
  const dashboardSheet = ss.getSheetByName("Dashboard");

  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert(
      'Dashboard sheet not found.'
    );
    return;
  }

  dashboardSheet.activate();
}

function openALO() {
  const url = "https://alo.acadiencelearning.org/login";

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 18px; text-align: center;">
      <p style="font-size: 14px; margin-bottom: 16px;">
        Click the button below to open Acadience Learning Online.
      </p>

      <a href="${url}" target="_blank"
         style="
           display: inline-block;
           padding: 10px 18px;
           background: #5E35B1;
           color: white;
           text-decoration: none;
           border-radius: 8px;
           font-weight: bold;
         ">
        Open ALO
      </a>
    </div>
  `)
  .setWidth(360)
  .setHeight(160);

  SpreadsheetApp.getUi().showModalDialog(html, "Open Acadience Learning Online");
}

function saveWizardSettings() {
  const ss = getSpreadsheet_();

  const wizardSheet = ss.getSheetByName("Set-up Wizard"); // change if needed

  if (!wizardSheet) {
    SpreadsheetApp.getUi().alert('Set-up Wizard sheet not found.');
    return;
  }

  const uploadInput = ss.getRangeByName("Setup_UploadFolder").getDisplayValue().trim();
  const outputInput = ss.getRangeByName("Setup_OutputFolder").getDisplayValue().trim();  const templateInput = ss.getRangeByName("Setup_TemplateDoc").getDisplayValue().trim();
  const uploadFolderId = extractDriveFolderId_(uploadInput);
  const outputFolderId = extractDriveFolderId_(outputInput);
  const templateDocId = extractGoogleDocId_(templateInput);

  if (!uploadFolderId || !outputFolderId || !templateDocId) {
    SpreadsheetApp.getUi().alert(
      "Please enter a valid Upload Folder, Output Folder, and Parent Letter Template URL or ID before saving."
    );
    return;
  }

  saveSettingValue_("Upload Folder ID", uploadFolderId);
  saveSettingValue_("Output Folder ID", outputFolderId);
  saveSettingValue_("Template Doc ID", templateDocId);

  SpreadsheetApp.getUi().alert("Settings saved successfully.");
}

function extractDriveFolderId_(input) {
  if (!input) return "";

  const text = String(input).trim();

  const folderMatch = text.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch) return folderMatch[1];

  return text;
}

function extractGoogleDocId_(input) {
  if (!input) return "";

  const text = String(input).trim();

  const docMatch = text.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
  if (docMatch) return docMatch[1];

  return text;
}

function saveSettingValue_(settingName, settingValue) {
  const ss = getSpreadsheet_();
  const settingsSheet = ss.getSheetByName("Settings");

  if (!settingsSheet) {
    throw new Error("Settings sheet not found.");
  }

  const data = settingsSheet.getDataRange().getValues();

  for (let r = 0; r < data.length; r++) {
    if (String(data[r][0]).trim() === settingName) {
      settingsSheet.getRange(r + 1, 2).setValue(settingValue);
      return;
    }
  }

  throw new Error('Setting not found: ' + settingName);
}

function openUploadFolder() {
  const setup = requireSetupComplete_();
  openUrl_(`https://drive.google.com/drive/folders/${setup.uploadFolderId}`);
}

function openOutputFolder() {
  const setup = requireSetupComplete_();
  openUrl_(`https://drive.google.com/drive/folders/${setup.outputFolderId}`);
}

function requireSetupComplete_() {
  const setup = getSetupStatus_();

  if (!setup.isComplete) {
    openSetupWizard();
    SpreadsheetApp.getUi().alert(
      "Set-up is incomplete. Please complete the Set-up Wizard before continuing."
    );
    throw new Error("Set-up incomplete.");
  }

  return setup;
}

function getSetupStatus_() {
  const uploadFolderId = getSettingValue_("Upload Folder ID");
  const outputFolderId = getSettingValue_("Output Folder ID");
  const templateDocId = getSettingValue_("Template Doc ID");

  return {
    uploadFolderId,
    outputFolderId,
    templateDocId,
    isComplete: Boolean(uploadFolderId && outputFolderId && templateDocId)
  };
}

function openRunLog() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName("Run Log");

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Run Log sheet was not found.");
    return;
  }

  ss.setActiveSheet(sheet);
}

function openSetupWizard() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName("Set-up Wizard");

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Set-up Wizard sheet was not found.");
    return;
  }

  ss.setActiveSheet(sheet);
}

function openSkippedStudents() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName("Skipped Students");

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Skipped Students sheet was not found.");
    return;
  }

  ss.setActiveSheet(sheet);
}

function getSettingValue_(settingName) {
  const settings = getSettings_();
  return settings[settingName] || "";
}

function openUrl_(url) {
  const html = HtmlService.createHtmlOutput(`
    <script>
      window.open("${url}", "_blank");
      google.script.host.close();
    </script>
    <p>
      If the page did not open,
      <a href="${url}" target="_blank">click here</a>.
    </p>
  `)
    .setWidth(300)
    .setHeight(120);

  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
}

function generateParentLettersOneClick() {
  const ui = SpreadsheetApp.getUi();

  try {

    PropertiesService.getScriptProperties()
      .setProperty("SPREADSHEET_ID", getSpreadsheet_().getId());

    requireSetupComplete_();
    

    resetWorkflowStatus_();

    importLatestCsv();

    processImportedData();
    setDashboardStatus_("Acadience File Found", "✓", "Acadience data processed.");

    generateTeacherPdfs();

    ui.alert(
      "Teacher packet generation complete.\n\n" +
      "Packets have been saved to the Output Folder."
    );

  } catch (error) {
    setWorkflowErrorStatus_(error);

    ui.alert(
      "The Letter Generator could not finish.\n\n" +
      error.message
    );

    throw error;
  }
}
function resetWorkflowStatus_() {
  const ss = getSpreadsheet_();

  setDashboardStatus_("Acadience File Found", "⏳", "Looking for latest Acadience CSV...");
  setDashboardStatus_("Teacher Packets Created", "—");

  safeSetNamedRange_("Status_File_Pill", "CHECKING");
  safeSetNamedRange_("Status_Packets_Pill", "WAITING");

  safeSetNamedRange_("Status_PacketsCreated_Detail", "");
  ss.getRangeByName("Status_Packets_ProgressBar").clearContent();
  safeSetNamedRange_("Status_Packets_ProgressPercent", "");
}

function setWorkflowErrorStatus_(error) {
  safeSetNamedRange_("Status_PacketsCreated_Icon", "❌");
  safeSetNamedRange_("Status_Packets_Pill", "ERROR");
  safeSetNamedRange_("Status_PacketsCreated_Detail", error.message || String(error));
}

function safeSetNamedRange_(rangeName, value) {
  const ss = getSpreadsheet_();
  const range = ss.getRangeByName(rangeName);

  if (!range) {
    throw new Error("Missing named range: " + rangeName);
  }

  range.setValue(value);
}

function trySetNamedRange_(rangeName, value) {
  const ss = getSpreadsheet_();
  const range = ss.getRangeByName(rangeName);
  if (range) {
    range.setValue(value);
  }
}