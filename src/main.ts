import { Album } from './album';

/**
 * @todo Move to custom @types files plus additions made to the Apps Script types.
 */
interface SheetsFormSubmitEvent {
  authMode: GoogleAppsScript.Script.AuthMode;
  namedValues: { [key: string]: string[] };
  range: GoogleAppsScript.Spreadsheet.Range;
  triggerUid: number;
  values: string[];
}

/**
 * Creates the 'Album Collaborative' menu and calculates the Summary sheet
 * when the spreadsheet is opened. This function is called through a built-in apps-script trigger.
 */
function onOpen() {
  // Create menu within Spreadsheet.
  createMenu();

  // Calculate Summary sheet.
  calculate();
}

/**
 * Creates the 'Album Collaborative' menu for the opened speadsheet.
 */
function createMenu() {
  const ui = SpreadsheetApp.getUi();

  // create the 'Album Collaborative' menu
  ui.createMenu('Album Collaborative')
    .addItem('Next', 'next')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Utilities')
        .addItem('Back', 'back')
        .addItem('Calculate', 'calculate')
        .addItem('Generate', 'generate')
        .addItem('New Album', 'newAlbum'),
    )
    .addToUi();
}

/**
 * Creates a new album Form and formats the subsequently created form-linked spreadsheet.
 * The album is also added to the Summary sheet.
 * @param submitter The name of the submitter of the album.
 */
function newAlbum(submitter?: string) {
  const album = new Album();

  // Prompt the user for the album's info.
  if (!album.prompt(submitter)) {
    return;
  }

  // Create the form and format its sheet.
  const form = createForm(album);

  // Add the album to the Summary sheet and calculate it.
  addToSummarySheet(album, form);

  // Add the album to the Current Album sheet.
  addToCurrentAlbumSheet(album, form);
}

/**
 * Creates a new form for the specified album and formats the created sheet.
 * @param album The album to create a form for.
 * @returns The created form.
 */
function createForm(album: Album): GoogleAppsScript.Forms.Form {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const form = FormApp.create(album.formattedName);

  // Create a new form for the album.
  form
    .addScaleItem()
    .setTitle('Authentic')
    .setHelpText(
      `Emotions are real, genuine and truthful; creating good music for the sake of the music
 itself.`,
    )
    .setBounds(1, 5)
    .setRequired(true)

    .duplicate()
    .setTitle('Adventurous')
    .setHelpText(
      `The artist/band looks for new ways to express what they feel and have to communicate;
 the surprise element, the creativity, the musical vision.`,
    )

    .duplicate()
    .setTitle('Accurate')
    .setHelpText(
      `A "Yes, that's it!" reaction; the translation of feelings through the mastery of an
 instrument.`,
    )

    .duplicate()
    .setTitle('Artistic')
    .setHelpText(
      `The more cerebral aspect of music; a concept which leads to structure, balance, length,
 interplay, selection of instruments, etc.`,
    )

    .duplicate()
    .setTitle('Attention-grabbing')
    .setHelpText(
      `Music should require some effort from the listener, but it should also include a
 factor of entertainment; keeps the listener's attention.`,
    );

  form
    .addScaleItem()
    .setTitle('Overall')
    .setBounds(1, 10)
    .setRequired(true);

  form
    .addParagraphTextItem()
    .setTitle('Favorite song(s)')

    .duplicate()
    .setTitle('Analysis')
    .setRequired(true)
    .setValidation(
      FormApp.createParagraphTextValidation()
        .build(),
    );

  form
    .setAllowResponseEdits(true)
    .setDescription(`Submitted by ${album.submitter}.`)
    .setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId())
    .setLimitOneResponsePerUser(true)
    .setPublishingSummary(true);

  // Format the newly created form sheet.
  formatFormSheet(album);

  return form;
}

/**
 * Formats a newly created form sheet.
 * @param album The album used to create the form.
 */
function formatFormSheet(album: Album) {
  // Make sure the subsequently created form has been added to the spreadsheet before we begin.
  SpreadsheetApp.flush();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheet = sheets[0];

  // Format the sheet.
  sheet
    .activate()
    .setName(album.formattedName)
    .setColumnWidth(1, 200)
    .setColumnWidths(2, 6, 100)
    .setColumnWidths(8, 2, 300)
    .deleteColumns(10, sheet.getMaxColumns() - 9);

  const all = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  all
    .setBorder(true, true, true, true, true, true)
    .setFontSize(10)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setWrap(true)
    .applyRowBanding(SpreadsheetApp.BandingTheme.TEAL, true, false);

  const header = sheet.getDataRange();
  header
    .setBorder(
      null,
      null,
      true,
      null,
      null,
      null,
      null,
      SpreadsheetApp.BorderStyle.DOUBLE,
    )
    .setFontSize(11)
    .setFontWeight('bold')
    .createFilter();

  sheet.getRange('A:A').setNumberFormat('mmmm d, yyyy');

  // Move sheet to end.
  spreadsheet.moveActiveSheet(sheets.length);
}

/**
 * Adds an album and it's referenced form to the Summary sheet.
 * @param album The album to add to the Summary sheet.
 * @param form The form to reference within the Summary sheet.
 */
function addToSummarySheet(album: Album, form: GoogleAppsScript.Forms.Form) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = spreadsheet.getSheetByName('Summary');
  const timestamp = new Date();

  // Add new row with album/form data.
  summarySheet
    .activate()
    .appendRow([
      timestamp,
      album.title,
      album.artist,
      album.submitter,
      form.shortenFormUrl(form.getPublishedUrl()),
      album.spotifyUrl,
    ])
    .getRange(summarySheet.getLastRow(), 1)
    .setNumberFormat('mmmm d, yyyy');

  // Calculate the Summary sheet.
  calculate();
}

/**
 * Adds an album and it's referenced form to the Current Album sheet.
 * @param album The album to add to the Current Album sheet.
 * @param form The form to reference with the Current Album sheet.
 */
function addToCurrentAlbumSheet(
  album: Album,
  form: GoogleAppsScript.Forms.Form,
) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const currentAlbumSheet = spreadsheet.getSheetByName('Current Album');
  const sheetRange = currentAlbumSheet.getRange('C2:C5');
  const sheetValues = sheetRange.getValues();

  currentAlbumSheet.activate();

  sheetValues[0][0] = album.formattedName;
  sheetValues[1][0] = album.submitter;
  sheetValues[2][0] = form.shortenFormUrl(form.getPublishedUrl());
  sheetValues[3][0] = album.spotifyUrl;

  sheetRange.setValues(sheetValues);

  currentAlbumSheet.autoResizeColumn(3);
}

/**
 * Submit function called when a form response has been submitted.
 * This function will style the newly added response and calculate the Summary sheet.
 * @param e The form submitted event.
 */
function submit(e: SheetsFormSubmitEvent) {
  const range = e.range;

  if (range.getNumColumns() < 9) {
    return;
  }

  // Style the added range.
  range
    .setBorder(true, true, true, true, true, true)
    .setFontSize(10)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('top')
    .setWrap(true)
    .getCell(1, 1)
    .setNumberFormat('mmmm d, yyyy h:mm am/pm');

  range.offset(0, 7, 1, 2).setHorizontalAlignment('left');

  // Calculate the Summary sheet.
  calculate();
}

/**
 * Calculates the Summary sheet.
 * @todo Pass in a value to specify which row to calculate.
 * @todo Make calculations more dynamic; less magic numbers.
 */
function calculate() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = spreadsheet.getSheetByName('Summary');
  const sheetRange = summarySheet.getRange(
    2,
    2,
    summarySheet.getLastRow() - 1,
    summarySheet.getLastColumn() - 1,
  );
  const sheetValues = sheetRange.getValues();
  let sheet;
  let count;

  for (let i = 0; i < sheetValues.length; i += 1) {
    sheet = spreadsheet.getSheetByName(
      `${sheetValues[i][0]} — ${sheetValues[i][1]}`,
    );
    count = sheet.getLastRow() - 1;

    sheetValues[i][5] = count;

    for (let j = 6; j < sheetValues[i].length; j += 1) {
      sheetValues[i][j] =
        count < 1 ? 'TBD' : getAverageForSheetColumn(sheet, count, j - 4);
    }
  }

  sheetRange.setValues(sheetValues);
}

/**
 * Determines the average for a given sheet's column.
 * @param sheet The sheet to determine the average of the specified column.
 * @param count The total number of sheets to divide the value by.
 * @param column The column to find the average of.
 */
function getAverageForSheetColumn(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  count: number,
  column: number,
): string {
  let values;
  let sum = 0;

  values = sheet.getRange(2, column, count, 1).getValues();
  for (let i = 0; i < values.length; i += 1) {
    sum += values[i][0] as number;
  }

  return (sum / count).toPrecision(2);
}

/**
 * Generates a new order of submitters.
 */
function generate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    'Schedule',
  );
  const namesRange = sheet.getRange('D2:D8');
  const names = namesRange.getValues();

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Generate',
    'Generate a new order?',
    ui.ButtonSet.YES_NO,
  );

  if (response !== ui.Button.YES) {
    return;
  }

  for (let i = names.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    const x = names[i];
    names[i] = names[j];
    names[j] = x;
  }

  namesRange.setValues(names);
}

/**
 * Moves the pointer to the next submitter.
 */
function next() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Schedule')
    .activate();
  const range = sheet.getRange('B2:B8');
  const values = range.getValues();
  let i;

  for (i = 0; values[i][0] !== '->'; i += 1);

  values[i][0] = '';
  i += 1;
  if (i === values.length) {
    i = 0;
    generate();
  }

  values[i][0] = '->';
  range.setValues(values);

  newAlbum(
    range
      .offset(0, 2)
      .getCell(i + 1, 1)
      .getValue() as string,
  );
}

/**
 * Moves the pointer to the previous submitter.
 */
function back() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Schedule')
    .activate();
  const range = sheet.getRange('B2:B8');
  const values = range.getValues();
  let i;

  for (i = 0; values[i][0] !== '->'; i += 1);

  values[i][0] = '';
  i -= 1;
  if (i < 0) {
    values[values.length - 1][0] = '->';
  } else {
    values[i][0] = '->';
  }

  range.setValues(values);
}
