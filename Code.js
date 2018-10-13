function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Album Collaborative')
  .addItem('Next', 'next')
  .addSeparator()
  .addSubMenu(ui.createMenu('Utilities')
              .addItem('Back', 'back')
              .addItem('Calculate', 'calculate')
              .addItem('Generate', 'generate')
              .addItem('New Album', 'newAlbum'))
  .addToUi();
  
  calculate();
}

function newAlbum(submitter) {
  var timestamp = new Date();
  var ui = SpreadsheetApp.getUi();
  
  var titleResponse = ui.prompt('New Album', 'Enter the album\'s title.', ui.ButtonSet.OK_CANCEL);
  if (titleResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var artistResponse = ui.prompt('New Album', 'Enter the album\'s artist.', ui.ButtonSet.OK_CANCEL);
  if (artistResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  if (!submitter) {
    var submitterResponse = ui.prompt('New Album', 'Enter the album\'s submitter.', ui.ButtonSet.OK_CANCEL);
    if (submitterResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    submitter = submitterResponse.getResponseText();
  }
  
  var title = titleResponse.getResponseText();
  var artist = artistResponse.getResponseText();
  var album = title + ' — ' + artist;
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.create(album);
  
  form.addScaleItem()
  .setTitle('Authentic')
  .setHelpText('The emotions have to be real, genuine and truthful. ' +
               'The prime objective should be to create good music for the sake of the music itself.')
  .setBounds(1, 5)
  .setRequired(true)
  
  .duplicate()
  .setTitle('Adventurous')
  .setHelpText('The artist/band should be looking for new ways to express what they feel and have to communicate. ' +
               'The surprise element, the creativity, the musical vision are part of the adventure.')
  
  .duplicate()
  .setTitle('Accurate')
  .setHelpText('A "Yes, that\'s it!" reaction. A sublime translation of feelings through the skills and mastery of a instrument.')
  
  .duplicate()
  .setTitle('Artistic')
  .setHelpText('The more cerebral aspect of music. Some concept which leads to structure, balance, length, interplay, selection of instruments, of musicians, of new approaches.')
  
  .duplicate()
  .setTitle('Attention-grabbing')
  .setHelpText('Though music can and should require an effort from the listener, it should also include a factor of entertainment. ' +
               'In the sense of keeping the attention going, of being captivating.');
  
  form.addScaleItem()
  .setTitle('Overall')
  .setBounds(1, 10)
  .setRequired(true);
  
  form.addParagraphTextItem()
  .setTitle('Favorite song(s)')
  
  .duplicate()
  .setTitle('Analysis')
  .setRequired(true)
  .setValidation(
    FormApp.createParagraphTextValidation()
    .requireTextLengthGreaterThanOrEqualTo(500)
  )
  .setHelpText('Your analysis must be at least 500 characters long.');
  
  form.setAllowResponseEdits(true)
  .setDescription('Submitted by ' + submitter + '.')
  .setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId())
  .setLimitOneResponsePerUser(true)
  .setPublishingSummary(true);
  
  SpreadsheetApp.flush();
  
  var sheets = spreadsheet.getSheets();
  var sheet = sheets[0];
  
  sheet.activate()
  .setName(album)
  .setColumnWidth(1, 200)
  .setColumnWidths(2, 6, 100)
  .setColumnWidths(8, 2, 300)
  .deleteColumns(10, sheet.getMaxColumns() - 9);
  
  var all = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  all.setBorder(true, true, true, true, true, true)
  .setFontSize(10)
  .setFontWeight('normal')
  .setHorizontalAlignment('center')
  .setWrap(true)
  .applyRowBanding(SpreadsheetApp.BandingTheme.TEAL, true, false);
  
  var header = sheet.getDataRange();
  header.setBorder(null, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE)
  .setFontSize(11)
  .setFontWeight('bold')
  .createFilter();
  
  sheet.getRange('A:A').setNumberFormat('mmmm d, yyyy');
  
  spreadsheet.moveActiveSheet(sheets.length);
  
  var summary = spreadsheet.getSheetByName('Summary');
  summary.activate()
  .appendRow([
    timestamp,
    title,
    artist,
    submitter,
    form.shortenFormUrl(form.getPublishedUrl())
  ])
  .getRange(summary.getLastRow(), 1).setNumberFormat('mmmm d, yyyy');
  
  calculate();
}

function submit(e) {
  var range = e.range;
  
  if (range.getNumColumns() < 9) {
    return;
  }
  
  range.setBorder(true, true, true, true, true, true)
  .setFontSize(10)
  .setFontWeight('normal')
  .setHorizontalAlignment('center')
  .setVerticalAlignment('top')
  .setWrap(true)
  .getCell(1, 1)
  .setNumberFormat('mmmm d, yyyy h:mm am/pm');
  
  range.offset(0, 7, 1, 2)
  .setHorizontalAlignment('left');
  
  calculate();
}

function calculate() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var summary = spreadsheet.getSheetByName('Summary');
  var sheetRange = summary.getRange(2, 2, summary.getLastRow() - 1, 11);
  var sheetValues = sheetRange.getValues();
  var sheet;
  var count;
  var i, j;
  
  for (i = 0; i < sheetValues.length; i++) {
    sheet = spreadsheet.getSheetByName(sheetValues[i][0] + ' — ' + sheetValues[i][1]);
    count = sheet.getLastRow() - 1;
    
    sheetValues[i][4] = count;
    
    for (j = 5; j < sheetValues[i].length; j++) {
      sheetValues[i][j] = count < 1 ? 'TBD' : getAverageForSheetColumn(sheet, count, j - 3);
    }
  }
  
  sheetRange.setValues(sheetValues);
}

function getAverageForSheetColumn(sheet, count, column) {
  var values;
  var sum = 0;
  var i;
  
  values = sheet.getRange(2, column, count, 1).getValues();
  for (i = 0; i < values.length; i++) {
    sum += values[i][0];
  }
  
  return (sum / count).toPrecision(2);
}

function generate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule');
  var namesRange = sheet.getRange('D2:D8');
  var names = namesRange.getValues();
  var j, x, i;
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Generate', 'Generate a new order?', ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  for (i = names.length - 1; i > 0; i--) {
    j = Math.floor(Math.random() * (i + 1));
    x = names[i];
    names[i] = names[j];
    names[j] = x;
  }
  
  namesRange.setValues(names);
}

function next() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule').activate();
  var range = sheet.getRange('B2:B8');
  var values = range.getValues();
  var index;
  var i;
  
  for (i = 0; values[i][0] !== '->'; i++);
  
  values[i++][0] = '';
  if (i === values.length) {
    i = 0;
    generate();
  }
  
  values[i][0] = '->';
  range.setValues(values);
  
  newAlbum(range.offset(0, 2).getCell(i + 1, 1).getValue());
}

function back() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule').activate();
  var range = sheet.getRange('B2:B8');
  var values = range.getValues();
  var index;
  var i;
  
  for (i = 0; values[i][0] !== '->'; i++);
  
  values[i--][0] = '';
  if (i < 0) {
    values[values.length - 1][0] = '->';
  } else {
    values[i][0] = '->';
  }
  
  range.setValues(values);
}
