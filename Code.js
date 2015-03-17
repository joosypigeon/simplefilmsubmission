var FILM_SUBMISSION = "Film Submission",
  ID = "ID",

  // the 3 spreadsheet sheets
  FILM_SUBMISSIONS_SHEET = "Film Submissions",
  TEMPLATE_SHEET = "Templates",
  OPTIONS_SETTINGS_SHEET = "Options & Settings",

  //named ranges
  FESTIVAL_NAME = "Festival Name",
  FESTIVAL_WEBSITE = "Festival Website",
  CLOSE_OF_SUBMISSION = "Close Of Submission",
  EVENT_DATE = "Event Date",
  RELEASE_LINK = "Release Link",
  FIRST_FILM_ID = "First ID",
  CURRENT_FILM_ID = "Current Film ID",
  DAYS_BEFORE_REMINDER = "Days Before Reminder",
  TEST_FIRST_NAME = "Test First Name",
  TEST_LAST_NAME = "Test Last Name",
  TEST_TITLE = "Test Title",
  TEST_EMAIL = "Test Email",
  COLOR_DATA = "Color Data",
  FESTIVAL_DATA = "Festival Data",
  TEMPLATE_DATA = "Template Data",
  TEST_DATA = "Test Data",
  INTERNALS = "Internals",


  SUBMISSION_CONFIRMATION = "Submission Confirmation",
  SUBMISSION_CONFIRMATION_SUBJECT_LINE = "Submission Confirmation Subject Line",
  SUBMISSION_CONFIRMATION_BODY = "Submission Confirmation Body",

  RECEIPT_CONFIMATION = "Receipt Confimation",
  RECEIPT_CONFIMATION_SUBJECT_LINE = "Receipt Confimation Subject Line",
  RECEIPT_CONFIMATION_BODY = "Receipt Confimation Body",

  REMINDER = "Reminder",
  REMINDER_SUBJECT_LINE = "Reminder Subject Line",
  REMINDER_BODY = "Reminder Body",

  NOT_ACCEPTED = "Not Accepted",
  NOT_ACCEPTED_SUBJECT_LINE = "Not Accepted Subject Line",
  NOT_ACCEPTED_BODY = "Not Accepted Body",

  ACCEPTED = "Accepted",
  ACCEPTED_SUBJECT_LINE = "Accepted Subject Line",
  ACCEPTED_BODY = "Accepted Body",

  AD_HOC_EMAIL = "Ad Hoc Email",
  AD_HOC_EMAIL_SUBJECT_LINE = "Ad Hoc Email Subject Line",
  AD_HOC_EMAIL_BODY = "Ad Hoc Email Body",

  SUBJECT_LINE = 'Subject Line',

  ENABLE_REMINDER = "Enable Reminder",
  ENABLE_CONFIRMATION = "Enable Confirmation",
  CURRENT_AD_HOC_EMAIL = "Current Ad Hoc Email",
  CURRENT_SELECTION_NOTIFICATION = "Current Selection Notification",

  SUBJECT = "Subject",
  BODY = "Body",
  AVAILABLE_TAGS = "Available Tags",
  SEND_TEST_EMAIL = "Send Test Email",

  //sufix and prefix for ID
  HELP = "Help",
  LABEL = "Label",
  TEST = "Test",
  BUTTON = "Button",
  HIDDEN = "Hidden",
  PLABEL = "PLable",

  //IDs
  NEXT = "Next",
  PREVIOUS = "Previous",
  SAVE = "Save",
  OK = "OK",
  ENABLE = "Enable",
  UNENABLE = "Unenable",
  CANCEL = "Cancel",
  TEST = "Test",
  WAIT = "Wait",
  STATUS_HTML = "Status HTML",
  STATUS_WARNING_HTML = "Status Warning HTML",
  TEST_PROCESSING_LABEL = "Test Processing Label",
  DATE_DIFF = "Date Diff",
  FESTIVAL_DATA_NAMES = "Festival Data Names",
  TEMPLATE_DATA_NAMES = "Template Data Names",
  TEST_DATA_NAMES = "Test Data Names",
  TAGS = "Tags",

  //value for ENABLE_REMINDER and ENABLE_CONFIRMATION
  ENABLED = "Enabled",
  NOT_ENABLED = "Not Enabled",

  //values for CURRENT_SELECTION_NOTIFICATION, CURRENT_AD_HOC_EMAIL
  NOT_STARTED = "Not Started",
  PENDING = "Pending",

  //states
  NO_MEDIA = "No Media",
  MEDIA_PRESENT = "Media Present",
  PROBLEM = "Problem",
  SELECTED = "Selected",
  NOT_SELECTED = "Not Selected",
  CONFIRMED = "Confirmed",
  NOT_CONFIRMED = "Not Confirmed",

  DO_NOT_CHANGE = "Do Not Change",

  // process names
  SELECTION_NOTIFICATION = "Selection Notification",

  // spreadsheet menu
  MENU_ENTRIES = [{
      name: "Set selected range to: '" + MEDIA_PRESENT + ", " + NOT_CONFIRMED + "'",
      functionName: "mediaPresentNotConfirmed"
    }, {
      name: "Set selected range to: '" + PROBLEM + "'",
      functionName: "problem"
    }, {
      name: "Set selected range to: '" + SELECTED + "'",
      functionName: "selected"
    }, {
      name: "Set selected range to: '" + NOT_SELECTED + "'",
      functionName: "notSelected"
    }, {
      name: "Manually set state of selected range",
      functionName: "manual"
    },
    null, {
      name: "Settings & Options",
      functionName: "settingsOptions"
    }, {
      name: "Templates",
      functionName: "editAndSaveTemplates"
    },
    null, {
      name: AD_HOC_EMAIL,
      functionName: "adHocEmail"
    }, {
      name: SELECTION_NOTIFICATION,
      functionName: "selectionNotification"
    }
  ],

  // form data
  FIRST_NAME = "First Name",
  LAST_NAME = "Last Name",
  EMAIL = "Email",
  TITLE = "Title",
  LENGTH = "Length",
  COUNTRY = "Country",
  YEAR = "Year",
  GENRE = "Genre",
  WEBSITE = "Website",
  SYNOPSIS = "Synopsis",
  CAST_AND_CREW = "Cast & Crew",
  FESTIVAL_SELECTION_AND_AWARDS = "Festival Selection & Awards",

  // for Strawberry Shorts
  BEST_LOCAL_FILM = "Is this film submitted to the Best Local Film Competitive Programme?",
  BEST_LOCAL_FILM_ELIGIBILITY = "If this film is being submitted to The Best Local Film Competitive Programme, please explain its eligibility?",

  CONFIRM = "Confirm",

  // Timestamp column
  TIMESTAMP = "Timestamp",

  // Additional FILM_SUBMISSIONS_SHEET Columns
  COMMENTS = "Comments",
  FILM_ID = "ID",
  SCORE = "Score",
  CONFIRMATION = "Confirmation",
  SELECTION = "Selection",
  STATUS = "Status",
  LAST_CONTACT = "Last Contact",

  FORM_TITLE = "Online Film Submissions Form",
  FORM_RESPONSE = "You now need to download, print off and sign the permission slip. A link to the permission slip has been emailed to you. The permission slip must be mailed together with a DVD of your film, to the address given on the slip. We cannot screen your film without a signed permission slip. If you do not receive this email, please check your spam folder. If you still cannot find the email please email us at " + Session.getActiveUser().getEmail() + " .",
  FORM_DATA = [{
      type: "addTextItem",
      required: {},
      title: FIRST_NAME,
      help: "Please enter your first name"
    }, {
      type: "addTextItem",
      required: {},
      title: LAST_NAME,
      help: "Please enter your last name"
    }, {
      type: "addTextItem",
      required: {},
      title: EMAIL,
      help: "Please enter your email address"
    },
    "section", {
      type: "addTextItem",
      required: {},
      title: TITLE,
      help: "Please give the title of your film"
    }, {
      type: "addTextItem",
      required: {},
      title: LENGTH,
      help: "Please enter the length of the film in format MM:SS"
    }, {
      type: "addTextItem",
      required: {},
      title: COUNTRY,
      help: "Please enter the films country/countries of origin"
    }, {
      type: "addTextItem",
      required: {},
      title: YEAR,
      help: "Please enter the year of the films completion in format YYYY"
    }, {
      type: "addTextItem",
      required: {},
      title: GENRE,
      help: "Please enter the genre of the film"
    }, {
      type: "addTextItem",
      title: WEBSITE,
      help: "Please give film website, if any"
    }, {
      type: "addParagraphTextItem",
      required: {},
      title: SYNOPSIS,
      help: "Please give synopsis of film"
    }, {
      type: "addParagraphTextItem",
      title: CAST_AND_CREW,
      help: "Please give principal cast and crew"
    }, {
      type: "addParagraphTextItem",
      title: FESTIVAL_SELECTION_AND_AWARDS,
      help: "Please give the festivals for which your film has been selected. Also give any awards your film has received"
    }, {
      type: "addCheckboxItem",
      required: {},
      title: CONFIRM,
      help: "I confirm that I have the right to submit this film and that, should it be accepted, I give permission for this film to be screened at the cambridge strawberry shorts film festival 2015.",
      choice: "Confirm"
    }
  ],

  ADDITIONAL_COLUMNS = [{
    title: FILM_ID,
    columnIndex: 1
  }, {
    title: SCORE,
    columnIndex: 1
  }, {
    title: SELECTION,
    columnIndex: 1
  }, {
    title: CONFIRMATION,
    columnIndex: 1
  }, {
    title: STATUS,
    columnIndex: 1
  }, {
    title: LAST_CONTACT,
    columnIndex: 1
  }, {
    title: COMMENTS,
    columnIndex: "end"
  }],

  //columns for TEMPLATE_SHEET
  TEMPLATE_NAME = "Name",
  TEMPLATE = "Template",

  // Sets number of numerical places on film ID so an example of a film ID were the PAD_NUMBER were set to 3 would be ID029.
  // Legal values are 3,4,5,6
  // NOTE: PAD_NUMBER puts an upper bound on the number of film submissions the system can cope with
  PAD_NUMBER = 3,

  // Will stop sending emails for the day when email quota is reported as less than or equal to MIN_QUOTA.
  // Will attempt to send unsent emails the next day.
  MIN_QUOTA = 50,
  REMAINING_EMAIL_QUOTA = "Remaining Email Quota",

  // one contact value
  NO_CONTACT = "No Contact",

  // name of log file
  LOG_FILE = "Submissions Processing Log File",
  LOG_FOLDER = "Log",
  FORM_FOLDER = "Form",

  // String resources
  PROCESSING = 'Processing, please wait . . .',
  SAVE_AND_RETURN_TO = "Save and return To:",

  CACHE = {},

  //map = Array.prototype.map,

  TESTING = false,

  // For Tesing
  TEMPLATES_TESTING = 'Templates Testing';

function flatten(array) {
  var result = [],
    self = arguments.callee;
  array.forEach(function (item) {
    Array.prototype.push.apply(
      result,
      Array.isArray(item) ? self(item) : [item]
    );
  });
  return result;
}

// build custom menu for spreadsheet
function onOpen() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
      initialised = ScriptProperties.getProperty("initialised");

    if(initialised) {
      ss.addMenu(FILM_SUBMISSION, MENU_ENTRIES);
    } else {
      ss.addMenu(FILM_SUBMISSION, [{
        name: "Setup",
        functionName: "setup"
      }]);
    }
  } catch(e) {
    Logger.log("Error in onOpen:" + e);
  }
}

function mediaPresentNotConfirmed() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  statusDialog(MEDIA_PRESENT, NOT_CONFIRMED, DO_NOT_CHANGE, 450, 'Set selection to Media Present, Not Confirmed', '<p>The state of a submission is given by the value of three fields, <b>' + STATUS + '</b>, <b>' + CONFIRMATION + '</b> and <b>' + SELECTION + '</b>.</p><p>Setting a submission to ' + STATUS + ': <b>' + MEDIA_PRESENT + '</b>, ' + CONFIRMATION + ': <b>' + NOT_CONFIRMED + '</b> (and leaving ' + SELECTION + ' unchanged) means the filmmaker will be automatically emailed after midnight to confirm the receipt of the Physical Media and the Permission Slip at the festival office, provided that:<ul><li>close of submission on ' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + ' has not been reached at the time that the email confirmation is attempted</li><li>and that confirmations are currently enabled.</li></ul></p><p>Note: receipt confirmations are currently ' + getNamedValue(ss, ENABLE_CONFIRMATION) + '.</p><p>If you re-set the submission to a new different state before midnight, the receipt confirmation email will not be sent.</p><p>After the receipt confirmation email is sent the state of the submission will be set to ' + STATUS + ': <b>' + MEDIA_PRESENT + '</b>, ' + CONFIRMATION + ': <b>' + CONFIRMED + '</b>.</p><p>To continue with setting the submission state to ' + STATUS + ': <b>' + MEDIA_PRESENT + '</b>, ' + CONFIRMATION + ': <b>' + NOT_CONFIRMED + '</b> press OK, to not do this press CANCEL.</p><br/>');
}

function problem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  statusDialog(PROBLEM, DO_NOT_CHANGE, DO_NOT_CHANGE, 400, 'Set selection to Problem', '<p>The state of a submission is given by the value of three fields, <b>' + STATUS + '</b>, <b>' + CONFIRMATION + '</b> and <b>' + SELECTION + '</b>.</p><p>Setting a submission to ' + STATUS + ': <b>' + PROBLEM + '</b> (and leaving ' + CONFIRMATION + ' and ' + SELECTION + ' unchanged) means the filmmaker will be sent no reminder or receipt confirmation emails from now on. They will however, receive the final selection advisory email after close of submissions(' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + '), when it is sent.</p><p>If you set a submission to ' + STATUS + ': <b>' + PROBLEM + '</b>, please update the submission comment field with the details of the problem.</p><p>To continue with setting the submission state to ' + STATUS + ': <b>' + PROBLEM + '</b> press OK, to not do this press CANCEL.</p><br/>');
}

function selected() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  statusDialog(DO_NOT_CHANGE, DO_NOT_CHANGE, SELECTED, 480, 'Set selection to Selected', '<p>The state of a submission is given by the value of three fields, <b>' + STATUS + '</b>, <b>' + CONFIRMATION + '</b> and <b>' + SELECTION + '</b>.</p><p>A submission set to ' + SELECTION + ': <b>' + SELECTED + '</b> (and leaving ' + STATUS + ' and ' + CONFIRMATION + ' unchanged) will still receive reminder and receipt confirmation emails dependent on its <b>' + STATUS + '</b> and <b>' + CONFIRMATION + '</b> provided that:<ul><li>it is at least ' + getNamedValue(ss, DAYS_BEFORE_REMINDER) + ' days before the close of submission on ' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + ' at the time that the email is attempted</li><li>and that receipt confirmations are currently enabled for receipt confirmation,</li><li>and that reminders are currently enabled for reminders.</li></ul></p><p>Note: receipt confirmations are currently ' + getNamedValue(ss, ENABLE_CONFIRMATION) + '.</p><p>Note: reminders are currently ' + getNamedValue(ss, ENABLE_REMINDER) + '.</p><p>Submissions set to ' + SELECTION + ': <b>' + SELECTED + '</b> will receive the festival submission acceptance email after close of submissions(' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + '), when it is sent.</p><p>You will usually only set this state after the close of submissions.</p><p>To continue with setting the submission state to ' + SELECTION + ': <b>' + SELECTED + '</b> press OK, to not do this press CANCEL.</p><br/>');
}

function notSelected() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  statusDialog(DO_NOT_CHANGE, DO_NOT_CHANGE, NOT_SELECTED, 480, 'Set selection to Not Selected', '<p>The state of a submission is given by the value of three fields, <b>' + STATUS + '</b>, <b>' + CONFIRMATION + '</b> and <b>' + SELECTION + '</b>.</p><p>A submission set to ' + SELECTION + ': <b>' + NOT_SELECTED + '</b> (and leaving ' + STATUS + ' and ' + CONFIRMATION + ' unchanged) will still receive reminder and receipt confirmation emails dependent on its <b>' + STATUS + '</b> and <b>' + CONFIRMATION + '</b> provided that:<ul><li>it is at least ' + getNamedValue(ss, DAYS_BEFORE_REMINDER) + ' days before the close of submission on ' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + ' at the time that the email is attempted</li><li>and that receipt confirmations are currently enabled for receipt confirmation,</li><li>and that reminders are currently enabled for reminders.</li></ul></p><p>Note: receipt confirmations are currently ' + getNamedValue(ss, ENABLE_CONFIRMATION) + '.</p><p>Note: reminders are currently ' + getNamedValue(ss, ENABLE_REMINDER) + '.</p><p>Submissions set to ' + SELECTION + ': <b>' + NOT_SELECTED + '</b> will receive the festival submission not acceptanced email after close of submissions(' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + '), when it is sent.</p><p>You will usually only set this state after the close of submissions.</p><p>To continue with setting the submission state to ' + SELECTION + ': <b>' + NOT_SELECTED + '</b> press OK, to not do this press CANCEL.</p><br/>');
}

function statusDialog(status, confirmation, selection, height, title, html) {
  var value = ScriptProperties.getProperty(normalizeHeader(status + ' ' + confirmation));
  if(value && value === 'true') {
    setStatus(status, confirmation, selection);
  } else {

    var ss = SpreadsheetApp.getActiveSpreadsheet(),
      app = UiApp.createApplication(),
      cPanel = app.createCaptionPanel(title).setId('cPanel'),
      vPanel = app.createVerticalPanel(),
      checkPanel = app.createHorizontalPanel(),
      handler = app.createServerHandler("buttonAction"),
      buttonGrid = app.createGrid(1, 3),
      ok = app.createButton(OK, handler).setId(OK),
      cancel = app.createButton(CANCEL, handler).setId(CANCEL),
      check = app.createCheckBox("Do not show me this diolog again.").setName('check'),
      label = app.createLabel(PROCESSING).setVisible(false);

    vPanel.add(app.createHTML(html));
    buttonGrid.setWidget(0, 0, ok);
    buttonGrid.setWidget(0, 1, cancel);
    buttonGrid.setWidget(0, 2, label);
    vPanel.add(buttonGrid);
    checkPanel.add(check);
    vPanel.add(checkPanel);
    cPanel.add(vPanel);
    app.add(cPanel);

    app.setHeight(height);

    var statusTextBox = app.createTextBox(),
      confirmationTextBox = app.createTextBox(),
      selectionBox = app.createTextBox();
    statusTextBox.setName(STATUS);
    statusTextBox.setVisible(false);
    statusTextBox.setValue(status);
    confirmationTextBox.setName(CONFIRMATION);
    confirmationTextBox.setVisible(false);
    confirmationTextBox.setValue(confirmation);
    selectionBox.setName(SELECTION);
    selectionBox.setVisible(false);
    selectionBox.setValue(selection);
    vPanel.add(statusTextBox);
    vPanel.add(confirmationTextBox);
    vPanel.add(selectionBox);
    handler.addCallbackElement(vPanel);

    var clientHandler = app.createClientHandler()
      .forTargets([ok, cancel]).setEnabled(false)
      .forTargets(label).setVisible(true);

    ok.addClickHandler(clientHandler);
    cancel.addClickHandler(clientHandler);

    ss.show(app);
  }
}

function manual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
    app = UiApp.createApplication().setTitle('Set state of selection');

  var grid = app.createGrid(3, 2);
  grid.setWidget(0, 0, app.createLabel(STATUS + ':'));
  var statusListBox = app.createListBox(false).setName(STATUS);
  statusListBox.addItem(DO_NOT_CHANGE);
  statusListBox.addItem(NO_MEDIA);
  statusListBox.addItem(MEDIA_PRESENT);
  statusListBox.addItem(PROBLEM);
  grid.setWidget(0, 1, statusListBox);
  grid.setWidget(1, 0, app.createLabel(CONFIRMATION + ':'));
  var confirmationListBox = app.createListBox(false).setName(CONFIRMATION);
  confirmationListBox.addItem(DO_NOT_CHANGE);
  confirmationListBox.addItem(NOT_CONFIRMED);
  confirmationListBox.addItem(CONFIRMED);
  grid.setWidget(1, 1, confirmationListBox);
  grid.setWidget(2, 0, app.createLabel(SELECTION + ':'));
  var selectionListBox = app.createListBox(false).setName(SELECTION);
  selectionListBox.addItem(DO_NOT_CHANGE);
  selectionListBox.addItem(NOT_SELECTED);
  selectionListBox.addItem(SELECTED);
  grid.setWidget(2, 1, selectionListBox);

  var cPanel = app.createCaptionPanel("Set State Of Range").setId('cPanel'),
    vPanel = app.createVerticalPanel();
  vPanel.add(app.createHTML('<p>The status of a submission is given by the value of three fields which are <b>' + STATUS + '</b>, <b>' + CONFIRMATION + '</b> and <b>' + SELECTION + '</b>.</p><p>For full details on what the various setting mean, please refer to documentation.</p>'));
  vPanel.add(grid);

  var handler = app.createServerHandler("buttonAction"),
    buttonGrid = app.createGrid(1, 3),
    ok = app.createButton(OK, handler).setId(OK),
    cancel = app.createButton(CANCEL, handler).setId(CANCEL),
    label = app.createLabel(PROCESSING).setVisible(false);
  handler.addCallbackElement(vPanel);
  buttonGrid.setWidget(0, 0, ok);
  buttonGrid.setWidget(0, 1, cancel);
  buttonGrid.setWidget(0, 2, label);

  var clientHandler = app.createClientHandler()
    .forTargets([ok, cancel]).setEnabled(false)
    .forTargets(label).setVisible(true);
  ok.addClickHandler(clientHandler);
  cancel.addClickHandler(clientHandler);

  vPanel.add(buttonGrid);
  handler.addCallbackElement(vPanel);

  cPanel.add(vPanel);
  app.add(cPanel);

  ss.show(app);
}

function buttonAction(e) {
  var app = UiApp.getActiveApplication(),
    ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptProperties.setProperty(normalizeHeader(e.parameter[STATUS] + ' ' + e.parameter[CONFIRMATION]), e.parameter.check);

  if(e.parameter.source === OK) {
    setStatus(e.parameter[STATUS], e.parameter[CONFIRMATION], e.parameter[SELECTION]);
  } else {
    ss.toast("Canceling operation.", "WARNING", 5);
  }
  return app.close();
}

function setStatus(status, confirmation, selection) {
  try {
    log('setStatus:(status, confirmation, selection):(' + status + ',' + confirmation + ',' + selection + ')');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if(ss.getSheetName() === FILM_SUBMISSIONS_SHEET) {
      var activeRange = SpreadsheetApp.getActiveRange(),
        filmSheet = ss.getSheetByName(FILM_SUBMISSIONS_SHEET),
        dataRange = filmSheet.getDataRange(),
        statusRange = findStatusRange(filmSheet, activeRange, dataRange);
      if(statusRange) {
        setStatusByStatusRange(ss, filmSheet, statusRange, status, confirmation, selection);
        //log status change
        var headersRange = filmSheet.getRange(1, 1, 1, filmSheet.getDataRange().getLastColumn()),
          headers = normalizeHeaders(headersRange.getValues()[0]),
          idIndex = headers.indexOf(normalizeHeader(FILM_ID)) + 1,
          ids = filmSheet.getRange(statusRange.getRow(), idIndex, statusRange.getHeight(), 1).getValues();
        log('The following film submissions have been set to "' + status + '", "' + confirmation + '", "' + selection + '": ' + ids.concat());
      } else {
        ss.toast("No films selected!", "WARNING", 5);
      }
    } else {
      ss.toast("Can only set Sate on '" + FILM_SUBMISSIONS_SHEET + "' sheet.", "WARNING", 5);
    }
  } catch(e) {
    log('There has been an error in setStatus:' + e);
  }
}

function setStatusByStatusRange(ss, filmSheet, statusRange, status, confirmation, selection) {
  var statusHeader = normalizeHeader(STATUS),
    confirmationHeader = normalizeHeader(CONFIRMATION),
    selectionHeader = normalizeHeader(SELECTION),
    rows = getRowsData(filmSheet, statusRange, 1);

  for(var i = 0; i < rows.length; i++) {
    if(status !== DO_NOT_CHANGE) {
      rows[i][statusHeader] = status;
    } else {
      status = rows[i][statusHeader];
    }
    if(confirmation !== DO_NOT_CHANGE) {
      rows[i][confirmationHeader] = confirmation;
    } else {
      confirmation = rows[i][confirmationHeader];
    }
    if(selection !== DO_NOT_CHANGE) {
      rows[i][selectionHeader] = selection;
    } else {
      selection = rows[i][selectionHeader];
    }
  }
  setRowsData(filmSheet, rows, filmSheet.getRange(1, statusRange.getColumn(), 1, statusRange.getWidth()), statusRange.getRow());

  var color = findStatusColor(ss, status, confirmation, selection);
  filmSheet.getRange(statusRange.getRow(), 1, statusRange.getHeight(), filmSheet.getDataRange().getLastColumn()).setBackground(color);
}

function findStatusColor(ss, status, confirmation, selection) {
  var color = NOT_SELECTED;

  if(selection === SELECTED) {
    color = SELECTED;
  } else if(status === PROBLEM) {
    color = PROBLEM;
  } else {
    color = status + ' ' + confirmation;
  }
  return getNamedValue(ss, color);
}

function findStatusRange(filmSheet, activeRange, dataRange) {
  var firstRowA = activeRange.getRow(),
    lastRowA = activeRange.getLastRow(),
    firstRowB = dataRange.getRow(),
    lastRowB = dataRange.getLastRow(),
    range = null,
    low = Math.max(firstRowA, firstRowB, 2), //cannot include sheet header
    high = Math.min(lastRowA, lastRowB);

  if(low <= high) {
    var minMaxColumn = findMinMaxColumns(filmSheet, [STATUS, CONFIRMATION, SELECTION]);
    range = filmSheet.getRange(low, minMaxColumn.min, high - low + 1, minMaxColumn.max - minMaxColumn.min + 1);
  }
  return range;
}

function pleaseWait(ss) {
  var app = UiApp.createApplication(),
    label = app.createLabel(PROCESSING).setStyleAttribute("fontSize", "200%");
  app.add(label);
  ss.show(app);
}

function editAndSaveTemplates(template) {
  try {
    var returnTo = template;
    template = template || SUBMISSION_CONFIRMATION;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    pleaseWait(ss);

    loadData(ss, TEST_DATA);
    loadData(ss, TEMPLATE_DATA);

    var app = UiApp.createApplication(),

      // Use these to turn off save, previous and next buttons if we have come from AD_HOC_MAIL or SELECTION_NOTIFICATION
      // As I turn them all off or all on together, logically I only need one hidden field for this information
      previousHidden = app.createHidden(normalizeHeader(PREVIOUS + ' ' + HIDDEN), !returnTo),
      nextHidden = app.createHidden(normalizeHeader(NEXT + ' ' + HIDDEN), !returnTo),
      saveHidden = app.createHidden(normalizeHeader(SAVE + ' ' + HIDDEN), !returnTo),

      vRoot = app.createVerticalPanel().setId(normalizeHeader('vRoot')),
      hTop = app.createHorizontalPanel(),
      handler = app.createServerHandler("templatesButtonAction"),
      saveButton = app.createButton(SAVE, handler).setId(normalizeHeader(SAVE)).setEnabled(!returnTo),
      cancelButton = app.createButton(CANCEL, handler).setId(normalizeHeader(CANCEL)),

      previousButtons = {},
      nextButtons = {},
      testButtons = {},

      previousButtonsArray = [],
      nextButtonsArray = [],
      testButtonsArray = [],

      processingLabel = app.createLabel(PROCESSING).setId(normalizeHeader(WAIT)).setVisible(false),

      returnToNotificationLabel = app.createLabel(SAVE_AND_RETURN_TO),
      returnToNotificationButton = app.createButton(SELECTION_NOTIFICATION, handler).setId(normalizeHeader(SELECTION_NOTIFICATION + ' ' + BUTTON)),
      returnToNotificationPLabel = app.createLabel(PROCESSING).setId(normalizeHeader(SELECTION_NOTIFICATION + ' ' + PLABEL)).setVisible(false),

      returnToAdHocLabel = app.createLabel(SAVE_AND_RETURN_TO),
      returnToAdHocButton = app.createButton(AD_HOC_EMAIL, handler).setId(normalizeHeader(AD_HOC_EMAIL + ' ' + BUTTON)),
      returnToAdHocPLabel = app.createLabel(PROCESSING).setId(normalizeHeader(AD_HOC_EMAIL + ' ' + PLABEL)).setVisible(false),

      templates = [SUBMISSION_CONFIRMATION, RECEIPT_CONFIMATION, REMINDER, NOT_ACCEPTED, ACCEPTED, AD_HOC_EMAIL],
      remainingDailyQuota = MailApp.getRemainingDailyQuota(),
      remainingEmailQuota = remainingDailyQuota - MIN_QUOTA,
      remainingEmailQuotaHidden = app.createHidden(normalizeHeader(REMAINING_EMAIL_QUOTA), remainingEmailQuota).setId(normalizeHeader(REMAINING_EMAIL_QUOTA)),

      testButtonGrid = app.createGrid(1, 3),
      //testButton          = app.createButton(TEST, handler).setId(normalizeHeader(TEST)).setEnabled(remainingEmailQuota>0),
      testButtonLabel = app.createLabel('Send test email:').setStyleAttribute('font-weight', 'bold'),
      testProcessingLabel = app.createLabel(PROCESSING).setStyleAttribute("fontSize", "50%").setVisible(false).setId(normalizeHeader(TEST_PROCESSING_LABEL)),

      testData = [{
        namedValue: FIRST_NAME,
        type: 'createTextBox',
        width: '200',
        validate: {
          validateLength: {
            min: 1,
            max: null
          }
        },
        label: 'First name required'
      }, {
        namedValue: LAST_NAME,
        type: 'createTextBox',
        width: '200',
        validate: {
          validateLength: {
            min: 1,
            max: null
          }
        },
        label: 'Last name required'
      }, {
        namedValue: TITLE,
        type: 'createTextBox',
        width: '200',
        validate: {
          validateLength: {
            min: 1,
            max: null
          }
        },
        label: 'Title required'
      }, {
        namedValue: EMAIL,
        type: 'createTextBox',
        width: '200',
        validate: {
          validateEmail: {}
        },
        label: 'Valid email required'
      }],

      height = '400',
      appHeight = returnTo ? '530' : '500',

      vAvaliableTages = app.createVerticalPanel(),
      cAvaliableTages = app.createCaptionPanel(AVAILABLE_TAGS).setWidth('200').setHeight(height),
      testGrid = app.createGrid(2 * testData.length, 2),
      vTest = app.createVerticalPanel(),
      cTest = app.createCaptionPanel(SEND_TEST_EMAIL).setWidth('300').setHeight(height),

      statusHTML = app.createHTML().setHTML('<p>Remaining email quota for today:<b>' + remainingDailyQuota + '</b>. System will pause sending emails for the day when remaining email quota is less than or equal to: <b>' + MIN_QUOTA + '</b>.</p>').setId(normalizeHeader(STATUS_HTML)),

      statusWarningHTML = app.createHTML().setHTML('<p>NOTE: Todays remaining email quota is now less than or equal to <b>' + MIN_QUOTA + '</b>. Hence the system has now paused sending emails for the day.</p>').setStyleAttribute("color", "red").setVisible(remainingEmailQuota <= 0).setId(normalizeHeader(STATUS_WARNING_HTML)),

      avalibleTagsHTML = app.createHTML('<p>Name of film festival<br/><b>${"Festival Name"}</b></p><p>Address of festival website<br/><b>${"Festival Website"}</b></p><p>Link to release form<br/><b>${"Release Link"}</b></p><p>Date of close of submission<br/><b>${"Close Of Submission"}</b></p><p>Date of start of film festival<br/><b>${"Event Date"}</b></p><p>First name of filmmaker<br/><b>${"First Name"}</b></p><p>Last name of filmmaker<br/><b>${"Last Name"}</b></p><p>Title of submission<br/><b>${"Title"}</b></p>');

    // Every template needs its own previous, next and test button
    var templateData = {}, templateName, previousTemplate, nextTemplate;
    for(var i = 0; i < templates.length; i++) {
      templateName = templates[i];
      previousButtons[templateName] = app.createButton(PREVIOUS).setVisible(templateName === template).setId(normalizeHeader(templateName + ' ' + PREVIOUS)).setEnabled(!returnTo);
      previousButtonsArray.push(previousButtons[templateName]);
      nextButtons[templateName] = app.createButton(NEXT).setVisible(templateName === template).setId(normalizeHeader(templateName + ' ' + NEXT)).setEnabled(!returnTo);
      nextButtonsArray.push(nextButtons[templateName]);
      testButtons[templateName] = app.createButton(TEST, handler).setVisible(templateName === template).setId(normalizeHeader(templateName + ' ' + TEST)).setEnabled(remainingEmailQuota > 0),
      testButtonsArray.push(testButtons[templateName]);
    }
    for(i = 0; i < templates.length; i++) {
      templateName = templates[i];

      var cTemplate = app.createCaptionPanel(templateName).setId(normalizeHeader(templateName)),
        subjectLabel = app.createLabel(SUBJECT + ':').setStyleAttribute('font-weight', 'bold'),
        subjectHelpLabel = app.createLabel(SUBJECT + ' required').setStyleAttribute("fontSize", "50%").setId(normalizeHeader(SUBJECT + ' ' + HELP)),
        bodyLabel = app.createLabel(BODY + ':').setStyleAttribute('font-weight', 'bold'),
        bodyHelpLabel = app.createLabel(BODY + ' required').setStyleAttribute("fontSize", "50%").setId(normalizeHeader(BODY + ' ' + HELP)),
        subjectTextBox = app.createTextBox().setName(normalizeHeader(SUBJECT_LINE)).setText(getNamedValue(ss, templateName + ' ' + SUBJECT_LINE)).setWidth('300').setName(normalizeHeader(templateName + ' ' + SUBJECT_LINE)),
        bodyTextArea = app.createTextArea().setName(normalizeHeader(BODY)).setText(getNamedValue(ss, templateName + ' ' + BODY)).setWidth('300').setHeight('280').setName(normalizeHeader(templateName + ' ' + BODY)),
        grid = app.createGrid(6, 1);

      grid.setWidget(0, 0, subjectLabel);
      grid.setWidget(1, 0, subjectTextBox);
      grid.setWidget(2, 0, subjectHelpLabel);
      grid.setWidget(3, 0, bodyLabel);
      grid.setWidget(4, 0, bodyTextArea);
      grid.setWidget(5, 0, bodyHelpLabel);
      cTemplate.add(grid).setHeight(height);

      var items = [{
        field: subjectTextBox,
        help: subjectHelpLabel
      }, {
        field: bodyTextArea,
        help: bodyHelpLabel
      }],
        item, right, wrong;
      for(var j = 0; j < items.length; j++) {
        item = items[j];

        wrong = app.createClientHandler()
          .validateNotLength(item.field, 1, null)
          .forTargets(item.help).setStyleAttribute("color", "red")
          .forTargets([saveButton].concat(previousButtonsArray, nextButtonsArray, testButtonsArray)).setEnabled(false);
        item.field.addKeyUpHandler(wrong);

        right = app.createClientHandler()
          .validateLength(item.field, 1, null)
          .forTargets(item.help).setStyleAttribute("color", "black");
        item.field.addKeyUpHandler(right);
      }

      templateData[templateName] = {
        panel: cTemplate,
        subject: subjectTextBox,
        body: bodyTextArea
      };

      hTop.add(templateData[templateName].panel.setVisible(templates[i] === template));
    }

    function isNotPreviousTemplate(x) {
      return x !== previousTemplate;
    }

    function getPanel(x) {
      return templateData[x].panel;
    }

    function getPreviousButton(x) {
      return previousButtons[x];
    }

    function getNextButton(x) {
      return nextButtons[x];
    }

    function getTestButton(x) {
      return testButtons[x];
    }
    for(i = 0; i < templates.length; i++) {
      templateName = templates[i];
      previousTemplate = templates[((+i) - 1 + templates.length) % templates.length];
      nextTemplate = templates[((+i) + 1) % templates.length];
      previousButtons[templateName].addClickHandler(
        app.createClientHandler()
        // set all panels invisible except the previous panel
        .forTargets(templates.filter(isNotPreviousTemplate).map(getPanel)).setVisible(false)
        .forTargets(templateData[previousTemplate].panel).setVisible(true)

        // the same as above but for buttons
        .forTargets(templates.filter(isNotPreviousTemplate).map(getPreviousButton)).setVisible(false)
        .forTargets(templates.filter(isNotPreviousTemplate).map(getNextButton)).setVisible(false)
        .forTargets(templates.filter(isNotPreviousTemplate).map(getTestButton)).setVisible(false)
        .forTargets(previousButtons[previousTemplate]).setVisible(true)
        .forTargets(nextButtons[previousTemplate]).setVisible(true)
        .forTargets(testButtons[previousTemplate]).setVisible(true)
      );

      nextButtons[templateName].addClickHandler(
        app.createClientHandler()
        // set all panels invisible except the previous panel
        .forTargets(templates.filter(isNotPreviousTemplate).map(getPanel)).setVisible(false)
        .forTargets(templateData[nextTemplate].panel).setVisible(true)

        // the same as above but for buttons
        .forTargets(templates.filter(isNotPreviousTemplate).map(getPreviousButton)).setVisible(false)
        .forTargets(templates.filter(isNotPreviousTemplate).map(getNextButton)).setVisible(false)
        .forTargets(templates.filter(isNotPreviousTemplate).map(getTestButton)).setVisible(false)
        .forTargets(previousButtons[nextTemplate]).setVisible(true)
        .forTargets(nextButtons[nextTemplate]).setVisible(true)
        .forTargets(testButtons[nextTemplate]).setVisible(true)
      );
    }

    vAvaliableTages.add(avalibleTagsHTML);
    cAvaliableTages.add(vAvaliableTages);

    hTop.add(cAvaliableTages);

    //send test GUI
    var name, testItem = {}, testLabel = {}, itemData, testItemName = {};
    for(i = 0; i < testData.length; i++) {
      itemData = testData[i];
      testGrid.setWidget(2 * (+i), 0, app.createLabel(itemData.namedValue + ':').setStyleAttribute('font-weight', 'bold'));
      name = normalizeHeader(TEST + ' ' + itemData.namedValue);
      testItemName[i] = name;
      testItem[name] = app[itemData.type]().setName(name).setValue(getNamedValue(ss, normalizeHeader(name))).setWidth(itemData.width);
      testGrid.setWidget(2 * (+i), 1, testItem[name]);
      testLabel[name] = app.createLabel(itemData.label).setStyleAttribute("fontSize", "50%").setId(normalizeHeader(LABEL + ' ' + itemData.namedValue));
      testGrid.setWidget(2 * (+i) + 1, 1, testLabel[name]);
    }

    //client side validation
    var allGoodTest = app.createClientHandler()
      .validateLength(testItem[normalizeHeader(TEST_FIRST_NAME)], 1, null)
      .validateLength(testItem[normalizeHeader(TEST_LAST_NAME)], 1, null)
      .validateLength(testItem[normalizeHeader(TEST_TITLE)], 1, null)
      .validateEmail(testItem[normalizeHeader(TEST_EMAIL)])
      .validateRange(remainingEmailQuotaHidden, 1, null)

    .validateLength(templateData[SUBMISSION_CONFIRMATION].subject, 1, null)
      .validateLength(templateData[SUBMISSION_CONFIRMATION].body, 1, null)

    .validateLength(templateData[RECEIPT_CONFIMATION].subject, 1, null)
      .validateLength(templateData[RECEIPT_CONFIMATION].body, 1, null)

    .validateLength(templateData[REMINDER].subject, 1, null)
      .validateLength(templateData[REMINDER].body, 1, null)

    .validateLength(templateData[NOT_ACCEPTED].subject, 1, null)
      .validateLength(templateData[NOT_ACCEPTED].body, 1, null)

    .validateLength(templateData[ACCEPTED].subject, 1, null)

    .validateLength(templateData[ACCEPTED].body, 1, null)

    .validateLength(templateData[AD_HOC_EMAIL].subject, 1, null)
      .validateLength(templateData[AD_HOC_EMAIL].body, 1, null)

    .forTargets([returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray)).setEnabled(true),

      allGoodSavePreviousNext = app.createClientHandler()
        .validateLength(testItem[normalizeHeader(TEST_FIRST_NAME)], 1, null)
        .validateLength(testItem[normalizeHeader(TEST_LAST_NAME)], 1, null)
        .validateLength(testItem[normalizeHeader(TEST_TITLE)], 1, null)
        .validateEmail(testItem[normalizeHeader(TEST_EMAIL)])
        .validateRange(remainingEmailQuotaHidden, 1, null)

      .validateLength(templateData[SUBMISSION_CONFIRMATION].subject, 1, null)
        .validateLength(templateData[SUBMISSION_CONFIRMATION].body, 1, null)

      .validateLength(templateData[RECEIPT_CONFIMATION].subject, 1, null)
        .validateLength(templateData[RECEIPT_CONFIMATION].body, 1, null)

      .validateLength(templateData[REMINDER].subject, 1, null)
        .validateLength(templateData[REMINDER].body, 1, null)

      .validateLength(templateData[NOT_ACCEPTED].subject, 1, null)
        .validateLength(templateData[NOT_ACCEPTED].body, 1, null)

      .validateLength(templateData[ACCEPTED].subject, 1, null)
        .validateLength(templateData[ACCEPTED].body, 1, null)

      .validateLength(templateData[AD_HOC_EMAIL].subject, 1, null)
        .validateLength(templateData[AD_HOC_EMAIL].body, 1, null)

      .validateOptions(previousHidden, ['true'])
        .validateOptions(nextHidden, ['true'])
        .validateOptions(saveHidden, ['true'])

      .forTargets([saveButton].concat(previousButtonsArray, nextButtonsArray)).setEnabled(true);

    var testItemRight, testItemWrong;
    for(i = 0; i < testData.length; i++) {
      itemData = testData[i];
      if(itemData.validate) {
        name = testItemName[i];
        testItemRight = null;
        testItemWrong = null;
        if(itemData.validate.validateLength) {
          var min = itemData.validate.validateLength.min,
            max = itemData.validate.validateLength.max;
          testItemRight = app.createClientHandler().validateLength(testItem[name], min, max)
            .forTargets(testLabel[name]).setStyleAttribute("color", "black"),
          testItemWrong = app.createClientHandler().validateNotLength(testItem[name], min, max)
            .forTargets(testLabel[name]).setStyleAttribute("color", "red")
            .forTargets([saveButton, returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray, previousButtonsArray, nextButtonsArray)).setEnabled(false);
        } else if(itemData.validate.validateEmail) {
          testItemRight = app.createClientHandler().validateEmail(testItem[name])
            .forTargets(testLabel[name]).setStyleAttribute("color", "black"),
          testItemWrong = app.createClientHandler().validateNotEmail(testItem[name])
            .forTargets(testLabel[name]).setStyleAttribute("color", "red")
            .forTargets([saveButton, returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray, previousButtonsArray, nextButtonsArray)).setEnabled(false);
        }
        if(testItemRight) {
          testItem[name].addKeyUpHandler(testItemRight);
        }
        if(testItemWrong) {
          testItem[name].addKeyUpHandler(testItemWrong);
        }
        testItem[name].addKeyUpHandler(allGoodTest);
        testItem[name].addKeyUpHandler(allGoodSavePreviousNext);
      }
    }

    for(i = 0; i < templates.length; i++) {
      templateName = templates[i];
      templateData[templateName].subject.addKeyUpHandler(allGoodTest);
      templateData[templateName].subject.addKeyUpHandler(allGoodSavePreviousNext);
      templateData[templateName].body.addKeyUpHandler(allGoodTest);
      templateData[templateName].body.addKeyUpHandler(allGoodSavePreviousNext);
    }

    // end of client side validation

    // Add test buttons to grid
    var off = app.createClientHandler()
      .forTargets([saveButton, returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray, previousButtonsArray, nextButtonsArray)).setEnabled(false)
      .forTargets(testProcessingLabel).setVisible(true);
    vTest.add(testGrid);
    vTest.add(statusHTML);
    testButtonGrid.setWidget(0, 0, testButtonLabel);
    var hTestButtonsPanel = app.createHorizontalPanel();
    for(i = 0; i < testButtonsArray.length; i++) {
      testButtonsArray[i].addClickHandler(off);
      hTestButtonsPanel.add(testButtonsArray[i]);
    }
    testButtonGrid.setWidget(0, 1, hTestButtonsPanel);
    testButtonGrid.setWidget(0, 2, testProcessingLabel);
    vTest.add(testButtonGrid);
    vTest.add(statusWarningHTML);

    cTest.add(vTest);
    // end of send test GUI

    hTop.add(cTest);

    vRoot.add(hTop);

    // Build control panel
    var cControls = app.createCaptionPanel("Controls").setId(normalizeHeader('cControls')),
      buttonGrid = app.createGrid(1, 5),
      returnGrid = app.createGrid(1, 3),
      vControls = app.createVerticalPanel(),
      clickSCPN = app.createClientHandler()
        .forTargets([saveButton, cancelButton, returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray, previousButtonsArray, nextButtonsArray)).setEnabled(false)
        .forTargets(processingLabel).setVisible(true);
    saveButton.addClickHandler(clickSCPN);
    cancelButton.addClickHandler(clickSCPN);

    var clickAdHocNot = app.createClientHandler()
      .forTargets([saveButton, cancelButton, returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray, previousButtonsArray, nextButtonsArray)).setEnabled(false)
      .forTargets(returnToNotificationPLabel).setVisible(true)
      .forTargets(returnToAdHocPLabel).setVisible(true);
    returnToNotificationButton.addClickHandler(clickAdHocNot);
    returnToAdHocButton.addClickHandler(clickAdHocNot);

    buttonGrid.setWidget(0, 0, saveButton);
    buttonGrid.setWidget(0, 1, cancelButton);
    var hPreviousButtons = app.createHorizontalPanel();
    for(i = 0; i < previousButtonsArray.length; i++) {
      hPreviousButtons.add(previousButtonsArray[i]);
    }
    buttonGrid.setWidget(0, 2, hPreviousButtons);
    var hNextButtons = app.createHorizontalPanel();
    for(i = 0; i < nextButtonsArray.length; i++) {
      hNextButtons.add(nextButtonsArray[i]);
    }
    buttonGrid.setWidget(0, 3, hNextButtons);
    buttonGrid.setWidget(0, 4, processingLabel);
    vControls.add(buttonGrid);

    if(returnTo) {
      if(returnTo === ACCEPTED || returnTo === NOT_ACCEPTED) {
        returnGrid.setWidget(0, 0, returnToNotificationLabel);
        returnGrid.setWidget(0, 1, returnToNotificationButton);
        returnGrid.setWidget(0, 2, returnToNotificationPLabel);
      } else {
        returnGrid.setWidget(0, 0, returnToAdHocLabel);
        returnGrid.setWidget(0, 1, returnToAdHocButton);
        returnGrid.setWidget(0, 2, returnToAdHocPLabel);
      }
      vControls.add(returnGrid);
    }

    cControls.add(vControls);

    vRoot.add(cControls);

    // hidden fields
    vRoot.add(remainingEmailQuotaHidden);
    vRoot.add(previousHidden);
    vRoot.add(nextHidden);
    vRoot.add(saveHidden);
    var tags = [TEST_TITLE, TEST_FIRST_NAME, TEST_LAST_NAME, TEST_EMAIL];
    vRoot.add(app.createHidden(normalizeHeader(TAGS), tags.join(',')));
    vRoot.add(app.createHidden(normalizeHeader(TEMPLATE_DATA_NAMES), templates.join(',')));
    vRoot.add(app.createHidden(normalizeHeader(TEST_DATA_NAMES), testData.map(function (item) {
      return TEST + ' ' + item.namedValue;
    }).join(',')));

    handler.addCallbackElement(vRoot);

    app.add(vRoot);

    app.setWidth('900');
    app.setHeight(appHeight);

    ss.show(app);

  } catch(e) {
    log('updateTemplates:' + catchToString(e));
  }
}



function templatesButtonAction(e) {
  function saveTestAndTemplateData() {
    var range = ss.getRangeByName(normalizeHeader(TEST_DATA)),
      values = testDataNames.map(function (itemName) {
        return [itemName, e.parameter[normalizeHeader(itemName)]];
      });
    range.setValues(values);

    var templatesAndSuffix = templates.map(function (item) {
      return [item + ' ' + SUBJECT_LINE, item + ' ' + BODY];
    });
    templatesAndSuffix = flatten(templatesAndSuffix);
    values = [
      [TEMPLATE_NAME, TEMPLATE]
    ].concat(templatesAndSuffix.map(function (itemName) {
      return [itemName, e.parameter[normalizeHeader(itemName)]];
    })); // Add column titles in.
    range = ss.getRangeByName(normalizeHeader(TEMPLATE_DATA));
    range.setValues(values);

    ss.toast("Data has been saved.", "INFORMATION", 5);
  }
  try {
    var app = UiApp.getActiveApplication(),
      ss = SpreadsheetApp.getActiveSpreadsheet(),
      testDataNames   = e.parameter[normalizeHeader(TEST_DATA_NAMES)].split(','),
      templates = e.parameter[normalizeHeader(TEMPLATE_DATA_NAMES)].split(','),
      testButtonNames = templates.map(function (x) {
        return normalizeHeader(x + ' ' + TEST);
      });

    if(e.parameter.source === normalizeHeader(SELECTION_NOTIFICATION + ' ' + BUTTON)) {
      saveTestAndTemplateData();
      selectionNotification();
    } else if(e.parameter.source === normalizeHeader(AD_HOC_EMAIL + ' ' + BUTTON)) {
      saveTestAndTemplateData();
      adHocEmail();
    } else if(testButtonNames.indexOf(e.parameter.source) > -1) {
      var template = templates[testButtonNames.indexOf(e.parameter.source)],
        remainingEmailQuotaTextBox = app.getElementById(normalizeHeader(REMAINING_EMAIL_QUOTA)),
        remainingDailyQuota = MailApp.getRemainingDailyQuota();

      var tags = e.parameter[normalizeHeader(TAGS)].split(','),
        mergeData = {}, tag;

      for(var i = 0; i < tags.length; i++) {
        tag = tags[i];
        mergeData[normalizeHeader(tag.replace(/^Test /, ''))] = e.parameter[normalizeHeader(tag)];
      }

      var statusHTML = app.getElementById(normalizeHeader(STATUS_HTML)),
        statusWarningHTML = app.getElementById(normalizeHeader(STATUS_WARNING_HTML)),
        testProcessingLabel = app.getElementById(normalizeHeader(TEST_PROCESSING_LABEL)),
        saveButton = app.getElementById(normalizeHeader(SAVE)),
        cancelButton = app.getElementById(normalizeHeader(CANCEL)),
        returnToNotificationButton = app.getElementById(normalizeHeader(SELECTION_NOTIFICATION + ' ' + BUTTON)),
        returnToAdHocButton = app.getElementById(normalizeHeader(AD_HOC_EMAIL + ' ' + BUTTON));

      mergeData[normalizeHeader(EMAIL)] = e.parameter[normalizeHeader(TEST_EMAIL)];
      loadData(ss, FESTIVAL_DATA);

      if(remainingDailyQuota > MIN_QUOTA) { // check just in case
        remainingDailyQuota = mergeAndSend(ss, mergeData, template, remainingDailyQuota);
        ss.toast("Test email has been sent.", "INFORMATION", 5);
      } else {
        ss.toast("Not enough email quota remained, email not sent!", "WARNING!", 5);
      }

      remainingEmailQuotaTextBox.setValue((remainingDailyQuota - MIN_QUOTA).toString());
      statusHTML.setHTML('<p>Remaining email quota for today:<b>' + remainingDailyQuota + '</b>. System will pause sending emails for the day when remaining email quota is less than or equal to: <b>' + MIN_QUOTA + '</b>.</p>');
      if(remainingDailyQuota <= MIN_QUOTA) {
        statusWarningHTML.setVisible(true);
      }
      saveButton.setEnabled(e.parameter[normalizeHeader(SAVE + ' ' + HIDDEN)] === 'true');
      cancelButton.setEnabled(true);
      for(i = 0; i < templates.length; i++) {
        app.getElementById(normalizeHeader(templates[i] + ' ' + PREVIOUS)).setEnabled(e.parameter[normalizeHeader(PREVIOUS + ' ' + HIDDEN)] === 'true');
        app.getElementById(normalizeHeader(templates[i] + ' ' + NEXT)).setEnabled(e.parameter[normalizeHeader(NEXT + ' ' + HIDDEN)] === 'true');
        app.getElementById(normalizeHeader(templates[i] + ' ' + TEST)).setEnabled(remainingDailyQuota > MIN_QUOTA);
      }
      returnToNotificationButton.setEnabled(true);
      returnToAdHocButton.setEnabled(true);
      testProcessingLabel.setVisible(false);
    } else if(e.parameter.source === normalizeHeader(SAVE)) {
      saveTestAndTemplateData();
      app.close();
    } else {
      ss.toast("Canceling operation.", "WARNING", 5);
      app.close();
    }
  } catch(error) {
    log('templatesButtonAction:error::' + catchToString(error));
  }
  return app;
}

function settingsOptions() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    pleaseWait(ss);

    var festivaDataNames = loadData(ss, FESTIVAL_DATA),
      app = UiApp.createApplication(),
      festivaDataNamesHidden = app.createHidden(normalizeHeader(FESTIVAL_DATA_NAMES), festivaDataNames.join(',')),
      gridData = [{
        namedValue: FESTIVAL_NAME,
        type: 'createTextBox',
        width: '300px',
        help: 'Required: please give the full name of the festival.',
        error: 'You must give the full festival name.',
        validate: {
          validateLength: {
            min: 1,
            max: null
          }
        }
      }, {
        namedValue: FESTIVAL_WEBSITE,
        type: 'createTextBox',
        width: '300px',
        help: 'Not required.'
      }, {
        namedValue: CLOSE_OF_SUBMISSION,
        type: 'createDateBox',
        width: '75px',
        help: 'Required (default ' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + ').',
        validate: {
          valadateDateDependance: {}
        }
      }, {
        namedValue: EVENT_DATE,
        type: 'createDateBox',
        width: '75px',
        help: 'Required (default ' + prettyPrintDate(getNamedValue(ss, EVENT_DATE)) + ').',
        validate: {
          valadateDateDependance: {}
        }
      }, {
        namedValue: RELEASE_LINK,
        type: 'createTextBox',
        width: '300px',
        help: 'Not required.'
      }, {
        namedValue: DAYS_BEFORE_REMINDER,
        type: 'createTextBox',
        width: '30px',
        help: 'Required: please give days before reminder email is sent.',
        error: 'Must be whole positive number. No spaces.',
        validate: {
          validateRange: {
            min: 1,
            max: null
          },
          validateInteger: {}
        }
      }, {
        namedValue: ENABLE_REMINDER,
        type: 'createListBox',
        width: '100px',
        help: 'Required (default ' + getNamedValue(ss, ENABLE_REMINDER) + ').',
        list: [NOT_ENABLED, ENABLED]
      }, {
        namedValue: ENABLE_CONFIRMATION,
        type: 'createListBox',
        width: '100px',
        help: 'Required (default ' + getNamedValue(ss, ENABLE_REMINDER) + ').',
        list: [NOT_ENABLED, ENABLED]
      }, {
        namedValue: FIRST_FILM_ID,
        type: 'createHidden'
      }],
      handler = app.createServerHandler("settingsOptionsButtonAction"),
      save = app.createButton(SAVE, handler).setId(normalizeHeader(SAVE)),
      cancel = app.createButton(CANCEL, handler).setId(normalizeHeader(CANCEL)),
      cPanel = app.createCaptionPanel("Settings and Options").setId('cPanel'),
      vPanel = app.createVerticalPanel(),
      buttonGrid = app.createGrid(1, 3);

    var grid = app.createGrid(2 * gridData.length, 2),
      item = {}, label = {}, back = {},
      itemName, itemValue, correct, wrong, i;

    // build GUI
    for(i = 0; i < gridData.length; i++) {
      itemName = normalizeHeader(gridData[i].namedValue);
      itemValue = getNamedValue(ss, gridData[i].namedValue);
      if(gridData[i].type === 'createHidden') {
        vPanel.add(app.createHidden(itemName, itemValue));
        continue;
      }
      grid.setWidget(2 * (+i), 0, app.createLabel(gridData[i].namedValue + ': '));
      if(gridData[i].type === 'createListBox') {
        item[itemName] = app.createListBox();
        for(var j = 0; j < gridData[i].list.length; j++) {
          item[itemName].addItem(gridData[i].list[j]);
        }
        item[itemName].setSelectedIndex(gridData[i].list.indexOf(itemValue));
        grid.setWidget(2 * (+i), 1, item[itemName]);
      } else {
        item[itemName] = app[gridData[i].type]().setValue(itemValue);
        grid.setWidget(2 * (+i), 1, item[itemName]);
      }
      back[itemName] = i;
      item[itemName].setWidth(gridData[i].width).setId(itemName).setName(itemName);

      label[itemName] = app.createLabel('').setStyleAttribute("fontSize", "50%").setId(itemName + HELP);
      grid.setWidget(2 * (+i) + 1, 1, label[itemName]);

      if(gridData[i].help) {
        label[itemName].setText(gridData[i].help);
      }
    }

    // add validators to GUI
    for(i in gridData) {
      if(gridData[i].validate) {
        itemName = normalizeHeader(gridData[i].namedValue);
        if(gridData[i].validate.valadateDateDependance) {
          item[itemName].addValueChangeHandler(app.createServerHandler("validateDates").addCallbackElement(vPanel));
        }
        if(gridData[i].validate.validateLength) {
          wrong = app.createClientHandler().validateNotLength(item[itemName], gridData[i].validate.validateLength.min, gridData[i].validate.validateLength.max)
            .forTargets(label[itemName]).setText(gridData[i].error).setStyleAttribute("color", "red")
            .forTargets(save).setEnabled(false);
          item[itemName].addKeyUpHandler(wrong);
          correct = app.createClientHandler().validateLength(item[itemName], gridData[i].validate.validateLength.min, gridData[i].validate.validateLength.max)
            .forTargets(label[itemName]).setText(gridData[i].help).setStyleAttribute("color", "black");
          item[itemName].addKeyUpHandler(correct);
        }
        if(gridData[i].validate.validateRange) {
          wrong = app.createClientHandler().validateNotRange(item[itemName], gridData[i].validate.validateRange.min, gridData[i].validate.validateRange.max)
            .forTargets(label[itemName]).setText(gridData[i].error).setStyleAttribute("color", "red")
            .forTargets(save).setEnabled(false);
          item[itemName].addKeyUpHandler(wrong);
          correct = app.createClientHandler().validateRange(item[itemName], gridData[i].validate.validateRange.min, gridData[i].validate.validateRange.max)
            .forTargets(label[itemName]).setText(gridData[i].help).setStyleAttribute("color", "black");
          item[itemName].addKeyUpHandler(correct);
        }
        if(gridData[i].validate.validateInteger) {
          wrong = app.createClientHandler().validateNotInteger(item[itemName])
            .forTargets(label[itemName]).setText(gridData[i].error).setStyleAttribute("color", "red")
            .forTargets(save).setEnabled(false);
          item[itemName].addKeyUpHandler(wrong);
          correct = app.createClientHandler().validateInteger(item[itemName])
            .forTargets(label[itemName]).setText(gridData[i].help).setStyleAttribute("color", "black");
          item[itemName].addKeyUpHandler(correct);
        }
      }
    }
    var dateDiff = app.createHidden().setValue(diffDays(getNamedValue(ss, EVENT_DATE), getNamedValue(ss, CLOSE_OF_SUBMISSION))).setId(normalizeHeader(DATE_DIFF));

    var allCorrect = app.createClientHandler()
      .validateLength(item[normalizeHeader(FESTIVAL_NAME)], 1, null)
      .validateRange(dateDiff, 1, null)
      .validateRange(item[normalizeHeader(DAYS_BEFORE_REMINDER)], 1, null)
      .forTargets(save).setEnabled(true);
    item[normalizeHeader(FESTIVAL_NAME)].addKeyUpHandler(allCorrect);
    item[normalizeHeader(DAYS_BEFORE_REMINDER)].addKeyUpHandler(allCorrect);
    //dateDiff.addValueChangeHandler(allCorrect); this does not work :(

    vPanel.add(app.createHTML('<p>Please enter the settings and options for your film festival below.</p>'));
    vPanel.add(grid);
    vPanel.add(dateDiff);

    label = app.createLabel(PROCESSING).setId(normalizeHeader(WAIT)).setVisible(false);
    handler.addCallbackElement(vPanel);
    buttonGrid.setWidget(0, 0, save);
    buttonGrid.setWidget(0, 1, cancel);
    buttonGrid.setWidget(0, 2, label);

    var clientHandler = app.createClientHandler()
      .forTargets([save, cancel]).setEnabled(false)
      .forTargets(label).setVisible(true);
    save.addClickHandler(clientHandler);
    cancel.addClickHandler(clientHandler);

    vPanel.add(buttonGrid);
    vPanel.add(festivaDataNamesHidden);

    handler.addCallbackElement(vPanel);

    cPanel.add(vPanel);
    app.add(cPanel);

    app.setHeight('560');

    ss.show(app);
  } catch(e) {
    log('settingsOptions:error:' + catchToString(e));
  }
}

function validateDates(e) {
  var app = UiApp.getActiveApplication(),
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    eventDate = e.parameter[normalizeHeader(EVENT_DATE)],
    cosDate = e.parameter[normalizeHeader(CLOSE_OF_SUBMISSION)],
    eventLabel = app.getElementById(normalizeHeader(EVENT_DATE + ' ' + HELP)),
    cosLabel = app.getElementById(normalizeHeader(CLOSE_OF_SUBMISSION + ' ' + HELP)),
    save = app.getElementById(normalizeHeader(SAVE)),
    dateDiff = app.getElementById(normalizeHeader(DATE_DIFF));

  if(eventDate instanceof Date && cosDate instanceof Date) {
    // case: both dates


    if(eventDate < cosDate) {
      // error condition
      cosLabel.setText('"' + CLOSE_OF_SUBMISSION + '" must be before "' + EVENT_DATE + '".').setStyleAttribute("color", "red");
      eventLabel.setText('"' + CLOSE_OF_SUBMISSION + '" must be before "' + EVENT_DATE + '".').setStyleAttribute("color", "red");
      save.setEnabled(false);
    } else {
      cosLabel.setText('Required (default ' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + ').').setStyleAttribute("color", "black");
      eventLabel.setText('Required (default ' + prettyPrintDate(getNamedValue(ss, EVENT_DATE)) + ').').setStyleAttribute("color", "black");
      var festivalName = e.parameter[normalizeHeader(FESTIVAL_NAME)],
        daysBeforeReminder = e.parameter[normalizeHeader(DAYS_BEFORE_REMINDER)];

      if(festivalName && festivalName.length > 0 && daysBeforeReminder && parseInt(daysBeforeReminder) > 0) {
        save.setEnabled(true); //only enable save if everything is OK
      }
    }
    dateDiff.setValue(diffDays(eventDate, cosDate));
  } else {
    //case: at least one is not a date
    if(eventDate instanceof Date) {
      eventLabel.setText('Required (default ' + prettyPrintDate(getNamedValue(ss, EVENT_DATE)) + ').').setStyleAttribute("color", "black");
    } else {
      eventLabel.setText('Not a valid date').setStyleAttribute("color", "red");
    }

    if(cosDate instanceof Date) {
      cosLabel.setText('Required (default ' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + ').').setStyleAttribute("color", "black");
    } else {
      cosLabel.setText('Not a valid date').setStyleAttribute("color", "red");
    }

    save.setEnabled(false);
    dateDiff.setValue(NaN);
  }

  return app;
}

function settingsOptionsButtonAction(e) {
  var app = UiApp.getActiveApplication(),
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    eventDate = e.parameter[normalizeHeader(EVENT_DATE)],
    cosDate = e.parameter[normalizeHeader(CLOSE_OF_SUBMISSION)],
    eventLabel = app.getElementById(normalizeHeader(EVENT_DATE + ' ' + HELP)),
    cosLabel = app.getElementById(normalizeHeader(CLOSE_OF_SUBMISSION + ' ' + HELP)),
    save = app.getElementById(normalizeHeader(SAVE)),
    cancel = app.getElementById(normalizeHeader(CANCEL)),
    wait = app.getElementById(normalizeHeader(WAIT));

  if(e.parameter.source === normalizeHeader(SAVE)) {
    // need to check that dates are not wrong before saving
    if(!(eventDate instanceof Date)) {
      eventLabel.setText('Not a valid date').setStyleAttribute("color", "red");
      save.setEnabled(false);
      cancel.setEnabled(true);
      wait.setVisible(false);
      return app; // do not close the dialog
    }
    if(!(cosDate instanceof Date)) {
      cosLabel.setText('Not a valid date').setStyleAttribute("color", "red");
      save.setEnabled(false);
      cancel.setEnabled(true);
      wait.setVisible(false);
      return app; // do not close the dialog
    }
    //if we got here, the dates should indeed be dates!
    if(eventDate < cosDate) {
      // error condition
      cosLabel.setText('"' + CLOSE_OF_SUBMISSION + '" must be before "' + EVENT_DATE + '".').setStyleAttribute("color", "red");
      eventLabel.setText('"' + CLOSE_OF_SUBMISSION + '" must be before "' + EVENT_DATE + '".').setStyleAttribute("color", "red");
      save.setEnabled(false);
      cancel.setEnabled(true);
      wait.setVisible(false);
      return app; // do not close the dialog
    }


    var festivaDataNames = e.parameter[normalizeHeader(FESTIVAL_DATA_NAMES)].split(','),
      range = ss.getRangeByName(normalizeHeader(FESTIVAL_DATA)),
      values = festivaDataNames.map(function (itemName) {
        return [itemName, e.parameter[normalizeHeader(itemName)]];
      });
    range.setValues(values);
  } else {
    ss.toast("Canceling operation.", "WARNING", 5);
  }

  return app.close();
}



function selectionNotificationAcceptedTemplate() {
  editAndSaveTemplates(ACCEPTED);
}

function selectionNotificationNotAcceptedTemplate() {
  editAndSaveTemplates(NOT_ACCEPTED);
}

function selectionNotification() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    pleaseWait(ss);

    var app = UiApp.createApplication(),
      cPanel = app.createCaptionPanel("Queue Selecttion Notification").setId('cPanel'),
      vPanel = app.createVerticalPanel(),
      buttonGrid = app.createGrid(1, 3),
      handler = app.createServerHandler("selectionNotificationButtonAction"),
      enable = app.createButton(ENABLE, handler).setId(ENABLE),
      unenable = app.createButton(UNENABLE, handler).setId(UNENABLE),
      cancel = app.createButton(CANCEL, handler).setId(CANCEL),
      waitHTML = app.createHTML(PROCESSING).setVisible(false),
      cannotHTML = app.createHTML('<p>NOTE: You have not reached <b>' + CLOSE_OF_SUBMISSION + '</b> on ' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + ' yet. Hence you cannot yet enable <b>' + SELECTION_NOTIFICATION + '</b>.</p><br/>').setStyleAttribute("color", "red").setVisible(false),
      height = '680',
      currentDate = new Date(),
      closeOfSubmission = getNamedValue(ss, CLOSE_OF_SUBMISSION);


    vPanel.add(app.createHTML('<p><b>' + SELECTION_NOTIFICATION + '</b> is the process of mailing each submitting filmmaker with the selection status of their submission.</p><p>It is intended that you only enable <b>' + SELECTION_NOTIFICATION + '</b> once per film festival, after close ofsubmission.</p><p>Before you enable <b>' + SELECTION_NOTIFICATION + '</b>, each film submission that you are selecting for your film festival, must have its <b>' + SELECTION + '</b> field set to <b>' + SELECTED + '</b>.</p><p>When you enable <b>' + SELECTION_NOTIFICATION + '</b>, the system will start to process the submissions after midnight of that day, in their submission order.</p><p>If you unenable <b>' + SELECTION_NOTIFICATION + '</b> before midnight, the notification emails will not be sent and the <b>' + SELECTION_NOTIFICATION + '</b> will have been canceled.</p><p>Each submission that the system processes (that does not have <b>' + STATUS + '</b> set to <b>' + PROBLEM + '</b>) will receive the <b>' + ACCEPTED + '</b> email if it has <b>' + SELECTION + '</b> set to <b>' + SELECTED + '</b>, otherwise it will receive the <b>' + NOT_ACCEPTED + '</b> email.</p><p>A submission with <b>' + STATUS + '</b> set to <b>' + PROBLEM + '</b> will not receive a <b>' + SELECTION_NOTIFICATION + '</b> email.</p>'));

    // link to edit and test ACCEPTED template
    var grid = app.createGrid(2, 3),
      acceptedHandle = app.createServerHandler('selectionNotificationAcceptedTemplate'),
      acceptedButton = app.createButton(ACCEPTED, acceptedHandle).setId(normalizeHeader(ACCEPTED)),
      acceptedPLabel = app.createLabel(PROCESSING).setId(normalizeHeader(ACCEPTED + ' ' + PLABEL)).setStyleAttribute("fontSize", "50%").setVisible(false);
    grid.setWidget(0, 0, app.createHTML('Edit and test <b>' + ACCEPTED + '</b> template:'));
    grid.setWidget(0, 1, acceptedButton);
    grid.setWidget(0, 2, acceptedPLabel);

    // link to edit and test NOT_ACCEPTED template
    var notAcceptedHandle = app.createServerHandler('selectionNotificationNotAcceptedTemplate'),
      notAcceptedButton = app.createButton(NOT_ACCEPTED, notAcceptedHandle).setId(normalizeHeader(NOT_ACCEPTED)),
      notAcceptedPLabel = app.createLabel(PROCESSING).setId(normalizeHeader(NOT_ACCEPTED + ' ' + PLABEL)).setStyleAttribute("fontSize", "50%").setVisible(false);
    grid.setWidget(1, 0, app.createHTML('Edit and test <b>' + NOT_ACCEPTED + '</b> template:'));
    grid.setWidget(1, 1, notAcceptedButton);
    grid.setWidget(1, 2, notAcceptedPLabel);

    vPanel.add(grid);

    var acceptedProcessing = app.createClientHandler()
      .forTargets([enable, unenable, acceptedButton, notAcceptedButton]).setEnabled(false)
      .forTargets(acceptedPLabel).setVisible(true);
    acceptedButton.addClickHandler(acceptedProcessing);

    var notAcceptedProcessing = app.createClientHandler()
      .forTargets([enable, unenable, acceptedButton, notAcceptedButton]).setEnabled(false)
      .forTargets(notAcceptedPLabel).setVisible(true);
    notAcceptedButton.addClickHandler(notAcceptedProcessing);

    vPanel.add(app.createHTML('<p>If the system gets within ' + MIN_QUOTA + ' of using up the daily email quota, the <b>' + SELECTION_NOTIFICATION + '</b> will be paused till the next day.</p><p>NOTE: your daily email quote is currently at ' + MailApp.getRemainingDailyQuota() + '.</p><p>You cannot enable <b>' + SELECTION_NOTIFICATION + '</b> before <b>' + CLOSE_OF_SUBMISSION + '</b> on ' + prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION)) + '.</p><p><p>To continue with enabling <b>' + SELECTION_NOTIFICATION + '</b> press <b>' + ENABLE + '</b>, to not do this press <b>' + CANCEL + '</b>.</p><br/>'));

    if(currentDate < closeOfSubmission) {
      cannotHTML.setVisible(true);
      enable.setEnabled(false);
    }

    vPanel.add(cannotHTML);
    buttonGrid.setWidget(0, 0, getNamedValue(ss, CURRENT_SELECTION_NOTIFICATION) === NOT_STARTED ? enable : unenable);
    buttonGrid.setWidget(0, 1, cancel);
    buttonGrid.setWidget(0, 2, waitHTML);
    vPanel.add(buttonGrid);
    cPanel.add(vPanel);
    app.add(cPanel);

    app.setHeight(height);

    handler.addCallbackElement(vPanel);

    var clientHandler = app.createClientHandler()
      .forTargets([enable, unenable, cancel, acceptedButton, notAcceptedButton, ]).setEnabled(false)
      .forTargets(waitHTML).setVisible(true);

    unenable.addClickHandler(clientHandler);
    enable.addClickHandler(clientHandler);
    cancel.addClickHandler(clientHandler);

    ss.show(app);
  } catch(e) {
    log('selectionNotification:error:' + catchToString(e));
  }
}

function selectionNotificationButtonAction(e) {
  var app = UiApp.getActiveApplication(),
    ss = SpreadsheetApp.getActiveSpreadsheet();

  if(e.parameter.source === ENABLE) {
    log(SELECTION_NOTIFICATION + ' set to:' + PENDING);
    setNamedValue(ss, CURRENT_SELECTION_NOTIFICATION, PENDING);
    ss.toast("Selection Notification has been queued.", "INFORMATION", 5);
  } else if(e.parameter.source === UNENABLE) {
    log(SELECTION_NOTIFICATION + ' set to:' + NOT_STARTED);
    setNamedValue(ss, CURRENT_SELECTION_NOTIFICATION, NOT_STARTED);
    ss.toast("Selection Notification has been unqueued.", "INFORMATION", 5);
  } else {
    ss.toast("Canceling operation.", "WARNING", 5);
  }
  return app.close();
}

function adHocEmailTemplate() {
  selectionNotification(AD_HOC_EMAIL);
}

function adHocEmail() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    pleaseWait(ss);

    var app = UiApp.createApplication(),
      cPanel = app.createCaptionPanel("Queue " + AD_HOC_EMAIL).setId('cPanel'),
      vPanel = app.createVerticalPanel(),
      handler = app.createServerHandler("adHocEmailButtonAction"),
      enable = app.createButton(ENABLE, handler).setId(normalizeHeader(ENABLE)),
      unenable = app.createButton(UNENABLE, handler).setId(normalizeHeader(UNENABLE)),
      cancel = app.createButton(CANCEL, handler).setId(normalizeHeader(CANCEL)),
      adHocEmailButton = app.createButton(AD_HOC_EMAIL, handler).setId(normalizeHeader(AD_HOC_EMAIL)),
      adHocEmailPLabel = app.createLabel(PROCESSING).setVisible(false).setId(normalizeHeader(AD_HOC_EMAIL + ' ' + PLABEL)),
      waitLabel = app.createLabel(PROCESSING).setVisible(false),
      buttonGrid = app.createGrid(1, 3),
      adHocEamilGrid = app.createGrid(1, 3);

    vPanel.add(app.createHTML('<p>The system can send an <b>' + AD_HOC_EMAIL + '</b> to all submitting filmmakers.</p><p>This can be used to inform the filmmakers of changes in the circumstances of the festival, such as a change in the close of submission date.</p>'));

    adHocEamilGrid.setWidget(0, 0, app.createLabel('Edit and test ' + AD_HOC_EMAIL + ':'));


    adHocEamilGrid.setWidget(0, 1, adHocEmailButton);
    adHocEamilGrid.setWidget(0, 2, adHocEmailPLabel);
    vPanel.add(adHocEamilGrid);

    vPanel.add(app.createHTML('<p>When you enable <b>' + AD_HOC_EMAIL + '</b>, the system will start to process the submissions after midnight of that day, in their submission order.</p><p>If you unenable <b>' + AD_HOC_EMAIL + '</b> before midnight, the emails will not be sent and <b>' + AD_HOC_EMAIL + '</b> will have been canceled.</p><p>NOTE: submissions with <b>' + STATUS + '</b> set to <b>' + PROBLEM + '</b> will not receive an email.</p><br/>'));

    buttonGrid.setWidget(0, 0, getNamedValue(ss, CURRENT_AD_HOC_EMAIL) === NOT_STARTED ? enable : unenable);
    buttonGrid.setWidget(0, 1, cancel);
    buttonGrid.setWidget(0, 2, waitLabel);
    vPanel.add(buttonGrid);

    cPanel.add(vPanel);
    app.add(cPanel);

    handler.addCallbackElement(vPanel);

    var click = app.createClientHandler()
      .forTargets([enable, unenable, cancel, adHocEmailButton]).setEnabled(false)
      .forTargets(waitLabel).setVisible(true);
    unenable.addClickHandler(click);
    enable.addClickHandler(click);
    cancel.addClickHandler(click);

    var adHocClick = app.createClientHandler()
      .forTargets([enable, unenable, cancel, adHocEmailButton]).setEnabled(false)
      .forTargets(adHocEmailPLabel).setVisible(true);
    adHocEmailButton.addClickHandler(adHocClick);

    ss.show(app);
  } catch(e) {
    log('adHocEmail:error:' + catchToString(e));
  }
}

function adHocEmailButtonAction(e) {
  var app = UiApp.getActiveApplication(),
    ss = SpreadsheetApp.getActiveSpreadsheet();
  log('adHocEmailButtonAction:e.parameter.source:' + e.parameter.source);
  if(e.parameter.source === normalizeHeader(UNENABLE)) {
    setNamedValue(ss, CURRENT_AD_HOC_EMAIL, NOT_STARTED);
    ss.toast("Ad Hoc Email has been unqueued.", "INFORMATION", 5);
    app.close();
  } else if(e.parameter.source === normalizeHeader(ENABLE)) {
    setNamedValue(ss, CURRENT_AD_HOC_EMAIL, PENDING);
    ss.toast("Ad Hoc Email has been queued.", "INFORMATION", 5);
    app.close();
  } else if(e.parameter.source === normalizeHeader(AD_HOC_EMAIL)) {
    adHocEmailTemplate();
  } else {
    app.close();
    ss.toast("Canceling operation.", "WARNING", 5);
  }
  return app;
}

function setup() {
  try {
    if(TESTING || !ScriptProperties.getProperty("initialised")) {
      if(!TESTING) {
        ScriptProperties.setProperty("initialised", "initialised");
      }

      log("setup start.");
      var filmSheetColumnWidth = [{
        header: TIMESTAMP,
        width: 120
      }, {
        header: LAST_CONTACT,
        width: 120
      }, {
        header: STATUS,
        width: 93
      }, {
        header: CONFIRMATION,
        width: 92
      }, {
        header: SELECTION,
        width: 85
      }, {
        header: SCORE,
        width: 48
      }, {
        header: FILM_ID,
        width: 40
      }, {
        header: FIRST_NAME,
        width: 120
      }, {
        header: LAST_NAME,
        width: 120
      }, {
        header: EMAIL,
        width: 52
      }, {
        header: TITLE,
        width: 166
      }, {
        header: LENGTH,
        width: 52
      }, {
        header: COUNTRY,
        width: 120
      }, {
        header: YEAR,
        width: 37
      }, {
        header: GENRE,
        width: 120
      }, {
        header: WEBSITE,
        width: 120
      }, {
        header: SYNOPSIS,
        width: 778
      }, {
        header: CAST_AND_CREW,
        width: 463
      }, {
        header: FESTIVAL_SELECTION_AND_AWARDS,
        width: 503
      }, {
        header: BEST_LOCAL_FILM,
        width: 463
      }, {
        header: BEST_LOCAL_FILM_ELIGIBILITY,
        width: 463
      }, {
        header: CONFIRM,
        width: 120
      }, {
        header: COMMENTS,
        width: 120
      }],

        templateSheetColumnWidth = [{
          header: TEMPLATE_NAME,
          width: 235
        }, {
          header: TEMPLATE,
          width: 410
        }],

        // data for Options & Settings sheet
        states = [
          [NO_MEDIA + " " + NOT_CONFIRMED, [NO_MEDIA + " " + NOT_CONFIRMED, "IndianRed"]],
          [NO_MEDIA + " " + CONFIRMED, [NO_MEDIA + " " + CONFIRMED, "LightPink"]],
          [MEDIA_PRESENT + " " + NOT_CONFIRMED, [MEDIA_PRESENT + " " + NOT_CONFIRMED, "DarkSeaGreen"]],
          [MEDIA_PRESENT + " " + CONFIRMED, [MEDIA_PRESENT + " " + CONFIRMED, "LightGreen"]],
          [PROBLEM, [PROBLEM, "Orange"]],
          [SELECTED, [SELECTED, "Gold"]],
          [NOT_SELECTED, [NOT_SELECTED, "LightSteelBlue"]]
        ],

        currentDate = new Date(),
        closeDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + 7, currentDate.getDate()),
        eventDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + 9, currentDate.getDate()),

        festivalData = [
          [FESTIVAL_NAME, [FESTIVAL_NAME, 'The Awesome Short Film Festival']],
          [FESTIVAL_WEBSITE, [FESTIVAL_WEBSITE, 'http://festivalwebsite']],
          [CLOSE_OF_SUBMISSION, [CLOSE_OF_SUBMISSION, closeDate]],
          [RELEASE_LINK, [RELEASE_LINK, 'http://festivalwebsite/releaseform']],
          [EVENT_DATE, [EVENT_DATE, eventDate]],
          [DAYS_BEFORE_REMINDER, [DAYS_BEFORE_REMINDER, 28]],
          [ENABLE_REMINDER, [ENABLE_REMINDER, ENABLED]],
          [ENABLE_CONFIRMATION, [ENABLE_CONFIRMATION, ENABLED]]
        ],

        testData = [
          [TEST_FIRST_NAME, [TEST_FIRST_NAME, 'Maya']],
          [TEST_LAST_NAME, [TEST_LAST_NAME, 'Deren']],
          [TEST_TITLE, [TEST_TITLE, 'Meshes of the Afternoon']],
          [TEST_EMAIL, [TEST_EMAIL, Session.getActiveUser().getEmail()]]
        ],

        internals = [
          [FIRST_FILM_ID, [FIRST_FILM_ID, 2]],
          [CURRENT_FILM_ID, [CURRENT_FILM_ID, -1]],
          [CURRENT_SELECTION_NOTIFICATION, [CURRENT_SELECTION_NOTIFICATION, NOT_STARTED]],
          [CURRENT_AD_HOC_EMAIL, [CURRENT_AD_HOC_EMAIL, NOT_STARTED]]
        ],

        // data for Template sheet
        templates = [
          [TEMPLATE_NAME, TEMPLATE],
          [SUBMISSION_CONFIRMATION_SUBJECT_LINE, [SUBMISSION_CONFIRMATION_SUBJECT_LINE, '${"Festival Name"} Submissions Confirmation: ${"Title"}']],
          [SUBMISSION_CONFIRMATION_BODY, [SUBMISSION_CONFIRMATION_BODY, 'Dear ${"First Name"},\n\nthank you for submitting your film "${"Title"}" to ${"Festival Name"}.\n\nNow you need to download and sign the permission slip that can be found on our website via the link below.\n${"Release Link"}\n\nThe signed permission slip must be mailed together with a DVD of your film to our festival office. If we do not receive these by the close of submission on ${"Close Of Submission"}, we will not be able to consider your film for selection. The address of our festival office is given on the permission slip.\n\nYou will receive an email confirmation of the receipt of your film and after the close of submissions you will be advised by emailed whether or not your film has been selected.\n\nMany Thanks\n${"Festival Name"}\n${"Festival Website"}']],
          [RECEIPT_CONFIMATION_SUBJECT_LINE, [RECEIPT_CONFIMATION_SUBJECT_LINE, '${"Festival Name"} Receipt Confimation: ${"Title"}']],
          [RECEIPT_CONFIMATION_BODY, [RECEIPT_CONFIMATION_BODY, 'Dear ${"First Name"}\n\nthis is to confirm the receipt of the DVD of your film "${"Title"}" by ${"Festival Name"}.\n\nSelection will take place after the close of submission on ${"Close Of Submission"}. You will be informed after this date if your film has been selected.\n\n${"Festival Name"} runs on the evening of ${"Event Date"}.\n\nMany thanks for your film submission.\n${"Festival Name"}\n${"Festival Website"}']],
          [REMINDER_SUBJECT_LINE, [REMINDER_SUBJECT_LINE, '${"Festival Name"} Submission Reminder: ${"Title"}']],
          [REMINDER_BODY, [REMINDER_BODY, 'Dear ${"First Name"}\n\nthank you for submitting your film "${"Title"}" to ${"Festival Name"}.\n\nWe are still awaiting a DVD of your film together with a signed permission slip. If we do not receive these at our festival office by the close of submission on ${"Close Of Submission"}, we will not be able to consider your film for selection.\n\nYou can download the permission slip from our website via the link below.\n${"Release Link"}\nThe address of our festival office is given on the permission slip.\n\nWe will confirm the receipt of your film by email and after the close of submissions you will be advised by emailed whether or not your film has been selected.\n\n${"Festival Name"} runs on the evening of ${"Event Date"}.\n\nMany Thanks\n${"Festival Name"}\n${"Festival Website"}']],
          [NOT_ACCEPTED_SUBJECT_LINE, [NOT_ACCEPTED_SUBJECT_LINE, '${"Festival Name"}: ${"Title"}']],
          [NOT_ACCEPTED_BODY, [NOT_ACCEPTED_BODY, 'Dear ${"First Name"},\n\nit has been very difficult to make the selection for ${"Festival Name"} this year. We have received an enormous number of accomplished short films from all over the world and we have had to make some difficult choices. We are sorry to tell you that your film ${"Title"} has not been selected for ${"Festival Name"}.\n\nIt is extremely important to point out that this years film selection is intended to be a broad representative sample from the huge number of submissions we have received. Many film merited being in the selection, but could not be include due to the limited screening time of the festival.\n\n\We wish to thank you for your film submission and to encourage you to submit to ${"Festival Name"} next year.\n\nMany Thanks\n${"Festival Name"}\n${"Festival Website"}']],
          [ACCEPTED_SUBJECT_LINE, [ACCEPTED_SUBJECT_LINE, '${"Festival Name"}: ${"Title"}']],
          [ACCEPTED_BODY, [ACCEPTED_BODY, 'Dear ${"First Name"},\n\nit has been very difficult to make the selection for ${"Festival Name"} this year. We have received an enormous number of accomplished short films from all over the world and we have had to make some difficult choices. We are very happy to tell you that your film ${"Title"} has been selected for ${"Festival Name"}.\n\nDue to limit seating at our festival venue, we can only give two free tickets to each attending filmmaker.\n\nWe will be in touch shortly to confirm the details of the screening.\n\nMany Thanks\n${"Festival Name"}\n${"Festival Website"}']],
          [AD_HOC_EMAIL_SUBJECT_LINE, [AD_HOC_EMAIL_SUBJECT_LINE, '${"Festival Name"}: ${"Title"}']],
          [AD_HOC_EMAIL_BODY, [AD_HOC_EMAIL_BODY, 'Dear ${"First Name"},\n\njust to thank you for submitting your film "${"Title"}" to ${"Festival Name"}.\n\nAnd to let you know that ${"Festival Name"} is still completely awesome.\n\nMany Thanks\n${"Festival Name"}\n${"Festival Website"}']]
        ],
        itemData, sheet;


      // build online submission form
      var form = FormApp.create(FORM_TITLE).setConfirmationMessage(FORM_RESPONSE);
      for(var itemIndex = 0; itemIndex < FORM_DATA.length; itemIndex++) {
        itemData = FORM_DATA[itemIndex];
        if(itemData === "section") {
          form.addSectionHeaderItem();
        } else {
          var formItem = form[itemData.type]().setTitle(itemData.title).setHelpText(itemData.help).setRequired(itemData.required ? true : false);
          if(itemData.type === "addCheckboxItem") {
            formItem.setChoices([formItem.createChoice(itemData.choice)]);
          }
          if(itemData.type === "addMultipleChoiceItem") {
            formItem.setChoiceValues(itemData.choices);
          }
        }
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet(),
        ssId = ss.getId(),
        formId = form.getId();

      // find parent folder of spreadsheet and move form to that folder
      var ssParentFolder = DocsList.getFileById(ssId).getParents()[0],
        formFile = DocsList.getFileById(formId),
        creationFolder = formFile.getParents()[0], // expect this to be ROOT and the same as the folder that formFolder is created into
        formFolder = DocsList.createFolder(FORM_FOLDER);
      formFolder.addToFolder(ssParentFolder);
      formFolder.removeFromFolder(creationFolder);
      formFile.addToFolder(formFolder);
      formFile.removeFromFolder(creationFolder);
      log('Built online submission form.');

      // initialise spreadsheet and set spreadsheet as the destination of the form

      form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

      // Need this for testing.
      // Otherwise 'Form Responses 1' not present when looked for!!!
      // Why does this only break testing!!!
      SpreadsheetApp.flush();
      
      // create the FILM_SUBMISSIONS_SHEET, OPTIONS_SETTINGS_SHEET and TEMPLATE_SHEET sheet
      ss.getSheetByName('Form Responses 1').setName(FILM_SUBMISSIONS_SHEET);
      ss.insertSheet(OPTIONS_SETTINGS_SHEET, 2);
      ss.insertSheet(TEMPLATE_SHEET, 3);
      ss.deleteSheet(ss.getSheetByName('Sheet1'));

      // add extra columns to FILM_SUBMISSIONS_SHEET
      sheet = ss.getSheetByName(FILM_SUBMISSIONS_SHEET);
      for(var i = 0; i < ADDITIONAL_COLUMNS.length; i++) {
        columnIndex = ADDITIONAL_COLUMNS[i].columnIndex;
        if(columnIndex === 'end') {
          columnIndex = sheet.getLastColumn();
        }
        sheet.insertColumnAfter(columnIndex);
        sheet.getRange(1, columnIndex + 1, 1, 1).setValue(ADDITIONAL_COLUMNS[i].title);
      }

      //add templates to TEMPLATE_SHEET)
      saveData(ss, ss.getSheetByName(TEMPLATE_SHEET).getRange(1, 1, 1, 1), templates, TEMPLATE_DATA, 'mistyrose');

      //initialise OPTIONS_SETTINGS_SHEET
      var data = [{
        name: FESTIVAL_DATA,
        data: festivalData
      }, {
        name: TEST_DATA,
        data: testData
      }, {
        name: COLOR_DATA,
        data: states
      }, {
        name: INTERNALS,
        data: internals
      }];
      for(i = 0; i < data.length; i++) {
        var dataItem = data[i];
        saveData(ss, ss.getSheetByName(OPTIONS_SETTINGS_SHEET).getRange(1, 2 * (+i) + 1, 1, 1), dataItem.data, dataItem.name, (+i) % 2 ? 'aliceblue' : 'mistyrose');

      }

      // update menu
      ss.updateMenu(FILM_SUBMISSION, MENU_ENTRIES);

      setColumnWidth(ss, FILM_SUBMISSIONS_SHEET, filmSheetColumnWidth);
      setColumnWidth(ss, TEMPLATE_SHEET, templateSheetColumnWidth);

      log('Built spreadsheet.');

      if(!TESTING) {
        settingsOptions(); // let user set options and settings
      }

      // Only need the triggers if we are not testing.
      if(!TESTING) {
        // enable film submission processing on form submission
        ScriptApp.newTrigger("hProcessSubmission")
          .forSpreadsheet(ss)
          .onFormSubmit()
          .create();
        
      // enable reminders and confirmations, sent out after midnight each day
      ScriptApp.newTrigger("hReminderConfirmation")
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .create();
      
        log('Added triggers.');
        log("setup end.");
      }
    }
  } catch(e) {
    log('setup:error:' + catchToString(e));
  }
}

function setColumnWidth(ss, sheetName, columnWidths) {
  try {
    var sheet = ss.getSheetByName(sheetName),
      headers = sheet.getRange(1, 1, 1, sheet.getDataRange().getLastColumn()).getValues()[0],
      width = {};
    for(var columnIndex = 0; columnIndex < columnWidths.length; columnIndex++) {
      width[columnWidths[columnIndex].header] = columnWidths[columnIndex].width;
    }

    for(columnIndex in headers) {
      if(width[headers[columnIndex]]) {
        sheet.setColumnWidth(1 + (+columnIndex), width[headers[columnIndex]]);
      }
    }
    SpreadsheetApp.flush();
  } catch(e) {
    log('setColumnWidth:error:' + catchToString(e));
  }
}

function log(info) {
  try {
    var logFileId = ScriptProperties.getProperty(normalizeHeader(LOG_FILE)),
      logFile;

    if(logFileId) {
      logFile = DocumentApp.openById(logFileId);
    }

    if(!logFile) {
      logFile = DocumentApp.create(LOG_FILE);
      var ss = SpreadsheetApp.getActiveSpreadsheet(),
        ssParentFolder = DocsList.getFileById(ss.getId()).getParents()[0],
        logFileFile = DocsList.getFileById(logFile.getId()),
        creationFolder = logFileFile.getParents()[0], // expect this to be ROOT and the same as the folder that logFolder is created into
        logFolder = DocsList.createFolder(LOG_FOLDER);
      logFolder.addToFolder(ssParentFolder);
      logFolder.removeFromFolder(creationFolder);
      logFileFile.addToFolder(logFolder);
      logFileFile.removeFromFolder(creationFolder);
      ScriptProperties.setProperty(normalizeHeader(LOG_FILE), logFile.getId());
      logFile.getBody().appendParagraph('Created log file!');
    }
    logFile.getBody().appendParagraph((new Date()) + ": " + info);
  } catch(e) {
    log('log:error:' + catchToString(e));
  }
}

function catchToString(err) {
  var errInfo = "Catched something:\n";
  for(var prop in err) {
    if(err.hasOwnProperty(prop)) {
      errInfo += "  property: " + prop + "\n    value: [" + err[prop] + "]\n";
    }
  }
  errInfo += "  toString(): " + " value: [" + err.toString() + "]";
  return errInfo;
}

function mergeAndSend(ss, submission, templateName, emailQuotaRemaining) {
  try {
    emailQuotaRemaining = emailQuotaRemaining ? emailQuotaRemaining : 0; // Set to 0 if not supplied i.e. it dosnt matter in this case

    submission[normalizeHeader(FESTIVAL_NAME)] = getNamedValue(ss, FESTIVAL_NAME);
    submission[normalizeHeader(CLOSE_OF_SUBMISSION)] = prettyPrintDate(getNamedValue(ss, CLOSE_OF_SUBMISSION));
    submission[normalizeHeader(EVENT_DATE)] = prettyPrintDate(getNamedValue(ss, EVENT_DATE));
    submission[normalizeHeader(FESTIVAL_WEBSITE)] = getNamedValue(ss, FESTIVAL_WEBSITE);
    submission[normalizeHeader(RELEASE_LINK)] = getNamedValue(ss, RELEASE_LINK);

    var subjectTemplate = getNamedValue(ss, templateName + ' ' + SUBJECT_LINE),
      bodyTemplate = getNamedValue(ss, templateName + ' ' + BODY);

    if(submission[normalizeHeader(EMAIL)]) {
      MailApp.sendEmail(submission[normalizeHeader(EMAIL)], fillInTemplateFromObject(subjectTemplate, submission), fillInTemplateFromObject(bodyTemplate, submission));
      emailQuotaRemaining -= 1;
      if(TESTING) {
        var templatesTesting = ScriptProperties.getProperty(normalizeHeader(TEMPLATES_TESTING));
        if(templatesTesting) {
          templatesTesting += ',' + templateName;
        } else {
          templatesTesting = templateName;
        }
        ScriptProperties.setProperty(normalizeHeader(TEMPLATES_TESTING), templatesTesting);
      }
      log('mergeAndSend:email sent');
    } else {
      log('mergeAndSend:error:no email found!');
    }


  } catch(e) {
    log('mergeAndSend:error:' + catchToString(e));
  }
  return emailQuotaRemaining;
}

function hProcessSubmission(event) {
  log('hProcessSubmission start');
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // load data in advance
    loadData(ss, FESTIVAL_DATA);
    loadData(ss, COLOR_DATA);
    loadData(ss, TEMPLATE_DATA);
    loadData(ss, INTERNALS);

    var filmSheet = ss.getSheetByName(FILM_SUBMISSIONS_SHEET),
      eventRow = event.range.getRow(),
      currentFilmID = -1,
      emailQuotaRemaining = MailApp.getRemainingDailyQuota(),
      filmSubmission, filmSubmissionData, rows, length, headersOfInterest, minColumn, maxColumn, minMaxColumns, lastContact, sendEmail, confirmation, color;
  
    sendEmail = MIN_QUOTA < emailQuotaRemaining;
    if(!sendEmail) {
      log('hProcessSubmission:ran out of email quota, (MIN_QUOTA, emailQuotaRemaining):(' + MIN_QUOTA + ',' + emailQuotaRemaining + ')');
    }

    // Insert LAST_CONTACT,FILM_ID and STATUS into film submission. Also make sure LENGTH has correct form
    // Use smallest range that will work with one call to getRowsData and setRowsData
    headersOfInterest = [TIMESTAMP, LAST_CONTACT, FILM_ID, STATUS, CONFIRMATION, SELECTION, LENGTH];
    minMaxColumns = findMinMaxColumns(filmSheet, headersOfInterest);
    
    
    var lastContactIndex = minMaxColumns.indices[normalizeHeader(LAST_CONTACT)];
    var lastContactCell = filmSheet.getRange(eventRow, lastContactIndex, 1, 1);
    lastContact = lastContactCell.getValue();
    
    if (lastContact instanceof Date) {
      throw {message: 'ERROR: this submission has already been processed!'};
    }
    
   // Update Find and update Film ID for current submission
    if(eventRow === 2) {
      // first submission
      currentFilmID = getNamedValue(ss, FIRST_FILM_ID);
    } else {
      currentFilmID = getNamedValue(ss, CURRENT_FILM_ID);
      currentFilmID = currentFilmID + 1;
    }
    setNamedValue(ss, CURRENT_FILM_ID, currentFilmID);
    currentFilmID = ID + pad(currentFilmID);
    log('hProcessSubmission:processing film submission ' + currentFilmID);
  
    maxColumn = minMaxColumns.max;
    minColumn = minMaxColumns.min;
    filmSubmission = event.range;
    filmSubmissionData = {};

    var properties = [];
    for(var property in event.namedValues) {
      if(event.namedValues.hasOwnProperty(property)) {
        properties.push(property);
        filmSubmissionData[normalizeHeader(property)] = event.namedValues[property][0];// Suprise . . . the value is an array of length 1!!!
      }
    }

    length = normaliseAndValidateDuration(filmSubmissionData[normalizeHeader(LENGTH)]);
    if(length) {
      filmSubmissionData[normalizeHeader(LENGTH)] = length;
    }
    filmSubmissionData[normalizeHeader(STATUS)] = NO_MEDIA;
    filmSubmissionData[normalizeHeader(FILM_ID)] = currentFilmID;
    if(sendEmail) {
      lastContact = filmSubmissionData[normalizeHeader(TIMESTAMP)];
      confirmation = CONFIRMED;
    } else {
      lastContact = NO_CONTACT;
      confirmation = NOT_CONFIRMED;
    }
    filmSubmissionData[normalizeHeader(CONFIRMATION)] = confirmation;
    filmSubmissionData[normalizeHeader(SELECTION)]    = NOT_SELECTED;
    filmSubmissionData[normalizeHeader(LAST_CONTACT)] = lastContact;
    filmSubmissionData[normalizeHeader(SCORE)]        = -1;
    

    rows = [];
    rows[0] = filmSubmissionData;
    setRowsData(filmSheet, rows, filmSheet.getRange(1, minColumn, 1, maxColumn - minColumn + 1), eventRow);

    //set status color on film submission
    color = findStatusColor(ss, NO_MEDIA, confirmation, NOT_SELECTED);
    filmSheet.getRange(eventRow, 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);

    if(sendEmail) {
      // send confirmation email
      mergeAndSend(ss, filmSubmissionData, SUBMISSION_CONFIRMATION);
    }
  } catch(e) {
    log('hProcessSubmission:error:' + catchToString(e));
  }
  lock.releaseLock();
  log('hProcessSubmission end');
}

function prettyPrintDate(date) {
  if(!date) {
    return null;
  } else {
    return ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][date.getDay()] + " the " + (function (d) {
      var s = d.toString(),
        l = s[s.length - 1];
      return s + (['th', 'th', 'th'][d - 11] || ['st', 'nd', 'rd'][l - 1] || 'th');
    })(date.getDate()) + " of " + ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][date.getMonth()] + ", " + date.getFullYear();
  }
}

function normaliseAndValidateDuration(length) {
  log('normaliseAndValidateDuration:start:length:"'+length+'"');
  function babylonian(myNumber) {
    return 0 <= myNumber && myNumber <= 59;
  }
  var returnLength, mymatch;
  try {
    var hours = '0'
      minutes = '0',
      seconds = '0',
      myregexp = /^[^\d]*([\d]+)[^\d]*$|^[^\d]*([\d]+)[^\d]+([\d]+)[^\d]*$|^[^\d]*([\d]+)[^\d]+([\d]+)[^\d]+([\d]+)[^\d]*$/i;
      
    mymatch = myregexp.exec(length);
    if(mymatch) {
      if (mymatch[1]) {
        // One number found, assume it is minutes
        minutes = mymatch[1];
      } else if (mymatch[2]) {
        // NOTE: due to regexp mymatch[2] is true ifandonlyif mymatch[3] is true
        // Found two numbers, assume they are minutes and seconds
        minutes = mymatch[2];
        seconds = mymatch[3];
      } else if (mymatch[4]) {
        // NOTE: due to regexp mymatch[4] is true ifandonlyif mymatch[5] is true ifandonlyif mymatch[6] is true
        // Found three numbers, assume they are hours, minutes and seconds
        hours   = mymatch[4];
        minutes = mymatch[5];
        seconds = mymatch[6];
      }
      
      if(babylonian(hours) && babylonian(minutes) && babylonian(seconds)) {
        returnLength = ('0' + hours).slice(-2)+':'+('0' + minutes).slice(-2)+':'+('0' + seconds).slice(-2);
      }
    }
  } catch(e) {
    log('normaliseAndValidateDuration:error:' + catchToString(e));
  }
  log('normaliseAndValidateDuration:end');
  return returnLength ? returnLength : length;
}

function pad(number) {
  if(number > 0) {
    return("000000" + number).slice(-PAD_NUMBER);
  } else {
    return("\\d\\d\\d\\d\\d\\d").slice(-2 * PAD_NUMBER);
  }
}

function setNamedValueInCache(name, value) {
  CACHE[name] = value;
}

function getNamedValueFromCache(name) {
  return CACHE[name];
}

function getNamedValue(ss, name) {
  var value;
  try {
    name = normalizeHeader(name);
    value = getNamedValueFromCache(name);
    if(!value) {
      value = ss.getRangeByName(name).getValue();
      setNamedValueInCache(name, value);
    }
  } catch(e) {
    log('getNamedValue:name:' + name + ':error:' + catchToString(e));
  }
  return value;
}

//NOTE: possible to set a value for a named range that does not exist.
//It will be set in the cache just for this single run of the script.
function setNamedValue(ss, name, value) {
  try {
    name = normalizeHeader(name);
    var range = null;
    
    // shouldnt need to do this!!!
    try{
      range = ss.getRangeByName(name);
    } catch(e) {
      log('getRangeByName:throw exception:'+ catchToString(e));
    }
    
    
    if(range) {
      range.setValue(value);
    } else {
      log('setNamedValue: no range with name:' + name + ' found!');
    }
    setNamedValueInCache(name, value);
  } catch(e) {
    log('setNamedValue:name:' + name + ':error:' + catchToString(e));
  }
}

// Load data into CACHE so that getNamedRange will later get it from CACHE and not have to make an individual getValue on a spreadsheet cell
function loadData(ss, name) {
  var items = [];
  try {
    log('loadData:loading data block:' + name);
    name = normalizeHeader(name);
    var loaded = getNamedValueFromCache(name);
    if(!loaded) {
      var item;

      // Named range assumed to have width of 2.
      // Each row is a name value pair.
      var values = ss.getRangeByName(name).getValues();
      for(var i = 0; i < values.length; i++) {
        item = values[i];
        items.push(item[0]);
        CACHE[normalizeHeader(item[0])] = item[1];
      }
      setNamedValueInCache(name, true);
    } else {
      log('loadData:date block has already been loaded!');
    }
  } catch(e) {
    log('loadData:error:' + catchToString(e));
  }
  return items;
}

// sets a range with array values
// if an entry in values is a pair, the first member will be the name of a cell and the second member will be its value
function saveData(ss, startCell, values, name, color) {
  try {
    var rowIndex, columnIndex, rowValue, value, currentCell;

    var range = startCell.offset(0, 0, values.length, 2);
    range.setBackground(color);
    ss.setNamedRange(normalizeHeader(name), range);

    for(rowIndex = 0; rowIndex < values.length; rowIndex++) {
      rowValue = values[rowIndex];
      for(columnIndex = 0; columnIndex < rowValue.length; columnIndex++) {
        value = rowValue[columnIndex];
        currentCell = startCell.offset(rowIndex, columnIndex);
        if(Array.isArray(value)) { //expect array to have 2 elements
          ss.setNamedRange(normalizeHeader(value[0]), currentCell);
          value = value[1];
        }
        currentCell.setValue(value);
      }
    }
  } catch(e) {
    log('saveData:error:' + catchToString(e));
  }
}

////////////////////////////////////////////////////////
// We expect hReminderConfirmation to do the following:
// if(Ad Hoc Email is pending or in progress){
//    Continue with Ad Hoc Email.
//    Stop when all submissions processed
//    or daily mail allocation is down to MIN_QUOTA.
// } else if (not yet reached Close Of Submissions) {
//    While daily mail allocation is greater than MIN_QUOTA for each submission
//       if submission is not confirmed {
//          send submission confirmation
//       } else if media present but not confirmed {
//          send receipt confirmation
//       } else if (media not present and days from last contact is large
//                 enough for a reminder) {
//          send reminder
//       }
// } else if (Notification Confirmation is pending or in progress) {
//    Continue with Notification Confirmation.
//    Stop when all submissions processed
//    or daily mail allocation is down to MIN_QUOTA.
// }
///////////////////////////////////////////////////////////////////////////////
function hReminderConfirmation() {
  try {
    log('hReminderConfirmation start');
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    loadData(ss, FESTIVAL_DATA);
    loadData(ss, TEMPLATE_DATA);
    loadData(ss, INTERNALS);
    loadData(ss, COLOR_DATA);

    var emailQuotaRemaining = MailApp.getRemainingDailyQuota(),
      enabledReminder = getNamedValue(ss, ENABLE_REMINDER) === ENABLED,
      enableConfirmation = getNamedValue(ss, ENABLE_CONFIRMATION) === ENABLED,
      filmSheet = ss.getSheetByName(FILM_SUBMISSIONS_SHEET),
      lastRow = filmSheet.getLastRow(),
      currentDate = new Date(),
      closeOfSubmission = getNamedValue(ss, CLOSE_OF_SUBMISSION),
      daysBeforeReminder = getNamedValue(ss, DAYS_BEFORE_REMINDER),
      daysTillClose = diffDays(closeOfSubmission, currentDate);

    log('hReminderConfirmation:enabledReminder: ' + enabledReminder + '\nenableConfirmation: ' + enableConfirmation + '\ncloseOfSubmission: ' + closeOfSubmission + '\ndaysBeforeReminder: ' + daysBeforeReminder + '\ndaysTillClose: ' + daysTillClose + '\nemailQuotaRemaining: ' + emailQuotaRemaining);

    if(lastRow > 1) {
      log('hReminderConfirmation:processing');
      var headersRange = filmSheet.getRange(1, 1, 1, filmSheet.getDataRange().getLastColumn()),
        headers = normalizeHeaders(headersRange.getValues()[0]),
        confirmationColumnIndex = headers.indexOf(normalizeHeader(CONFIRMATION)) + 1,
        lastContactColumnIndex = headers.indexOf(normalizeHeader(LAST_CONTACT)) + 1,
        minMaxColumns = findMinMaxColumns(filmSheet, [LAST_CONTACT, STATUS, CONFIRMATION, SELECTION, FIRST_NAME, EMAIL, TITLE]),
        height = filmSheet.getDataRange().getLastRow() - 1,
        range = filmSheet.getRange(2, minMaxColumns.min, height, minMaxColumns.max - minMaxColumns.min + 1),
        submissions = getRowsData(filmSheet, range, 1),
        currentAdHocEmail = getNamedValue(ss, CURRENT_AD_HOC_EMAIL),
        currentSelectionNotification = getNamedValue(ss, CURRENT_SELECTION_NOTIFICATION),
        color = "",
        i, submission, status, confirmation, selection, filmId, filmIdFormat, processNextSubmission, setToFinishedAtEnd, matches;

      if(currentAdHocEmail !== NOT_STARTED) {
        // AD HOC EMAIL
        log('hReminderConfirmation:processing Ad Hoc Mail');
        filmIdFormat = FILM_ID + pad(-1);
        processNextSubmission = currentAdHocEmail === PENDING;
        setToFinishedAtEnd = true;
        log('hReminderConfirmation:currentAdHocEmail:' + currentAdHocEmail);
        // Loop through all submissions until we come to the first submission after CURRENT_AD_HOC_EMAIL.
        // Start with first if currentAdHocEmail === PENDING. Normally this will be the case.
        for(i = 0; i < submissions.length; i++) {
          if(emailQuotaRemaining <= MIN_QUOTA) {
            log('hReminderConfirmation:ran out of email quota during ' + AD_HOC_EMAIL + ', (MIN_QUOTA, emailQuotaRemaining):(' + MIN_QUOTA + ',' + emailQuotaRemaining + '). Terminating ' + AD_HOC_EMAIL + ' early.');
            setToFinishedAtEnd = false;
            break; //stop sending emails for the day
          }

          submission = submissions[i];
          status = submission[normalizeHeader(STATUS)];
          filmId = submission[normalizeHeader(FILM_ID)];

          if(!filmId) {
            log('hReminderConfirmation:ERROR: exepected film id!');
            continue;
          }

          if(status === PROBLEM) {
            // Ingnore submissions with status PROBLEM
            continue;
          }

          matches = filmId.match(filmIdFormat, 'i');
          if(matches && matches.length > 0) {
            filmId = matches[0];
            if(processNextSubmission) {
              // send Ad Hoc Email
              emailQuotaRemaining = mergeAndSend(ss, submission, AD_HOC_EMAIL, emailQuotaRemaining);
              setNamedValue(ss, CURRENT_AD_HOC_EMAIL, filmId); //Update current
              log('hReminderConfirmation:Ad Hoc Mail:processed:' + filmId);
            } else {
              processNextSubmission = filmId === currentAdHocEmail;
              log('hReminderConfirmation:found currentAdHocEmail');
            }
          } else {
            log('hReminderConfirmation:ERROR: exepected film id!');
          }
        } // for (var i in submissions)

        if(setToFinishedAtEnd) {
          setNamedValue(ss, CURRENT_AD_HOC_EMAIL, NOT_STARTED); // Job done. Reset.
        }

        // if (currentAdHocEmail!==NOT_STARTED)
      } else if(currentDate < closeOfSubmission) {
        for(i = 0; i < submissions.length; i++) {
          if(emailQuotaRemaining <= MIN_QUOTA) {
            log('Ran out of email quota, (MIN_QUOTA, emailQuotaRemaining):(' + MIN_QUOTA + ',' + emailQuotaRemaining + ').');
            log('Terminating daily reminder and confirmation processing early.');
            break; //stop sending emails for the day
          }

          submission = submissions[i];
          status = submission[normalizeHeader(STATUS)];
          confirmation = submission[normalizeHeader(CONFIRMATION)];
          selection = submission[normalizeHeader(SELECTION)];
          if(status === PROBLEM) {
            // Ingnore submissions with status PROBLEM
            continue;
          }

          if(status === NO_MEDIA) {
            if(confirmation === NOT_CONFIRMED) {
              //Send submission confirmation. This should be a very rare case!
              log('hReminderConfirmation:no submission confirmation has been sent to ' + submission[normalizeHeader(FILM_ID)] + '.');

              emailQuotaRemaining = mergeAndSend(ss, submission, SUBMISSION_CONFIRMATION, emailQuotaRemaining);

              filmSheet.getRange(2 + (+i), lastContactColumnIndex, 1, 1).setValue(currentDate);
              filmSheet.getRange(2 + (+i), confirmationColumnIndex, 1, 1).setValue(CONFIRMED);
              color = findStatusColor(ss, status, CONFIRMED, selection);
              filmSheet.getRange(2 + (+i), 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);
              log('hReminderConfirmation:submission confirmation has now been sent to ' + submission[normalizeHeader(FILM_ID)] + ' and its ' + LAST_CONTACT + ' updated.');
            } else if(daysTillClose > daysBeforeReminder && enabledReminder) {
              // Need to send reminder to send Media and release form
              var lastContact = submission[normalizeHeader(LAST_CONTACT)];
              
              var daysSinceLastContact = diffDays(currentDate, lastContact);

              if(daysSinceLastContact >= daysBeforeReminder) {
                //send reminder
                log('It is ' + daysSinceLastContact + ' days since ' + LAST_CONTACT + ' with submission ' + submission[normalizeHeader(FILM_ID)] + ' which has status ' + NO_MEDIA + '.');

                emailQuotaRemaining = mergeAndSend(ss, submission, REMINDER, emailQuotaRemaining);

                filmSheet.getRange(2 + (+i), lastContactColumnIndex, 1, 1).setValue(currentDate);
                log('hReminderConfirmation:a reminder has been sent to submission ' + submission[normalizeHeader(FILM_ID)] + ' and its ' + LAST_CONTACT + ' updated.');
              }
            }

          } else if(status === MEDIA_PRESENT && confirmation === NOT_CONFIRMED && enableConfirmation) {
            //send confirmation of receipt of media
            log('hReminderConfirmation:submission ' + submission[normalizeHeader(FILM_ID)] + ' has status ' + MEDIA_PRESENT + ', ' + NOT_CONFIRMED + '.');

            emailQuotaRemaining = mergeAndSend(ss, submission, RECEIPT_CONFIMATION, emailQuotaRemaining);

            filmSheet.getRange(2 + (+i), lastContactColumnIndex, 1, 1).setValue(currentDate);
            filmSheet.getRange(2 + (+i), confirmationColumnIndex, 1, 1).setValue(CONFIRMED);
            color = findStatusColor(ss, status, CONFIRMED, selection);
            filmSheet.getRange(2 + (+i), 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);
            log('hReminderConfirmation:the status of submission ' + submission[normalizeHeader(FILM_ID)] + ' has been updated to ' + MEDIA_PRESENT + ', ' + CONFIRMED + ' and a receipt confirmation email has been sent.');
          }
        } // for(var i in submissions)
        //  else if (currentDate<closeOfSubmission)
      } else if(currentSelectionNotification !== NOT_STARTED) {
        // SELECTION NOTIFICATION
        log('hReminderConfirmation:processing Selection Notification');
        filmIdFormat = FILM_ID + pad(-1);
        processNextSubmission = currentSelectionNotification === PENDING;
        setToFinishedAtEnd = true;
        log('hReminderConfirmation:currentSelectionNotification:' + currentSelectionNotification);
        // Loop through all submissions until we come to the first submission after CURRENT_SELECTION_NOTIFICATION.
        // Start with first if CURRENT_SELECTION_NOTIFICATION === PENDING. Normally this will be the case.

        for(i = 0; i < submissions.length; i++) {
          if(emailQuotaRemaining <= MIN_QUOTA) {
            log('hReminderConfirmation:ran out of email quota during ' + SELECTION_NOTIFICATION + ', (MIN_QUOTA, emailQuotaRemaining):(' + MIN_QUOTA + ',' + emailQuotaRemaining + '). Terminating ' + SELECTION_NOTIFICATION + ' early.');
            setToFinishedAtEnd = false;
            break; //stop sending emails for the day
          }

          submission = submissions[i];
          filmId = submission[normalizeHeader(FILM_ID)];
          status = submission[normalizeHeader(STATUS)];
          selection = submission[normalizeHeader(SELECTION)];

          if(!filmId) {
            log('hReminderConfirmation:ERROR: exepected film id!');
            continue;
          }

          if(status === PROBLEM) {
            // Ingnore submissions with status PROBLEM
            continue;
          }

          matches = filmId.match(filmIdFormat, 'i');
          if(matches && matches.length > 0) {
            filmId = matches[0];
            if(processNextSubmission) {
              //send SELECTION NOTIFICATION

              if(selection === SELECTED) {
                emailQuotaRemaining = mergeAndSend(ss, submission, ACCEPTED, emailQuotaRemaining);
              } else {
                emailQuotaRemaining = mergeAndSend(ss, submission, NOT_ACCEPTED, emailQuotaRemaining);
              }

              setNamedValue(ss, CURRENT_SELECTION_NOTIFICATION, filmId); //Update current
              log('hReminderConfirmation:processed ' + filmId + ' as ' + selection);
            } else {
              processNextSubmission = filmId === currentSelectionNotification;
              log('hReminderConfirmation:found currentSelectionNotification');
            }
          } else {
            log('hReminderConfirmation:ERROR: exepected film id!');
          }
        } //for (var i in submissions)

        if(setToFinishedAtEnd) {
          setNamedValue(ss, CURRENT_SELECTION_NOTIFICATION, NOT_STARTED); //Job done. Reset.
        }
        // else if(currentSelectionNotification!==NOT_STARTED)
      }
      log('hReminderConfirmation end');
    } // if(lastRow>1)
  } catch(e) {
    log('ThReminderConfirmation:error:' + catchToString(e));
  }

}

//returns the difference in days of two dates
function diffDays(a, b) {
  return Math.floor((a.getTime() - b.getTime()) / (1000 * 3600 * 24));
}


function findMinMaxColumns(filmSheet, headersOfInterest) {
  var headersRange = filmSheet.getRange(1, 1, 1, filmSheet.getDataRange().getLastColumn()),
    headers = normalizeHeaders(headersRange.getValues()[0]),
    headersOfInterestIndices = [],
    minColumn, maxColumn, indices = {};

  headersOfInterest = normalizeHeaders(headersOfInterest);

  for(var i = 0; i < headersOfInterest.length; i++) {
    var header = headersOfInterest[i],
      index = headers.indexOf(header);

    indices[header] = index + 1;
    
    if(index !== -1) {
      headersOfInterestIndices.push(index);
    }
  }

  minColumn = Math.min.apply(Math, headersOfInterestIndices) + 1;
  maxColumn = Math.max.apply(Math, headersOfInterestIndices) + 1;
  return {
    min: minColumn,
    max: maxColumn,
    indices: indices
  };
}

/////////////////////////////////////////////////////////////////////////
// The code below is from the Simple Mail Merge tutorial which can be
// found here https://developers.google.com/apps-script/articles/mail_merge
/////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////
// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  if(templateVars) {
    // Replace variables from the template with the actual values from the data object.
    // If no value is available, replace with the empty string.
    for(var i = 0; i < templateVars.length; ++i) {
      // normalizeHeader ignores ${"} so we can call it directly here.
      var variableData = data[normalizeHeader(templateVars[i])];
      email = email.replace(templateVars[i], variableData || "");
    }
  }

  return email;
}

// setRowsData fills in one row of data per object defined in the objects Array.
// For every Column, it checks if data objects define a value for it.
// Arguments:
//   - sheet: the Sheet Object where the data will be written
//   - objects: an Array of Objects, each of which contains data for a row
//   - optHeadersRange: a Range of cells where the column headers are defined. This
//     defaults to the entire first row in sheet.
//   - optFirstDataRowIndex: index of the first row where data should be written. This
//     defaults to the row immediately below the headers.
function setRowsData(sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  var headers = normalizeHeaders(headersRange.getValues()[0]);

  var data = [];
  for(var i = 0; i < objects.length; ++i) {
    var values = [];
    for(var j = 0; j < headers.length; ++j) {
      var header = headers[j];
      // If the header is non-empty and the object value is 0...
      if((header.length > 0) && (objects[i][header] == 0)) {
        values.push(0);
      }
      // If the header is empty or the object value is empty...
      else if((!(header.length > 0)) || (objects[i][header] == '')) {
        values.push('');
      } else {
        values.push(objects[i][header]);
      }
    }
    data.push(values);
  }

  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(),
    objects.length, headers.length);
  destinationRange.setValues(data);
}


// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getLastColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for(var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for(var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if(isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if(hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for(var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if(key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for(var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if(letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if(!isAlnum(letter)) {
      continue;
    }
    if(key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if(upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof (cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}