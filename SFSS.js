/*global DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */

var sfss = sfss || {};
sfss.s = (function (r, smm) {
    'use strict';

    var sfss_interface, CACHE = {},
        PROPERTIES = {},

        fillInTemplateFromObject = smm.fillInTemplateFromObject,
        setRowsData = smm.setRowsData,
        getRowsData = smm.getRowsData,
        normalizeHeaders = smm.normalizeHeaders,
        normalizeHeader = smm.normalizeHeader,

        logDoc, scriptProperties = PropertiesService.getScriptProperties(),
        logId = scriptProperties.getProperty(normalizeHeader(r.LOG_FILE.s));


    if (logId) {
        logDoc = DocumentApp.openById(logId);
    }

    function setLogDoc(lD) {
        logDoc = lD;
    }
    // set TESTING to true when running unit tests
    function setTesting(b) {
        r.TESTING.b = b;
    }

    // used in testing template emails
    function getProperty(p) {
        return PROPERTIES[p];
    }

    // used in testing template emails
    function setProperty(p, v) {
        PROPERTIES[p] = v;
    }

    // used in testing template emails
    function deleteProperty(p) {
        delete PROPERTIES[p];
    }

    // build custom menu for spreadsheet
    function onOpen() {
        // spreadsheet menu
        try {
            var ss = SpreadsheetApp.getActiveSpreadsheet(),
                initialised = scriptProperties.getProperty("initialised");

            if (initialised) {
                ss.addMenu(r.FILM_SUBMISSION.s, r.MENU_ENTRIES.d);
            } else {
                ss.addMenu(r.FILM_SUBMISSION.s, [{
                    name: r.INITIALISE.s,
                    functionName: "setup"
                }]);
            }
        } catch (e) {
            Logger.log("Error in onOpen:" + e);
        }
    }

    //////////////////////////////////////
    // spreadsheet menue items
    // ///////////////////////////////////
    // mediaPresentNotConfirmed,
    // problem,
    // selected,
    // notSelected,
    // manual,
    // settingsOptions,
    // editAndSaveTemplates,
    // adHocEmail,
    // selectionNotification,
    // ///////////////////////////////////
    function statusDialog(status, confirmation, selection, height, title, html) {
        var value = scriptProperties.getProperty(normalizeHeader(status + ' ' + confirmation));
        if (value && value === 'true') {
            setStatus(status, confirmation, selection);
        } else {
            var ss = SpreadsheetApp.getActiveSpreadsheet(),
                app = UiApp.createApplication(),
                cPanel = app.createCaptionPanel(title).setId('cPanel'),
                vPanel = app.createVerticalPanel(),
                checkPanel = app.createHorizontalPanel(),
                handler = app.createServerHandler("buttonAction"),
                buttonGrid = app.createGrid(1, 3),
                ok = app.createButton(r.OK.s, handler).setId(r.OK.s),
                cancel = app.createButton(r.CANCEL.s, handler).setId(r.CANCEL.s),
                check = app.createCheckBox("Do not show me this diolog again.").setName('check'),
                label = app.createLabel(r.PROCESSING.s).setVisible(false);

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
            statusTextBox.setName(r.STATUS.s);
            statusTextBox.setVisible(false);
            statusTextBox.setValue(status);
            confirmationTextBox.setName(r.CONFIRMATION.s);
            confirmationTextBox.setVisible(false);
            confirmationTextBox.setValue(confirmation);
            selectionBox.setName(r.SELECTION.s);
            selectionBox.setVisible(false);
            selectionBox.setValue(selection);
            vPanel.add(statusTextBox);
            vPanel.add(confirmationTextBox);
            vPanel.add(selectionBox);
            handler.addCallbackElement(vPanel);

            var clientHandler = app.createClientHandler().forTargets([ok, cancel]).setEnabled(false).forTargets(label).setVisible(true);

            ok.addClickHandler(clientHandler);
            cancel.addClickHandler(clientHandler);

            ss.show(app);
        }
    }

    function mediaPresentNotConfirmed() {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        statusDialog(r.MEDIA_PRESENT.s, r.NOT_CONFIRMED.s, r.DO_NOT_CHANGE.s, 450, 'Set selection to Media Present, Not Confirmed', '<p>The state of a submission is given by the value of three fields, <b>' + r.STATUS.s + '</b>, <b>' + r.CONFIRMATION.s + '</b> and <b>' + r.SELECTION.s + '</b>.</p><p>Setting a submission to ' + r.STATUS.s + ': <b>' + r.MEDIA_PRESENT.s + '</b>, ' + r.CONFIRMATION.s + ': <b>' + r.NOT_CONFIRMED.s + '</b> (and leaving ' + r.SELECTION.s + ' unchanged) means the filmmaker will be automatically emailed after midnight to confirm the receipt of the Physical Media and the Permission Slip at the festival office, provided that:<ul><li>close of submission on ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + ' has not been reached at the time that the email confirmation is attempted</li><li>and that confirmations are currently enabled.</li></ul></p><p>Note: receipt confirmations are currently ' + getNamedValue(ss, r.ENABLE_CONFIRMATION.s) + '.</p><p>If you re-set the submission to a new different state before midnight, the receipt confirmation email will not be sent.</p><p>After the receipt confirmation email is sent the state of the submission will be set to ' + r.STATUS.s + ': <b>' + r.MEDIA_PRESENT.s + '</b>, ' + r.CONFIRMATION.s + ': <b>' + r.CONFIRMED.s + '</b>.</p><p>To continue with setting the submission state to ' + r.STATUS.s + ': <b>' + r.MEDIA_PRESENT.s + '</b>, ' + r.CONFIRMATION.s + ': <b>' + r.NOT_CONFIRMED.s + '</b> press OK, to not do this press CANCEL.</p><br/>');
    }

    function problem() {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        statusDialog(r.PROBLEM.s, r.DO_NOT_CHANGE.s, r.DO_NOT_CHANGE.s, 400, 'Set selection to Problem', '<p>The state of a submission is given by the value of three fields, <b>' + r.STATUS.s + '</b>, <b>' + r.CONFIRMATION.s + '</b> and <b>' + r.SELECTION.s + '</b>.</p><p>Setting a submission to ' + r.STATUS.s + ': <b>' + r.PROBLEM.s + '</b> (and leaving ' + r.CONFIRMATION.s + ' and ' + r.SELECTION.s + ' unchanged) means the filmmaker will be sent no reminder or receipt confirmation emails from now on. They will however, receive the final selection advisory email after close of submissions(' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + '), when it is sent.</p><p>If you set a submission to ' + r.STATUS.s + ': <b>' + r.PROBLEM.s + '</b>, please update the submission comment field with the details of the problem.</p><p>To continue with setting the submission state to ' + r.STATUS.s + ': <b>' + r.PROBLEM.s + '</b> press OK, to not do this press CANCEL.</p><br/>');
    }

    function selected() {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        statusDialog(r.DO_NOT_CHANGE.s, r.DO_NOT_CHANGE.s, r.SELECTED.s, 480, 'Set selection to Selected', '<p>The state of a submission is given by the value of three fields, <b>' + r.STATUS.s + '</b>, <b>' + r.CONFIRMATION.s + '</b> and <b>' + r.SELECTION.s + '</b>.</p><p>A submission set to ' + r.SELECTION.s + ': <b>' + r.SELECTED.s + '</b> (and leaving ' + r.STATUS.s + ' and ' + r.CONFIRMATION.s + ' unchanged) will still receive reminder and receipt confirmation emails dependent on its <b>' + r.STATUS.s + '</b> and <b>' + r.CONFIRMATION.s + '</b> provided that:<ul><li>it is at least ' + getNamedValue(ss, r.DAYS_BEFORE_REMINDER.s) + ' days before the close of submission on ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + ' at the time that the email is attempted</li><li>and that receipt confirmations are currently enabled for receipt confirmation,</li><li>and that reminders are currently enabled for reminders.</li></ul></p><p>Note: receipt confirmations are currently ' + getNamedValue(ss, r.ENABLE_CONFIRMATION.s) + '.</p><p>Note: reminders are currently ' + getNamedValue(ss, r.ENABLE_REMINDER.s) + '.</p><p>Submissions set to ' + r.SELECTION.s + ': <b>' + r.SELECTED.s + '</b> will receive the festival submission acceptance email after close of submissions(' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + '), when it is sent.</p><p>You will usually only set this state after the close of submissions.</p><p>To continue with setting the submission state to ' + r.SELECTION.v + ': <b>' + r.SELECTED.v + '</b> press OK, to not do this press CANCEL.</p><br/>');
    }

    function notSelected() {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        statusDialog(r.DO_NOT_CHANGE.s, r.DO_NOT_CHANGE.s, r.NOT_SELECTED.s, 480, 'Set selection to Not Selected', '<p>The state of a submission is given by the value of three fields, <b>' + r.STATUS.s + '</b>, <b>' + r.CONFIRMATION.s + '</b> and <b>' + r.SELECTION.s + '</b>.</p><p>A submission set to ' + r.SELECTION.s + ': <b>' + r.NOT_SELECTED.s + '</b> (and leaving ' + r.STATUS.s + ' and ' + r.CONFIRMATION.s + ' unchanged) will still receive reminder and receipt confirmation emails dependent on its <b>' + r.STATUS.s + '</b> and <b>' + r.CONFIRMATION.s + '</b> provided that:<ul><li>it is at least ' + getNamedValue(ss, r.DAYS_BEFORE_REMINDER.s) + ' days before the close of submission on ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + ' at the time that the email is attempted</li><li>and that receipt confirmations are currently enabled for receipt confirmation,</li><li>and that reminders are currently enabled for reminders.</li></ul></p><p>Note: receipt confirmations are currently ' + getNamedValue(ss, r.ENABLE_CONFIRMATION.s) + '.</p><p>Note: reminders are currently ' + getNamedValue(ss, r.ENABLE_REMINDER.s) + '.</p><p>Submissions set to ' + r.SELECTION.s + ': <b>' + r.NOT_SELECTED.s + '</b> will receive the festival submission not acceptanced email after close of submissions(' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + '), when it is sent.</p><p>You will usually only set this state after the close of submissions.</p><p>To continue with setting the submission state to ' + r.SELECTION.s + ': <b>' + r.NOT_SELECTED.s + '</b> press OK, to not do this press CANCEL.</p><br/>');
    }



    function manual() {
        var ss = SpreadsheetApp.getActiveSpreadsheet(),
            app = UiApp.createApplication().setTitle('Set state of selection');

        var grid = app.createGrid(3, 2);
        grid.setWidget(0, 0, app.createLabel(r.STATUS.s + ':'));
        var statusListBox = app.createListBox(false).setName(r.STATUS.s);
        statusListBox.addItem(r.DO_NOT_CHANGE.s);
        statusListBox.addItem(r.NO_MEDIA.s);
        statusListBox.addItem(r.MEDIA_PRESENT.s);
        statusListBox.addItem(r.PROBLEM.s);
        grid.setWidget(0, 1, statusListBox);
        grid.setWidget(1, 0, app.createLabel(r.CONFIRMATION.s + ':'));
        var confirmationListBox = app.createListBox(false).setName(r.CONFIRMATION.s);
        confirmationListBox.addItem(r.DO_NOT_CHANGE.s);
        confirmationListBox.addItem(r.NOT_CONFIRMED.s);
        confirmationListBox.addItem(r.CONFIRMED.s);
        grid.setWidget(1, 1, confirmationListBox);
        grid.setWidget(2, 0, app.createLabel(r.SELECTION.s + ':'));
        var selectionListBox = app.createListBox(false).setName(r.SELECTION.s);
        selectionListBox.addItem(r.DO_NOT_CHANGE.s);
        selectionListBox.addItem(r.NOT_SELECTED.s);
        selectionListBox.addItem(r.SELECTED.s);
        grid.setWidget(2, 1, selectionListBox);

        var cPanel = app.createCaptionPanel("Set State Of Range").setId('cPanel'),
            vPanel = app.createVerticalPanel();
        vPanel.add(app.createHTML('<p>The status of a submission is given by the value of three fields which are <b>' + r.STATUS.s + '</b>, <b>' + r.CONFIRMATION.s + '</b> and <b>' + r.SELECTION.s + '</b>.</p><p>For full details on what the various setting mean, please refer to documentation.</p>'));
        vPanel.add(grid);

        var handler = app.createServerHandler("buttonAction"),
            buttonGrid = app.createGrid(1, 3),
            ok = app.createButton(r.OK.s, handler).setId(r.OK.s),
            cancel = app.createButton(r.CANCEL.s, handler).setId(r.CANCEL.s),
            label = app.createLabel(r.PROCESSING.s).setVisible(false);
        handler.addCallbackElement(vPanel);
        buttonGrid.setWidget(0, 0, ok);
        buttonGrid.setWidget(0, 1, cancel);
        buttonGrid.setWidget(0, 2, label);

        var clientHandler = app.createClientHandler().forTargets([ok, cancel]).setEnabled(false).forTargets(label).setVisible(true);
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

        scriptProperties.setProperty(normalizeHeader(e.parameter[r.STATUS.s] + ' ' + e.parameter[r.CONFIRMATION.s]), e.parameter.check);

        if (e.parameter.source === r.OK.s) {
            setStatus(e.parameter[r.STATUS.s], e.parameter[r.CONFIRMATION.s], e.parameter[r.SELECTION.s]);
        } else {
            ss.toast("Canceling operation.", "WARNING", 5);
        }

        return app.close();
    }

    function setStatus(status, confirmation, selection) {
        try {
            log('setStatus:(status, confirmation, selection):(' + status + ',' + confirmation + ',' + selection + ')');
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            if (ss.getSheetName() === r.FILM_SUBMISSIONS_SHEET.s) {
                var activeRange = SpreadsheetApp.getActiveRange(),
                    filmSheet = ss.getSheetByName(r.FILM_SUBMISSIONS_SHEET.s),
                    dataRange = filmSheet.getDataRange(),
                    statusRange = findStatusRange(filmSheet, activeRange, dataRange);
                if (statusRange) {
                    setStatusByStatusRange(ss, filmSheet, statusRange, status, confirmation, selection);
                    //log status change
                    var headersRange = filmSheet.getRange(1, 1, 1, filmSheet.getDataRange().getLastColumn()),
                        headers = normalizeHeaders(headersRange.getValues()[0]),
                        idIndex = headers.indexOf(normalizeHeader(r.FILM_ID.s)) + 1,
                        ids = filmSheet.getRange(statusRange.getRow(), idIndex, statusRange.getHeight(), 1).getValues();
                    log('The following film submissions have been set to "' + status + '", "' + confirmation + '", "' + selection + '": ' + ids.concat());
                } else {
                    ss.toast("No films selected!", "WARNING", 5);
                }
            } else {
                ss.toast("Can only set Sate on '" + r.FILM_SUBMISSIONS_SHEET.s + "' sheet.", "WARNING", 5);
            }
        } catch (e) {
            log('There has been an error in setStatus:' + e);
        }
    }

    function setStatusByStatusRange(ss, filmSheet, statusRange, status, confirmation, selection) {
        var statusHeader = normalizeHeader(r.STATUS.s),
            confirmationHeader = normalizeHeader(r.CONFIRMATION.s),
            selectionHeader = normalizeHeader(r.SELECTION.s),
            rows = getRowsData(filmSheet, statusRange, 1);

        for (var i = 0; i < rows.length; i++) {
            if (status !== r.DO_NOT_CHANGE.s) {
                rows[i][statusHeader] = status;
            } else {
                status = rows[i][statusHeader];
            }
            if (confirmation !== r.DO_NOT_CHANGE.s) {
                rows[i][confirmationHeader] = confirmation;
            } else {
                confirmation = rows[i][confirmationHeader];
            }
            if (selection !== r.DO_NOT_CHANGE.s) {
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
        var color = r.NOT_SELECTED.s;

        if (selection === r.SELECTED.s) {
            color = r.SELECTED.s;
        } else if (status === r.PROBLEM.s) {
            color = r.PROBLEM.s;
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
            low = Math.max(firstRowA, firstRowB, 2),
            //cannot include sheet header
            high = Math.min(lastRowA, lastRowB);

        if (low <= high) {
            var minMaxColumn = findMinMaxColumns(filmSheet, [r.STATUS.s, r.CONFIRMATION.s, r.SELECTION.s]);
            range = filmSheet.getRange(low, minMaxColumn.min, high - low + 1, minMaxColumn.max - minMaxColumn.min + 1);
        }
        return range;
    }

    function pleaseWait(ss) {
        var app = UiApp.createApplication(),
            label = app.createLabel(r.PROCESSING.s).setStyleAttribute("fontSize", "200%");
        app.add(label);
        ss.show(app);
    }

    function editAndSaveTemplates(template) {
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
        try {
            var returnTo = template;
            template = template || r.SUBMISSION_CONFIRMATION.s;

            var ss = SpreadsheetApp.getActiveSpreadsheet();
            pleaseWait(ss);

            loadData(ss, r.TEST_DATA.s);
            loadData(ss, r.TEMPLATE_DATA.s);

            var app = UiApp.createApplication(),

                // Use these to turn off save, previous and next buttons if we have come from AD_HOC_MAIL or SELECTION_NOTIFICATION
                // As I turn them all off or all on together, logically I only need one hidden field for this information
                previousHidden = app.createHidden(normalizeHeader(r.PREVIOUS.s + ' ' + r.HIDDEN.s), !returnTo),
                nextHidden = app.createHidden(normalizeHeader(r.NEXT.s + ' ' + r.HIDDEN.s), !returnTo),
                saveHidden = app.createHidden(normalizeHeader(r.SAVE.s + ' ' + r.HIDDEN.s), !returnTo),

                vRoot = app.createVerticalPanel().setId(normalizeHeader('vRoot')),
                hTop = app.createHorizontalPanel(),
                handler = app.createServerHandler("templatesButtonAction"),
                saveButton = app.createButton(r.SAVE.s, handler).setId(normalizeHeader(r.SAVE.s)).setEnabled(!returnTo),
                cancelButton = app.createButton(r.CANCEL.s, handler).setId(normalizeHeader(r.CANCEL.s)),

                previousButtons = {},
                nextButtons = {},
                testButtons = {},

                previousButtonsArray = [],
                nextButtonsArray = [],
                testButtonsArray = [],

                processingLabel = app.createLabel(r.PROCESSING.s).setId(normalizeHeader(r.WAIT.s)).setVisible(false),

                returnToNotificationLabel = app.createLabel(r.SAVE_AND_RETURN_TO.s),
                returnToNotificationButton = app.createButton(r.SELECTION_NOTIFICATION.s, handler).setId(normalizeHeader(r.SELECTION_NOTIFICATION.s + ' ' + r.BUTTON.s)),
                returnToNotificationPLabel = app.createLabel(r.PROCESSING.s).setId(normalizeHeader(r.SELECTION_NOTIFICATION.s + ' ' + r.PLABEL.s)).setVisible(false),

                returnToAdHocLabel = app.createLabel(r.SAVE_AND_RETURN_TO.s),
                returnToAdHocButton = app.createButton(r.AD_HOC_EMAIL.s, handler).setId(normalizeHeader(r.AD_HOC_EMAIL.s + ' ' + r.BUTTON.s)),
                returnToAdHocPLabel = app.createLabel(r.PROCESSING.s).setId(normalizeHeader(r.AD_HOC_EMAIL.s + ' ' + r.PLABEL.s)).setVisible(false),

                templates = [r.SUBMISSION_CONFIRMATION.s, r.RECEIPT_CONFIMATION.s, r.REMINDER.s, r.NOT_ACCEPTED.s, r.ACCEPTED.s, r.AD_HOC_EMAIL.s],
                remainingDailyQuota = MailApp.getRemainingDailyQuota(),
                remainingEmailQuota = remainingDailyQuota - r.MIN_QUOTA.n,
                remainingEmailQuotaHidden = app.createHidden(normalizeHeader(r.REMAINING_EMAIL_QUOTA.s), remainingEmailQuota).setId(normalizeHeader(r.REMAINING_EMAIL_QUOTA.s)),

                testButtonGrid = app.createGrid(1, 3),
                //testButton          = app.createButton(TEST, handler).setId(normalizeHeader(TEST)).setEnabled(remainingEmailQuota>0),
                testButtonLabel = app.createLabel('Send test email:').setStyleAttribute('font-weight', 'bold'),
                testProcessingLabel = app.createLabel(r.PROCESSING.s).setStyleAttribute("fontSize", "50%").setVisible(false).setId(normalizeHeader(r.TEST_PROCESSING_LABEL.s)),

                testData = [{
                    namedValue: r.FIRST_NAME.s,
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
                    namedValue: r.LAST_NAME.s,
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
                    namedValue: r.TITLE.s,
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
                    namedValue: r.EMAIL.s,
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
                cAvaliableTages = app.createCaptionPanel(r.AVAILABLE_TAGS.s).setWidth('200').setHeight(height),
                testGrid = app.createGrid(2 * testData.length, 2),
                vTest = app.createVerticalPanel(),
                cTest = app.createCaptionPanel(r.SEND_TEST_EMAIL.s).setWidth('300').setHeight(height),

                statusHTML = app.createHTML().setHTML('<p>Remaining email quota for today:<b>' + remainingDailyQuota + '</b>. System will pause sending emails for the day when remaining email quota is less than or equal to: <b>' + r.MIN_QUOTA.n + '</b>.</p>').setId(normalizeHeader(r.STATUS_HTML.s)),

                statusWarningHTML = app.createHTML().setHTML('<p>NOTE: Todays remaining email quota is now less than or equal to <b>' + r.MIN_QUOTA.n + '</b>. Hence the system has now paused sending emails for the day.</p>').setStyleAttribute("color", "red").setVisible(remainingEmailQuota <= 0).setId(normalizeHeader(r.STATUS_WARNING_HTML.s)),

                avalibleTagsHTML = app.createHTML('<p>Name of film festival<br/><b>${"Festival Name"}</b></p><p>Address of festival website<br/><b>${"Festival Website"}</b></p><p>Link to release form<br/><b>${"Release Link"}</b></p><p>Date of close of submission<br/><b>${"Close Of Submission"}</b></p><p>Date of start of film festival<br/><b>${"Event Date"}</b></p><p>First name of filmmaker<br/><b>${"First Name"}</b></p><p>Last name of filmmaker<br/><b>${"Last Name"}</b></p><p>Title of submission<br/><b>${"Title"}</b></p>');

            // Every template needs its own previous, next and test button
            var templateData = {},
                templateName, previousTemplate, nextTemplate;
            for (var i = 0; i < templates.length; i++) {
                templateName = templates[i];
                previousButtons[templateName] = app.createButton(r.PREVIOUS.s).setVisible(templateName === template).setId(normalizeHeader(templateName + ' ' + r.PREVIOUS.s)).setEnabled(!returnTo);
                previousButtonsArray.push(previousButtons[templateName]);
                nextButtons[templateName] = app.createButton(r.NEXT.s).setVisible(templateName === template).setId(normalizeHeader(templateName + ' ' + r.NEXT.s)).setEnabled(!returnTo);
                nextButtonsArray.push(nextButtons[templateName]);
                testButtons[templateName] = app.createButton(r.TEST.s, handler).setVisible(templateName === template).setId(normalizeHeader(templateName + ' ' + r.TEST.s)).setEnabled(remainingEmailQuota > 0), testButtonsArray.push(testButtons[templateName]);
            }
            for (i = 0; i < templates.length; i++) {
                templateName = templates[i];

                var cTemplate = app.createCaptionPanel(templateName).setId(normalizeHeader(templateName)),
                    subjectLabel = app.createLabel(r.SUBJECT.s + ':').setStyleAttribute('font-weight', 'bold'),
                    subjectHelpLabel = app.createLabel(r.SUBJECT.s + ' required').setStyleAttribute("fontSize", "50%").setId(normalizeHeader(r.SUBJECT.s + ' ' + r.HELP.s)),
                    bodyLabel = app.createLabel(r.BODY.s + ':').setStyleAttribute('font-weight', 'bold'),
                    bodyHelpLabel = app.createLabel(r.BODY.s + ' required').setStyleAttribute("fontSize", "50%").setId(normalizeHeader(r.BODY.s + ' ' + r.HELP.s)),
                    subjectTextBox = app.createTextBox().setName(normalizeHeader(r.SUBJECT_LINE.s)).setText(getNamedValue(ss, templateName + ' ' + r.SUBJECT_LINE.s)).setWidth('300').setName(normalizeHeader(templateName + ' ' + r.SUBJECT_LINE.s)),
                    bodyTextArea = app.createTextArea().setName(normalizeHeader(r.BODY.s)).setText(getNamedValue(ss, templateName + ' ' + r.BODY.s)).setWidth('300').setHeight('280').setName(normalizeHeader(templateName + ' ' + r.BODY.s)),
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
                for (var j = 0; j < items.length; j++) {
                    item = items[j];

                    wrong = app.createClientHandler().validateNotLength(item.field, 1, null).forTargets(item.help).setStyleAttribute("color", "red").forTargets([saveButton].concat(previousButtonsArray, nextButtonsArray, testButtonsArray)).setEnabled(false);
                    item.field.addKeyUpHandler(wrong);

                    right = app.createClientHandler().validateLength(item.field, 1, null).forTargets(item.help).setStyleAttribute("color", "black");
                    item.field.addKeyUpHandler(right);
                }

                templateData[templateName] = {
                    panel: cTemplate,
                    subject: subjectTextBox,
                    body: bodyTextArea
                };

                hTop.add(templateData[templateName].panel.setVisible(templates[i] === template));
            }


            for (i = 0; i < templates.length; i++) {
                templateName = templates[i];
                previousTemplate = templates[((+i) - 1 + templates.length) % templates.length];
                nextTemplate = templates[((+i) + 1) % templates.length];
                previousButtons[templateName].addClickHandler(
                app.createClientHandler()
                // set all panels invisible except the previous panel
                .forTargets(templates.filter(isNotPreviousTemplate).map(getPanel)).setVisible(false).forTargets(templateData[previousTemplate].panel).setVisible(true)

                // the same as above but for buttons
                .forTargets(templates.filter(isNotPreviousTemplate).map(getPreviousButton)).setVisible(false).forTargets(templates.filter(isNotPreviousTemplate).map(getNextButton)).setVisible(false).forTargets(templates.filter(isNotPreviousTemplate).map(getTestButton)).setVisible(false).forTargets(previousButtons[previousTemplate]).setVisible(true).forTargets(nextButtons[previousTemplate]).setVisible(true).forTargets(testButtons[previousTemplate]).setVisible(true));

                nextButtons[templateName].addClickHandler(
                app.createClientHandler()
                // set all panels invisible except the previous panel
                .forTargets(templates.filter(isNotPreviousTemplate).map(getPanel)).setVisible(false).forTargets(templateData[nextTemplate].panel).setVisible(true)

                // the same as above but for buttons
                .forTargets(templates.filter(isNotPreviousTemplate).map(getPreviousButton)).setVisible(false).forTargets(templates.filter(isNotPreviousTemplate).map(getNextButton)).setVisible(false).forTargets(templates.filter(isNotPreviousTemplate).map(getTestButton)).setVisible(false).forTargets(previousButtons[nextTemplate]).setVisible(true).forTargets(nextButtons[nextTemplate]).setVisible(true).forTargets(testButtons[nextTemplate]).setVisible(true));
            }

            vAvaliableTages.add(avalibleTagsHTML);
            cAvaliableTages.add(vAvaliableTages);

            hTop.add(cAvaliableTages);

            //send test GUI
            var name, testItem = {},
                testLabel = {},
                itemData, testItemName = {};
            for (i = 0; i < testData.length; i++) {
                itemData = testData[i];
                testGrid.setWidget(2 * (+i), 0, app.createLabel(itemData.namedValue + ':').setStyleAttribute('font-weight', 'bold'));
                name = normalizeHeader(r.TEST.s + ' ' + itemData.namedValue);
                testItemName[i] = name;
                testItem[name] = app[itemData.type]().setName(name).setValue(getNamedValue(ss, normalizeHeader(name))).setWidth(itemData.width);
                testGrid.setWidget(2 * (+i), 1, testItem[name]);
                testLabel[name] = app.createLabel(itemData.label).setStyleAttribute("fontSize", "50%").setId(normalizeHeader(r.LABEL.s + ' ' + itemData.namedValue));
                testGrid.setWidget(2 * (+i) + 1, 1, testLabel[name]);
            }

            //client side validation
            var allGoodTest = app.createClientHandler().validateLength(testItem[normalizeHeader(r.TEST_FIRST_NAME.s)], 1, null).validateLength(testItem[normalizeHeader(r.TEST_LAST_NAME.s)], 1, null).validateLength(testItem[normalizeHeader(r.TEST_TITLE.s)], 1, null).validateEmail(testItem[normalizeHeader(r.TEST_EMAIL.s)]).validateRange(remainingEmailQuotaHidden, 1, null)

            .validateLength(templateData[r.SUBMISSION_CONFIRMATION.s].subject, 1, null).validateLength(templateData[r.SUBMISSION_CONFIRMATION.s].body, 1, null)

            .validateLength(templateData[r.RECEIPT_CONFIMATION.s].subject, 1, null).validateLength(templateData[r.RECEIPT_CONFIMATION.s].body, 1, null)

            .validateLength(templateData[r.REMINDER.s].subject, 1, null).validateLength(templateData[r.REMINDER.s].body, 1, null)

            .validateLength(templateData[r.NOT_ACCEPTED.s].subject, 1, null).validateLength(templateData[r.NOT_ACCEPTED.s].body, 1, null)

            .validateLength(templateData[r.ACCEPTED.s].subject, 1, null)

            .validateLength(templateData[r.ACCEPTED.s].body, 1, null)

            .validateLength(templateData[r.AD_HOC_EMAIL.s].subject, 1, null).validateLength(templateData[r.AD_HOC_EMAIL.s].body, 1, null)

            .forTargets([returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray)).setEnabled(true),

                allGoodSavePreviousNext = app.createClientHandler().validateLength(testItem[normalizeHeader(r.TEST_FIRST_NAME.s)], 1, null).validateLength(testItem[normalizeHeader(r.TEST_LAST_NAME.s)], 1, null).validateLength(testItem[normalizeHeader(r.TEST_TITLE.s)], 1, null).validateEmail(testItem[normalizeHeader(r.TEST_EMAIL.s)]).validateRange(remainingEmailQuotaHidden, 1, null)

                .validateLength(templateData[r.SUBMISSION_CONFIRMATION.s].subject, 1, null).validateLength(templateData[r.SUBMISSION_CONFIRMATION.s].body, 1, null)

                .validateLength(templateData[r.RECEIPT_CONFIMATION.s].subject, 1, null).validateLength(templateData[r.RECEIPT_CONFIMATION.s].body, 1, null)

                .validateLength(templateData[r.REMINDER.s].subject, 1, null).validateLength(templateData[r.REMINDER.s].body, 1, null)

                .validateLength(templateData[r.NOT_ACCEPTED.s].subject, 1, null).validateLength(templateData[r.NOT_ACCEPTED.s].body, 1, null)

                .validateLength(templateData[r.ACCEPTED.s].subject, 1, null).validateLength(templateData[r.ACCEPTED.s].body, 1, null)

                .validateLength(templateData[r.AD_HOC_EMAIL.s].subject, 1, null).validateLength(templateData[r.AD_HOC_EMAIL.s].body, 1, null)

                .validateOptions(previousHidden, ['true']).validateOptions(nextHidden, ['true']).validateOptions(saveHidden, ['true'])

                .forTargets([saveButton].concat(previousButtonsArray, nextButtonsArray)).setEnabled(true);

            var testItemRight, testItemWrong;
            for (i = 0; i < testData.length; i++) {
                itemData = testData[i];
                if (itemData.validate) {
                    name = testItemName[i];
                    testItemRight = null;
                    testItemWrong = null;
                    if (itemData.validate.validateLength) {
                        var min = itemData.validate.validateLength.min,
                            max = itemData.validate.validateLength.max;
                        testItemRight = app.createClientHandler().validateLength(testItem[name], min, max).forTargets(testLabel[name]).setStyleAttribute("color", "black"), testItemWrong = app.createClientHandler().validateNotLength(testItem[name], min, max).forTargets(testLabel[name]).setStyleAttribute("color", "red").forTargets([saveButton, returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray, previousButtonsArray, nextButtonsArray)).setEnabled(false);
                    } else if (itemData.validate.validateEmail) {
                        testItemRight = app.createClientHandler().validateEmail(testItem[name]).forTargets(testLabel[name]).setStyleAttribute("color", "black"), testItemWrong = app.createClientHandler().validateNotEmail(testItem[name]).forTargets(testLabel[name]).setStyleAttribute("color", "red").forTargets([saveButton, returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray, previousButtonsArray, nextButtonsArray)).setEnabled(false);
                    }
                    if (testItemRight) {
                        testItem[name].addKeyUpHandler(testItemRight);
                    }
                    if (testItemWrong) {
                        testItem[name].addKeyUpHandler(testItemWrong);
                    }
                    testItem[name].addKeyUpHandler(allGoodTest);
                    testItem[name].addKeyUpHandler(allGoodSavePreviousNext);
                }
            }

            for (i = 0; i < templates.length; i++) {
                templateName = templates[i];
                templateData[templateName].subject.addKeyUpHandler(allGoodTest);
                templateData[templateName].subject.addKeyUpHandler(allGoodSavePreviousNext);
                templateData[templateName].body.addKeyUpHandler(allGoodTest);
                templateData[templateName].body.addKeyUpHandler(allGoodSavePreviousNext);
            }

            // end of client side validation

            // Add test buttons to grid
            var off = app.createClientHandler().forTargets([saveButton, returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray, previousButtonsArray, nextButtonsArray)).setEnabled(false).forTargets(testProcessingLabel).setVisible(true);
            vTest.add(testGrid);
            vTest.add(statusHTML);
            testButtonGrid.setWidget(0, 0, testButtonLabel);
            var hTestButtonsPanel = app.createHorizontalPanel();
            for (i = 0; i < testButtonsArray.length; i++) {
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
                clickSCPN = app.createClientHandler().forTargets([saveButton, cancelButton, returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray, previousButtonsArray, nextButtonsArray)).setEnabled(false).forTargets(processingLabel).setVisible(true);
            saveButton.addClickHandler(clickSCPN);
            cancelButton.addClickHandler(clickSCPN);

            var clickAdHocNot = app.createClientHandler().forTargets([saveButton, cancelButton, returnToNotificationButton, returnToAdHocButton].concat(testButtonsArray, previousButtonsArray, nextButtonsArray)).setEnabled(false).forTargets(returnToNotificationPLabel).setVisible(true).forTargets(returnToAdHocPLabel).setVisible(true);
            returnToNotificationButton.addClickHandler(clickAdHocNot);
            returnToAdHocButton.addClickHandler(clickAdHocNot);

            buttonGrid.setWidget(0, 0, saveButton);
            buttonGrid.setWidget(0, 1, cancelButton);
            var hPreviousButtons = app.createHorizontalPanel();
            for (i = 0; i < previousButtonsArray.length; i++) {
                hPreviousButtons.add(previousButtonsArray[i]);
            }
            buttonGrid.setWidget(0, 2, hPreviousButtons);
            var hNextButtons = app.createHorizontalPanel();
            for (i = 0; i < nextButtonsArray.length; i++) {
                hNextButtons.add(nextButtonsArray[i]);
            }
            buttonGrid.setWidget(0, 3, hNextButtons);
            buttonGrid.setWidget(0, 4, processingLabel);
            vControls.add(buttonGrid);

            if (returnTo) {
                if (returnTo === r.ACCEPTED.s || returnTo === r.NOT_ACCEPTED.s) {
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
            var tags = [r.TEST_TITLE.s, r.TEST_FIRST_NAME.s, r.TEST_LAST_NAME.s, r.TEST_EMAIL.s];
            vRoot.add(app.createHidden(normalizeHeader(r.TAGS.s), tags.join(',')));
            vRoot.add(app.createHidden(normalizeHeader(r.TEMPLATE_DATA_NAMES.s), templates.join(',')));
            vRoot.add(app.createHidden(normalizeHeader(r.TEST_DATA_NAMES.s), testData.map(function (item) {
                return r.TEST.s + ' ' + item.namedValue;
            }).join(',')));

            handler.addCallbackElement(vRoot);

            app.add(vRoot);

            app.setWidth('900');
            app.setHeight(appHeight);

            ss.show(app);

        } catch (e) {
            log('updateTemplates:' + catchToString(e));
        }
    }



    function templatesButtonAction(e) {
        function saveTestAndTemplateData() {
            var range = ss.getRangeByName(normalizeHeader(r.TEST_DATA.s)),
                values = testDataNames.map(function (itemName) {
                    return [itemName, e.parameter[normalizeHeader(itemName)]];
                });
            range.setValues(values);

            var templatesAndSuffix = templates.map(function (item) {
                return [item + ' ' + r.SUBJECT_LINE.s, item + ' ' + r.BODY.s];
            });
            templatesAndSuffix = Array.concat.apply([], templatesAndSuffix);
            values = [
                [r.TEMPLATE_NAME.s, r.TEMPLATE.s]
            ].concat(templatesAndSuffix.map(function (itemName) {
                return [itemName, e.parameter[normalizeHeader(itemName)]];
            })); // Add column titles in.
            range = ss.getRangeByName(normalizeHeader(r.TEMPLATE_DATA.s));
            range.setValues(values);

            ss.toast("Data has been saved.", "INFORMATION", 5);
        }
        try {
            var app = UiApp.getActiveApplication(),
                ss = SpreadsheetApp.getActiveSpreadsheet(),
                testDataNames = e.parameter[normalizeHeader(r.TEST_DATA_NAMES.s)].split(','),
                templates = e.parameter[normalizeHeader(r.TEMPLATE_DATA_NAMES.s)].split(','),
                testButtonNames = templates.map(function (x) {
                    return normalizeHeader(x + ' ' + r.TEST.s);
                });

            if (e.parameter.source === normalizeHeader(r.SELECTION_NOTIFICATION.s + ' ' + r.BUTTON.s)) {
                saveTestAndTemplateData();
                selectionNotification();
            } else if (e.parameter.source === normalizeHeader(r.AD_HOC_EMAIL.s + ' ' + r.BUTTON.s)) {
                saveTestAndTemplateData();
                adHocEmail();
            } else if (testButtonNames.indexOf(e.parameter.source) > -1) {
                var template = templates[testButtonNames.indexOf(e.parameter.source)],
                    remainingEmailQuotaTextBox = app.getElementById(normalizeHeader(r.REMAINING_EMAIL_QUOTA.s)),
                    remainingDailyQuota = MailApp.getRemainingDailyQuota();

                var tags = e.parameter[normalizeHeader(r.TAGS.s)].split(','),
                    mergeData = {},
                    tag;

                for (var i = 0; i < tags.length; i++) {
                    tag = tags[i];
                    mergeData[normalizeHeader(tag.replace(/^Test /, ''))] = e.parameter[normalizeHeader(tag)];
                }

                var statusHTML = app.getElementById(normalizeHeader(r.STATUS_HTML.s)),
                    statusWarningHTML = app.getElementById(normalizeHeader(r.STATUS_WARNING_HTML.s)),
                    testProcessingLabel = app.getElementById(normalizeHeader(r.TEST_PROCESSING_LABEL.s)),
                    saveButton = app.getElementById(normalizeHeader(r.SAVE.s)),
                    cancelButton = app.getElementById(normalizeHeader(r.CANCEL.s)),
                    returnToNotificationButton = app.getElementById(normalizeHeader(r.SELECTION_NOTIFICATION.s + ' ' + r.BUTTON.s)),
                    returnToAdHocButton = app.getElementById(normalizeHeader(r.AD_HOC_EMAIL.s + ' ' + r.BUTTON.s));

                mergeData[normalizeHeader(r.EMAIL.s)] = e.parameter[normalizeHeader(r.TEST_EMAIL.s)];
                loadData(ss, r.FESTIVAL_DATA.s);

                // load test template into cache
                setNamedValue(ss, template + ' ' + r.TEST.s + ' ' + r.SUBJECT_LINE.s, e.parameter[normalizeHeader(template + ' ' + r.SUBJECT_LINE.s)]);
                setNamedValue(ss, template + ' ' + r.TEST.s + ' ' + r.BODY.s, e.parameter[normalizeHeader(template + ' ' + r.BODY.s)]);
                template += (' ' + r.TEST.s);

                if (remainingDailyQuota > r.MIN_QUOTA.n) { // check just in case
                    remainingDailyQuota = mergeAndSend(ss, mergeData, template, remainingDailyQuota);
                    ss.toast("Test email has been sent.", "INFORMATION", 5);
                } else {
                    ss.toast("Not enough email quota remained, email not sent!", "WARNING!", 5);
                }

                remainingEmailQuotaTextBox.setValue((remainingDailyQuota - r.MIN_QUOTA.n).toString());
                statusHTML.setHTML('<p>Remaining email quota for today:<b>' + remainingDailyQuota + '</b>. System will pause sending emails for the day when remaining email quota is less than or equal to: <b>' + r.MIN_QUOTA.n + '</b>.</p>');
                if (remainingDailyQuota <= r.MIN_QUOTA.n) {
                    statusWarningHTML.setVisible(true);
                }
                saveButton.setEnabled(e.parameter[normalizeHeader(r.SAVE.s + ' ' + r.HIDDEN.s)] === 'true');
                cancelButton.setEnabled(true);
                for (i = 0; i < templates.length; i++) {
                    app.getElementById(normalizeHeader(templates[i] + ' ' + r.PREVIOUS.s)).setEnabled(e.parameter[normalizeHeader(r.PREVIOUS.s + ' ' + r.HIDDEN.s)] === 'true');
                    app.getElementById(normalizeHeader(templates[i] + ' ' + r.NEXT.s)).setEnabled(e.parameter[normalizeHeader(r.NEXT.s + ' ' + r.HIDDEN.s)] === 'true');
                    app.getElementById(normalizeHeader(templates[i] + ' ' + r.TEST.s)).setEnabled(remainingDailyQuota > r.MIN_QUOTA.n);
                }
                returnToNotificationButton.setEnabled(true);
                returnToAdHocButton.setEnabled(true);
                testProcessingLabel.setVisible(false);
            } else if (e.parameter.source === normalizeHeader(r.SAVE.s)) {
                saveTestAndTemplateData();
                app.close();
            } else {
                ss.toast("Canceling operation.", "WARNING", 5);
                app.close();
            }
        } catch (error) {
            log('templatesButtonAction:error::' + catchToString(error));
        }
        return app;
    }

    function settingsOptions() {
        try {
            var ss = SpreadsheetApp.getActiveSpreadsheet();

            pleaseWait(ss);

            var festivaDataNames = loadData(ss, r.FESTIVAL_DATA.s),
                app = UiApp.createApplication(),
                festivaDataNamesHidden = app.createHidden(normalizeHeader(r.FESTIVAL_DATA_NAMES.s), festivaDataNames.join(',')),
                gridData = [{
                    namedValue: r.FESTIVAL_NAME.s,
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
                    namedValue: r.FESTIVAL_WEBSITE.s,
                    type: 'createTextBox',
                    width: '300px',
                    help: 'Not required.'
                }, {
                    namedValue: r.CLOSE_OF_SUBMISSION.s,
                    type: 'createDateBox',
                    width: '75px',
                    help: 'Required (default ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + ').',
                    validate: {
                        valadateDateDependance: {}
                    }
                }, {
                    namedValue: r.EVENT_DATE.s,
                    type: 'createDateBox',
                    width: '75px',
                    help: 'Required (default ' + prettyPrintDate(getNamedValue(ss, r.EVENT_DATE.s)) + ').',
                    validate: {
                        valadateDateDependance: {}
                    }
                }, {
                    namedValue: r.RELEASE_LINK.s,
                    type: 'createTextBox',
                    width: '300px',
                    help: 'Not required.'
                }, {
                    namedValue: r.DAYS_BEFORE_REMINDER.s,
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
                    namedValue: r.ENABLE_REMINDER.s,
                    type: 'createListBox',
                    width: '100px',
                    help: 'Required (default ' + getNamedValue(ss, r.ENABLE_REMINDER.s) + ').',
                    list: [r.NOT_ENABLED.s, r.ENABLED.s]
                }, {
                    namedValue: r.ENABLE_CONFIRMATION.s,
                    type: 'createListBox',
                    width: '100px',
                    help: 'Required (default ' + getNamedValue(ss, r.ENABLE_REMINDER.s) + ').',
                    list: [r.NOT_ENABLED.s, r.ENABLED.s]
                }, {
                    namedValue: r.FIRST_FILM_ID.s,
                    type: 'createHidden'
                }],
                handler = app.createServerHandler("settingsOptionsButtonAction"),
                save = app.createButton(r.SAVE.s, handler).setId(normalizeHeader(r.SAVE.s)),
                cancel = app.createButton(r.CANCEL.s, handler).setId(normalizeHeader(r.CANCEL.s)),
                cPanel = app.createCaptionPanel("Settings and Options").setId('cPanel'),
                vPanel = app.createVerticalPanel(),
                buttonGrid = app.createGrid(1, 3);

            var grid = app.createGrid(2 * gridData.length, 2),
                item = {},
                label = {},
                back = {},
                itemName, itemValue, correct, wrong, i;

            // build GUI
            for (i = 0; i < gridData.length; i++) {
                itemName = normalizeHeader(gridData[i].namedValue);
                itemValue = getNamedValue(ss, gridData[i].namedValue);
                if (gridData[i].type === 'createHidden') {
                    vPanel.add(app.createHidden(itemName, itemValue));
                    continue;
                }
                grid.setWidget(2 * (+i), 0, app.createLabel(gridData[i].namedValue + ': '));
                if (gridData[i].type === 'createListBox') {
                    item[itemName] = app.createListBox();
                    for (var j = 0; j < gridData[i].list.length; j++) {
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

                label[itemName] = app.createLabel('').setStyleAttribute("fontSize", "50%").setId(itemName + r.HELP.s);
                grid.setWidget(2 * (+i) + 1, 1, label[itemName]);

                if (gridData[i].help) {
                    label[itemName].setText(gridData[i].help);
                }
            }

            // add validators to GUI
            for (i in gridData) {
                if (gridData[i].validate) {
                    itemName = normalizeHeader(gridData[i].namedValue);
                    if (gridData[i].validate.valadateDateDependance) {
                        item[itemName].addValueChangeHandler(app.createServerHandler("validateDates").addCallbackElement(vPanel));
                    }
                    if (gridData[i].validate.validateLength) {
                        wrong = app.createClientHandler().validateNotLength(item[itemName], gridData[i].validate.validateLength.min, gridData[i].validate.validateLength.max).forTargets(label[itemName]).setText(gridData[i].error).setStyleAttribute("color", "red").forTargets(save).setEnabled(false);
                        item[itemName].addKeyUpHandler(wrong);
                        correct = app.createClientHandler().validateLength(item[itemName], gridData[i].validate.validateLength.min, gridData[i].validate.validateLength.max).forTargets(label[itemName]).setText(gridData[i].help).setStyleAttribute("color", "black");
                        item[itemName].addKeyUpHandler(correct);
                    }
                    if (gridData[i].validate.validateRange) {
                        wrong = app.createClientHandler().validateNotRange(item[itemName], gridData[i].validate.validateRange.min, gridData[i].validate.validateRange.max).forTargets(label[itemName]).setText(gridData[i].error).setStyleAttribute("color", "red").forTargets(save).setEnabled(false);
                        item[itemName].addKeyUpHandler(wrong);
                        correct = app.createClientHandler().validateRange(item[itemName], gridData[i].validate.validateRange.min, gridData[i].validate.validateRange.max).forTargets(label[itemName]).setText(gridData[i].help).setStyleAttribute("color", "black");
                        item[itemName].addKeyUpHandler(correct);
                    }
                    if (gridData[i].validate.validateInteger) {
                        wrong = app.createClientHandler().validateNotInteger(item[itemName]).forTargets(label[itemName]).setText(gridData[i].error).setStyleAttribute("color", "red").forTargets(save).setEnabled(false);
                        item[itemName].addKeyUpHandler(wrong);
                        correct = app.createClientHandler().validateInteger(item[itemName]).forTargets(label[itemName]).setText(gridData[i].help).setStyleAttribute("color", "black");
                        item[itemName].addKeyUpHandler(correct);
                    }
                }
            }
            var dateDiff = app.createHidden().setValue(diffDays(getNamedValue(ss, r.EVENT_DATE.s), getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s))).setId(normalizeHeader(r.DATE_DIFF.s));

            var allCorrect = app.createClientHandler().validateLength(item[normalizeHeader(r.FESTIVAL_NAME.s)], 1, null).validateRange(dateDiff, 1, null).validateRange(item[normalizeHeader(r.DAYS_BEFORE_REMINDER.s)], 1, null).forTargets(save).setEnabled(true);
            item[normalizeHeader(r.FESTIVAL_NAME.s)].addKeyUpHandler(allCorrect);
            item[normalizeHeader(r.DAYS_BEFORE_REMINDER.s)].addKeyUpHandler(allCorrect);
            //dateDiff.addValueChangeHandler(allCorrect); this does not work :(
            vPanel.add(app.createHTML('<p>Please enter the settings and options for your film festival below.</p>'));
            vPanel.add(grid);
            vPanel.add(dateDiff);

            label = app.createLabel(r.PROCESSING.s).setId(normalizeHeader(r.WAIT.s)).setVisible(false);
            handler.addCallbackElement(vPanel);
            buttonGrid.setWidget(0, 0, save);
            buttonGrid.setWidget(0, 1, cancel);
            buttonGrid.setWidget(0, 2, label);

            var clientHandler = app.createClientHandler().forTargets([save, cancel]).setEnabled(false).forTargets(label).setVisible(true);
            save.addClickHandler(clientHandler);
            cancel.addClickHandler(clientHandler);

            vPanel.add(buttonGrid);
            vPanel.add(festivaDataNamesHidden);

            handler.addCallbackElement(vPanel);

            cPanel.add(vPanel);
            app.add(cPanel);

            app.setHeight('560');

            ss.show(app);
        } catch (e) {
            log('settingsOptions:error:' + catchToString(e));
        }
    }

    function validateDates(e) {
        var app = UiApp.getActiveApplication(),
            ss = SpreadsheetApp.getActiveSpreadsheet(),
            eventDate = e.parameter[normalizeHeader(r.EVENT_DATE.s)],
            cosDate = e.parameter[normalizeHeader(r.CLOSE_OF_SUBMISSION.s)],
            eventLabel = app.getElementById(normalizeHeader(r.EVENT_DATE.s + ' ' + r.HELP.s)),
            cosLabel = app.getElementById(normalizeHeader(r.CLOSE_OF_SUBMISSION.s + ' ' + r.HELP.s)),
            save = app.getElementById(normalizeHeader(r.SAVE.s)),
            dateDiff = app.getElementById(normalizeHeader(r.DATE_DIFF.s));

        if (eventDate instanceof Date && cosDate instanceof Date) {
            // case: both dates
            if (eventDate < cosDate) {
                // error condition
                cosLabel.setText('"' + r.CLOSE_OF_SUBMISSION.s + '" must be before "' + r.EVENT_DATE.s + '".').setStyleAttribute("color", "red");
                eventLabel.setText('"' + r.CLOSE_OF_SUBMISSION.s + '" must be before "' + r.EVENT_DATE.s + '".').setStyleAttribute("color", "red");
                save.setEnabled(false);
            } else {
                cosLabel.setText('Required (default ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + ').').setStyleAttribute("color", "black");
                eventLabel.setText('Required (default ' + prettyPrintDate(getNamedValue(ss, r.EVENT_DATE.s)) + ').').setStyleAttribute("color", "black");
                var festivalName = e.parameter[normalizeHeader(r.FESTIVAL_NAME.s)],
                    daysBeforeReminder = e.parameter[normalizeHeader(r.DAYS_BEFORE_REMINDER.s)];

                if (festivalName && festivalName.length > 0 && daysBeforeReminder && parseInt(daysBeforeReminder) > 0) {
                    save.setEnabled(true); //only enable save if everything is OK
                }
            }
            dateDiff.setValue(diffDays(eventDate, cosDate));
        } else {
            //case: at least one is not a date
            if (eventDate instanceof Date) {
                eventLabel.setText('Required (default ' + prettyPrintDate(getNamedValue(ss, r.EVENT_DATE.s)) + ').').setStyleAttribute("color", "black");
            } else {
                eventLabel.setText('Not a valid date').setStyleAttribute("color", "red");
            }

            if (cosDate instanceof Date) {
                cosLabel.setText('Required (default ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + ').').setStyleAttribute("color", "black");
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
            eventDate = e.parameter[normalizeHeader(r.EVENT_DATE.s)],
            cosDate = e.parameter[normalizeHeader(r.CLOSE_OF_SUBMISSION.s)],
            eventLabel = app.getElementById(normalizeHeader(r.EVENT_DATE.s + ' ' + r.HELP.s)),
            cosLabel = app.getElementById(normalizeHeader(r.CLOSE_OF_SUBMISSION.s + ' ' + r.HELP.s)),
            save = app.getElementById(normalizeHeader(r.SAVE.s)),
            cancel = app.getElementById(normalizeHeader(r.CANCEL.s)),
            wait = app.getElementById(normalizeHeader(r.WAIT.s));

        if (e.parameter.source === normalizeHeader(r.SAVE.s)) {
            // need to check that dates are not wrong before saving
            if (!(eventDate instanceof Date)) {
                eventLabel.setText('Not a valid date').setStyleAttribute("color", "red");
                save.setEnabled(false);
                cancel.setEnabled(true);
                wait.setVisible(false);
                return app; // do not close the dialog
            }
            if (!(cosDate instanceof Date)) {
                cosLabel.setText('Not a valid date').setStyleAttribute("color", "red");
                save.setEnabled(false);
                cancel.setEnabled(true);
                wait.setVisible(false);
                return app; // do not close the dialog
            }
            //if we got here, the dates should indeed be dates!
            if (eventDate < cosDate) {
                // error condition
                cosLabel.setText('"' + r.CLOSE_OF_SUBMISSION.s + '" must be before "' + r.EVENT_DATE.s + '".').setStyleAttribute("color", "red");
                eventLabel.setText('"' + r.CLOSE_OF_SUBMISSION.s + '" must be before "' + r.EVENT_DATE.s + '".').setStyleAttribute("color", "red");
                save.setEnabled(false);
                cancel.setEnabled(true);
                wait.setVisible(false);
                return app; // do not close the dialog
            }


            var festivaDataNames = e.parameter[normalizeHeader(r.FESTIVAL_DATA_NAMES.s)].split(','),
                range = ss.getRangeByName(normalizeHeader(r.FESTIVAL_DATA.s)),
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
        editAndSaveTemplates(r.ACCEPTED.s);
    }

    function selectionNotificationNotAcceptedTemplate() {
        editAndSaveTemplates(r.NOT_ACCEPTED.s);
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
                enable = app.createButton(r.ENABLE.s, handler).setId(r.ENABLE.s),
                unenable = app.createButton(r.UNENABLE.s, handler).setId(r.UNENABLE.s),
                cancel = app.createButton(r.CANCEL.s, handler).setId(r.CANCEL.s),
                waitHTML = app.createHTML(r.PROCESSING.s).setVisible(false),
                cannotHTML = app.createHTML('<p>NOTE: You have not reached <b>' + r.CLOSE_OF_SUBMISSION.s + '</b> on ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + ' yet. Hence you cannot yet enable <b>' + r.SELECTION_NOTIFICATION.s + '</b>.</p><br/>').setStyleAttribute("color", "red").setVisible(false),
                height = '680',
                currentDate = new Date(),
                closeOfSubmission = getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s);


            vPanel.add(app.createHTML('<p><b>' + r.SELECTION_NOTIFICATION.s + '</b> is the process of mailing each submitting filmmaker with the selection status of their submission.</p><p>It is intended that you only enable <b>' + r.SELECTION_NOTIFICATION.s + '</b> once per film festival, after close ofsubmission.</p><p>Before you enable <b>' + r.SELECTION_NOTIFICATION.s + '</b>, each film submission that you are selecting for your film festival, must have its <b>' + r.SELECTION.s + '</b> field set to <b>' + r.SELECTED.s + '</b>.</p><p>When you enable <b>' + r.SELECTION_NOTIFICATION.s + '</b>, the system will start to process the submissions after midnight of that day, in their submission order.</p><p>If you unenable <b>' + r.SELECTION_NOTIFICATION.s + '</b> before midnight, the notification emails will not be sent and the <b>' + r.SELECTION_NOTIFICATION.s + '</b> will have been canceled.</p><p>Each submission that the system processes (that does not have <b>' + r.STATUS.s + '</b> set to <b>' + r.PROBLEM.s + '</b>) will receive the <b>' + r.ACCEPTED.s + '</b> email if it has <b>' + r.SELECTION.s + '</b> set to <b>' + r.SELECTED.s + '</b>, otherwise it will receive the <b>' + r.NOT_ACCEPTED.s + '</b> email.</p><p>A submission with <b>' + r.STATUS.s + '</b> set to <b>' + r.PROBLEM.s + '</b> will not receive a <b>' + r.SELECTION_NOTIFICATION.s + '</b> email.</p>'));

            // link to edit and test ACCEPTED template
            var grid = app.createGrid(2, 3),
                acceptedHandle = app.createServerHandler('selectionNotificationAcceptedTemplate'),
                acceptedButton = app.createButton(r.ACCEPTED.s, acceptedHandle).setId(normalizeHeader(r.ACCEPTED.s)),
                acceptedPLabel = app.createLabel(r.PROCESSING.s).setId(normalizeHeader(r.ACCEPTED.s + ' ' + r.PLABEL.s)).setStyleAttribute("fontSize", "50%").setVisible(false);
            grid.setWidget(0, 0, app.createHTML('Edit and test <b>' + r.ACCEPTED.s + '</b> template:'));
            grid.setWidget(0, 1, acceptedButton);
            grid.setWidget(0, 2, acceptedPLabel);

            // link to edit and test NOT_ACCEPTED template
            var notAcceptedHandle = app.createServerHandler('selectionNotificationNotAcceptedTemplate'),
                notAcceptedButton = app.createButton(r.NOT_ACCEPTED.s, notAcceptedHandle).setId(normalizeHeader(r.NOT_ACCEPTED.s)),
                notAcceptedPLabel = app.createLabel(r.PROCESSING.s).setId(normalizeHeader(r.NOT_ACCEPTED.s + ' ' + r.PLABEL.s)).setStyleAttribute("fontSize", "50%").setVisible(false);
            grid.setWidget(1, 0, app.createHTML('Edit and test <b>' + r.NOT_ACCEPTED.s + '</b> template:'));
            grid.setWidget(1, 1, notAcceptedButton);
            grid.setWidget(1, 2, notAcceptedPLabel);

            vPanel.add(grid);

            var acceptedProcessing = app.createClientHandler().forTargets([enable, unenable, acceptedButton, notAcceptedButton]).setEnabled(false).forTargets(acceptedPLabel).setVisible(true);
            acceptedButton.addClickHandler(acceptedProcessing);

            var notAcceptedProcessing = app.createClientHandler().forTargets([enable, unenable, acceptedButton, notAcceptedButton]).setEnabled(false).forTargets(notAcceptedPLabel).setVisible(true);
            notAcceptedButton.addClickHandler(notAcceptedProcessing);

            vPanel.add(app.createHTML('<p>If the system gets within ' + r.MIN_QUOTA.n + ' of using up the daily email quota, the <b>' + r.SELECTION_NOTIFICATION.s + '</b> will be paused till the next day.</p><p>NOTE: your daily email quote is currently at ' + MailApp.getRemainingDailyQuota() + '.</p><p>You cannot enable <b>' + r.SELECTION_NOTIFICATION.s + '</b> before <b>' + r.CLOSE_OF_SUBMISSION.s + '</b> on ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + '.</p><p><p>To continue with enabling <b>' + r.SELECTION_NOTIFICATION.s + '</b> press <b>' + r.ENABLE.s + '</b>, to not do this press <b>' + r.CANCEL.s + '</b>.</p><br/>'));

            if (currentDate < closeOfSubmission) {
                cannotHTML.setVisible(true);
                enable.setEnabled(false);
            }

            vPanel.add(cannotHTML);
            buttonGrid.setWidget(0, 0, getNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s) === r.NOT_STARTED.s ? enable : unenable);
            buttonGrid.setWidget(0, 1, cancel);
            buttonGrid.setWidget(0, 2, waitHTML);
            vPanel.add(buttonGrid);
            cPanel.add(vPanel);
            app.add(cPanel);

            app.setHeight(height);

            handler.addCallbackElement(vPanel);

            var clientHandler = app.createClientHandler().forTargets([enable, unenable, cancel, acceptedButton, notAcceptedButton]).setEnabled(false).forTargets(waitHTML).setVisible(true);

            unenable.addClickHandler(clientHandler);
            enable.addClickHandler(clientHandler);
            cancel.addClickHandler(clientHandler);

            ss.show(app);
        } catch (e) {
            log('selectionNotification:error:' + catchToString(e));
        }
    }

    function selectionNotificationButtonAction(e) {
        var app = UiApp.getActiveApplication(),
            ss = SpreadsheetApp.getActiveSpreadsheet();

        if (e.parameter.source === r.ENABLE.s) {
            log(r.SELECTION_NOTIFICATION.s + ' set to:' + r.PENDING.s);
            setNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s, r.PENDING.s);
            ss.toast("Selection Notification has been queued.", "INFORMATION", 5);
        } else if (e.parameter.source === r.UNENABLE.s) {
            log(r.SELECTION_NOTIFICATION.s + ' set to:' + r.NOT_STARTED.s);
            setNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s, r.NOT_STARTED.s);
            ss.toast("Selection Notification has been unqueued.", "INFORMATION", 5);
        } else {
            ss.toast("Canceling operation.", "WARNING", 5);
        }
        return app.close();
    }

    function adHocEmailTemplate() {
        selectionNotification(r.AD_HOC_EMAIL.s);
    }

    function adHocEmail() {
        try {
            var ss = SpreadsheetApp.getActiveSpreadsheet();

            pleaseWait(ss);

            var app = UiApp.createApplication(),
                cPanel = app.createCaptionPanel("Queue " + r.AD_HOC_EMAIL.s).setId('cPanel'),
                vPanel = app.createVerticalPanel(),
                handler = app.createServerHandler("adHocEmailButtonAction"),
                enable = app.createButton(r.ENABLE.s, handler).setId(normalizeHeader(r.ENABLE.s)),
                unenable = app.createButton(r.UNENABLE.s, handler).setId(normalizeHeader(r.UNENABLE.s)),
                cancel = app.createButton(r.CANCEL.s, handler).setId(normalizeHeader(r.CANCEL.s)),
                adHocEmailButton = app.createButton(r.AD_HOC_EMAIL.s, handler).setId(normalizeHeader(r.AD_HOC_EMAIL.s)),
                adHocEmailPLabel = app.createLabel(r.PROCESSING.s).setVisible(false).setId(normalizeHeader(r.AD_HOC_EMAIL.s + ' ' + r.PLABEL.s)),
                waitLabel = app.createLabel(r.PROCESSING.s).setVisible(false),
                buttonGrid = app.createGrid(1, 3),
                adHocEamilGrid = app.createGrid(1, 3);

            vPanel.add(app.createHTML('<p>The system can send an <b>' + r.AD_HOC_EMAIL.s + '</b> to all submitting filmmakers.</p><p>This can be used to inform the filmmakers of changes in the circumstances of the festival, such as a change in the close of submission date.</p>'));

            adHocEamilGrid.setWidget(0, 0, app.createLabel('Edit and test ' + r.AD_HOC_EMAIL.s + ':'));


            adHocEamilGrid.setWidget(0, 1, adHocEmailButton);
            adHocEamilGrid.setWidget(0, 2, adHocEmailPLabel);
            vPanel.add(adHocEamilGrid);

            vPanel.add(app.createHTML('<p>When you enable <b>' + r.AD_HOC_EMAIL.s + '</b>, the system will start to process the submissions after midnight of that day, in their submission order.</p><p>If you unenable <b>' + r.AD_HOC_EMAIL.s + '</b> before midnight, the emails will not be sent and <b>' + r.AD_HOC_EMAIL.s + '</b> will have been canceled.</p><p>NOTE: submissions with <b>' + r.STATUS.s + '</b> set to <b>' + r.PROBLEM.s + '</b> will not receive an email.</p><br/>'));

            buttonGrid.setWidget(0, 0, getNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s) === r.NOT_STARTED.s ? enable : unenable);
            buttonGrid.setWidget(0, 1, cancel);
            buttonGrid.setWidget(0, 2, waitLabel);
            vPanel.add(buttonGrid);

            cPanel.add(vPanel);
            app.add(cPanel);

            handler.addCallbackElement(vPanel);

            var click = app.createClientHandler().forTargets([enable, unenable, cancel, adHocEmailButton]).setEnabled(false).forTargets(waitLabel).setVisible(true);
            unenable.addClickHandler(click);
            enable.addClickHandler(click);
            cancel.addClickHandler(click);

            var adHocClick = app.createClientHandler().forTargets([enable, unenable, cancel, adHocEmailButton]).setEnabled(false).forTargets(adHocEmailPLabel).setVisible(true);
            adHocEmailButton.addClickHandler(adHocClick);

            ss.show(app);
        } catch (e) {
            log('adHocEmail:error:' + catchToString(e));
        }
    }

    function adHocEmailButtonAction(e) {
        var app = UiApp.getActiveApplication(),
            ss = SpreadsheetApp.getActiveSpreadsheet();
        log('adHocEmailButtonAction:e.parameter.source:' + e.parameter.source);
        if (e.parameter.source === normalizeHeader(r.UNENABLE.s)) {
            setNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s, r.NOT_STARTED.s);
            ss.toast("Ad Hoc Email has been unqueued.", "INFORMATION", 5);
            app.close();
        } else if (e.parameter.source === normalizeHeader(r.ENABLE.s)) {
            setNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s, r.PENDING.s);
            ss.toast("Ad Hoc Email has been queued.", "INFORMATION", 5);
            app.close();
        } else if (e.parameter.source === normalizeHeader(r.AD_HOC_EMAIL.s)) {
            adHocEmailTemplate();
        } else {
            app.close();
            ss.toast("Canceling operation.", "WARNING", 5);
        }
        return app;
    }

    function setup() {
        try {
            if (r.TESTING.b || !scriptProperties.getProperty("initialised")) {
                if (!r.TESTING.b) {
                    scriptProperties.setProperty("initialised", "initialised");
                }

                log("setup start.");
                var itemData, sheet;

                // build online submission form
                var form = FormApp.create(r.FORM_TITLE.s).setConfirmationMessage(r.FORM_RESPONSE.s);
                for (var itemIndex = 0; itemIndex < r.FORM_DATA.d.length; itemIndex++) {
                    itemData = r.FORM_DATA.d[itemIndex];
                    if (itemData === "section") {
                        form.addSectionHeaderItem();
                    } else {
                        var formItem = form[itemData.type]().setTitle(itemData.title).setHelpText(itemData.help).setRequired(itemData.required ? true : false);
                        if (itemData.type === "addCheckboxItem") {
                            formItem.setChoices([formItem.createChoice(itemData.choice)]);
                        }
                        if (itemData.type === "addMultipleChoiceItem") {
                            formItem.setChoiceValues(itemData.choices);
                        }
                    }
                }

                var ss = SpreadsheetApp.getActiveSpreadsheet(),
                    ssId = ss.getId(),
                    formId = form.getId();

                // find parent folder of spreadsheet and move form to that folder
                var ssParentFolder = DriveApp.getFileById(ssId).getParents().next(),
                    formFile = DriveApp.getFileById(formId),
                    formFolder = DriveApp.createFolder(r.FORM_FOLDER.s);

                DriveApp.removeFolder(formFolder); // remove from root
                DriveApp.removeFile(formFile); // remove from root
                ssParentFolder.addFolder(formFolder);
                formFolder.addFile(formFile);

                log('Built online submission form.');

                // initialise spreadsheet and set spreadsheet as the destination of the form
                form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

                // Need this for testing.
                // Otherwise 'Form Responses 1' not present when looked for!!!
                // Why does this only break testing!!!
                SpreadsheetApp.flush();

                // create the r.FILM_SUBMISSIONS_SHEET.s, r.OPTIONS_SETTINGS_SHEET.s and r.TEMPLATE_SHEET.s sheet
                ss.getSheetByName('Form Responses 1').setName(r.FILM_SUBMISSIONS_SHEET.s);
                ss.insertSheet(r.OPTIONS_SETTINGS_SHEET.s, 2);
                ss.insertSheet(r.TEMPLATE_SHEET.s, 3);
                ss.deleteSheet(ss.getSheetByName('Sheet1'));

                // add extra columns to r.FILM_SUBMISSIONS_SHEET.s
                sheet = ss.getSheetByName(r.FILM_SUBMISSIONS_SHEET.s);
                for (var i = 0; i < r.ADDITIONAL_COLUMNS.d.length; i++) {
                    columnIndex = r.ADDITIONAL_COLUMNS.d[i].columnIndex;
                    if (columnIndex === 'end') {
                        columnIndex = sheet.getLastColumn();
                    }
                    sheet.insertColumnAfter(columnIndex);
                    sheet.getRange(1, columnIndex + 1, 1, 1).setValue(r.ADDITIONAL_COLUMNS.d[i].title);
                }

                //add templates to r.TEMPLATE_SHEET.s
                saveData(ss, ss.getSheetByName(r.TEMPLATE_SHEET.s).getRange(1, 1, 1, 1), r.TEMPLATE_DATA.d, r.TEMPLATE_DATA.s, 'mistyrose');

                //initialise r.OPTIONS_SETTINGS_SHEET.s
                for (i = 0; i < r.OPTION_SHEET_DATA.d.length; i++) {
                    var dataItem = r.OPTION_SHEET_DATA.d[i];
                    saveData(ss, ss.getSheetByName(r.OPTIONS_SETTINGS_SHEET.s).getRange(1, 2 * (+i) + 1, 1, 1), dataItem.data, dataItem.name, (+i) % 2 ? 'aliceblue' : 'mistyrose');

                }

                // update menu
                if (!r.TESTING.b) {
                    ss.updateMenu(r.FILM_SUBMISSION.s, r.MENU_ENTRIES.d);
                }

                setColumnWidth(ss, r.FILM_SUBMISSIONS_SHEET.s, r.filmSheetColumnWidth.d);
                setColumnWidth(ss, r.TEMPLATE_SHEET.s, r.templateSheetColumnWidth.d);

                log('Built spreadsheet.');

                if (!r.TESTING.b) {
                    settingsOptions(); // let user set options and settings
                }

                // Only need the triggers if we are not testing.
                if (!r.TESTING.b) {
                    // enable film submission processing on form submission
                    ScriptApp.newTrigger("hProcessSubmission").forSpreadsheet(ss).onFormSubmit().create();

                    // enable reminders and confirmations, sent out after midnight each day
                    ScriptApp.newTrigger("hReminderConfirmation").timeBased().everyDays(1).atHour(3).create();

                    log('Added triggers.');
                    log("setup end.");
                }
            }
        } catch (e) {
            log('setup:error:' + catchToString(e));
        }
    }

    function setColumnWidth(ss, sheetName, columnWidths) {
        try {
            var sheet = ss.getSheetByName(sheetName),
                headers = sheet.getRange(1, 1, 1, sheet.getDataRange().getLastColumn()).getValues()[0],
                width = {};
            for (var columnIndex = 0; columnIndex < columnWidths.length; columnIndex++) {
                width[columnWidths[columnIndex].header] = columnWidths[columnIndex].width;
            }

            for (columnIndex in headers) {
                if (width[headers[columnIndex]]) {
                    sheet.setColumnWidth(1 + (+columnIndex), width[headers[columnIndex]]);
                }
            }
            SpreadsheetApp.flush();
        } catch (e) {
            log('setColumnWidth:error:' + catchToString(e));
        }
    }

    function log(info) {
        try {
            if (!logDoc) {
                logDoc = DocumentApp.create(r.LOG_FILE.s), logId = logDoc.getId();
                var logFile = DriveApp.getFileById(logId);
                DriveApp.removeFile(logFile); // remove from root
                var ss = SpreadsheetApp.getActiveSpreadsheet(),
                    ssParentFolder = DriveApp.getFileById(ss.getId()).getParents().next(),
                    logFolder = DriveApp.createFolder(r.LOG_FOLDER.s);
                Logger.log('ssParentFolder: ' + ssParentFolder);
                Logger.log('logFolder: ' + logFolder);
                DriveApp.removeFolder(logFolder); // remove from root
                ssParentFolder.addFolder(logFolder);
                logFolder.addFile(logFile);

                logDoc.getBody().appendParagraph((new Date()) + ":Created log file!");

                scriptProperties.setProperty(normalizeHeader(r.LOG_FILE.s), logId);

            }
            logDoc.getBody().appendParagraph((new Date()) + ": " + info);
        } catch (e) {
            log('log:error:' + catchToString(e));
        }
    }

    function catchToString(err) {
        var errInfo = "Caught something:\n";
        for (var prop in err) {
            if (err.hasOwnProperty(prop)) {
                errInfo += "(" + prop + ", " + err[prop] + ")\n";
            }
        }
        errInfo += "  toString(): " + " value: [" + err.toString() + "]";
        return errInfo;
    }

    function mergeAndSend(ss, submission, templateName, emailQuotaRemaining) {
        try {
            emailQuotaRemaining = emailQuotaRemaining ? emailQuotaRemaining : 0; // Set to 0 if not supplied i.e. it dosnt matter in this case
            submission[normalizeHeader(r.FESTIVAL_NAME.s)] = getNamedValue(ss, r.FESTIVAL_NAME.s);
            submission[normalizeHeader(r.CLOSE_OF_SUBMISSION.s)] = prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s));
            submission[normalizeHeader(r.EVENT_DATE.s)] = prettyPrintDate(getNamedValue(ss, r.EVENT_DATE.s));
            submission[normalizeHeader(r.FESTIVAL_WEBSITE.s)] = getNamedValue(ss, r.FESTIVAL_WEBSITE.s);
            submission[normalizeHeader(r.RELEASE_LINK.s)] = getNamedValue(ss, r.RELEASE_LINK.s);

            var subjectTemplate = getNamedValue(ss, templateName + ' ' + r.SUBJECT_LINE.s),
                bodyTemplate = getNamedValue(ss, templateName + ' ' + r.BODY.s);

            if (submission[normalizeHeader(r.EMAIL.s)]) {
                MailApp.sendEmail(submission[normalizeHeader(r.EMAIL.s)], fillInTemplateFromObject(subjectTemplate, submission), fillInTemplateFromObject(bodyTemplate, submission));
                emailQuotaRemaining -= 1;
                if (r.TESTING.b) {
                    var templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
                    if (templatesTesting) {
                        templatesTesting += ',' + templateName;
                    } else {
                        templatesTesting = templateName;
                    }
                    setProperty(normalizeHeader(r.TEMPLATES_TESTING.s), templatesTesting);
                }
                log('mergeAndSend:email sent');
            } else {
                log('mergeAndSend:error:no email found!');
            }


        } catch (e) {
            log('mergeAndSend:error:' + catchToString(e));
        }
        return emailQuotaRemaining;
    }

    // This trigger function processes form submissions. Due to the fact that
    // On Form Submit event occationally does not fire
    function hProcessSubmission(event) {
        log('hProcessSubmission start');
        var lock = LockService.getPublicLock();
        lock.waitLock(30000);
        try {
            var ss = SpreadsheetApp.getActiveSpreadsheet();

            // load data in advance
            loadData(ss, r.FESTIVAL_DATA.s);
            loadData(ss, r.COLOR_DATA.s);
            loadData(ss, r.TEMPLATE_DATA.s);
            loadData(ss, r.INTERNALS.s);

            var filmSheet = ss.getSheetByName(r.FILM_SUBMISSIONS_SHEET.s),
                eventRow = event.range.getRow(),
                emailQuotaRemaining = MailApp.getRemainingDailyQuota(),
                headersOfInterest, minMaxColumns, lastContact, rowIndex;

            // Insert LAST_CONTACT,FILM_ID and STATUS into film submission. Also make sure LENGTH has correct form
            // Use smallest range that will work with one call to getRowsData and setRowsData
            headersOfInterest = [r.TIMESTAMP.s, r.LAST_CONTACT.s, r.FILM_ID.s, r.STATUS.s, r.CONFIRMATION.s, r.SELECTION.s, r.LENGTH.s];
            minMaxColumns = findMinMaxColumns(filmSheet, headersOfInterest);


            var lastContactIndex = minMaxColumns.indices[normalizeHeader(r.LAST_CONTACT.s)];
            var lastContactCell = filmSheet.getRange(eventRow, lastContactIndex, 1, 1);
            lastContact = lastContactCell.getValue();

            if (lastContact instanceof Date || lastContact === r.NO_CONTACT.s) {
                throw {
                    message: 'ERROR: this submission has already been processed!'
                };
            }

            // check that previous submissions have been processed
            var lastUnprocessedIndex = eventRow;
            if (eventRow > 2) { //NOTE: row 2 should contain the first submission
                // We know that we have least one previous row.
                // Count backward until we encounter a processed row.
                log('There are more than 2 submissions counting this one.');
                for (rowIndex = eventRow - 1; rowIndex > 1; rowIndex -= 1) {
                    log('rowIndex:' + rowIndex);
                    lastContact = filmSheet.getRange(rowIndex, lastContactIndex, 1, 1).getValue();
                    log('lastContact:' + lastContact);
                    if (lastContact instanceof Date || lastContact === r.NO_CONTACT.s) {
                        break;
                    } else {
                        lastUnprocessedIndex = rowIndex;
                    }
                }
            }
            log('(eventRow,lastUnprocessedIndex):(' + eventRow + ',' + lastUnprocessedIndex + ')');

            // assert 1 < lastUnprocessedIndex <= eventRow
            maxColumn = minMaxColumns.max;
            minColumn = minMaxColumns.min;
            unprocessedRange = filmSheet.getRange(lastUnprocessedIndex, minColumn, eventRow - lastUnprocessedIndex + 1, maxColumn - minColumn + 1);

            var rows = getRowsData(filmSheet, unprocessedRange, 1);

            var index, currentFilmID = -1,
                filmSubmissionData, length, confirmation, color, color_CONFIRMED, color_NOT_CONFIRMED, date = event.namedValues[r.TIMESTAMP.s][0];

            // Update Find and update Film ID for current submission
            if (eventRow === 2) {
                // first submission
                currentFilmID = getNamedValue(ss, r.FIRST_FILM_ID.s) - 1;
            } else {
                currentFilmID = getNamedValue(ss, r.CURRENT_FILM_ID.s);

            }

            for (index = 0; index < rows.length; index += 1) {
                currentFilmID = currentFilmID + 1;
                log('processSubmission:processing film submission ' + r.ID.s + pad(currentFilmID));
                filmSubmissionData = rows[index];

                length = normaliseAndValidateDuration(filmSubmissionData[normalizeHeader(r.LENGTH.s)]);
                if (length) {
                    filmSubmissionData[normalizeHeader(r.LENGTH.s)] = length;
                }
                filmSubmissionData[normalizeHeader(r.STATUS.s)] = r.NO_MEDIA.s;
                filmSubmissionData[normalizeHeader(r.FILM_ID.s)] = r.ID.s + pad(currentFilmID);
                if (r.MIN_QUOTA.n < emailQuotaRemaining) {
                    lastContact = date;
                    confirmation = r.CONFIRMED.s;
                } else {
                    lastContact = r.NO_CONTACT.s;
                    confirmation = r.NOT_CONFIRMED.s;
                }
                emailQuotaRemaining -= 1;
                filmSubmissionData[normalizeHeader(r.CONFIRMATION.s)] = confirmation;
                filmSubmissionData[normalizeHeader(r.SELECTION.s)] = r.NOT_SELECTED.s;
                filmSubmissionData[normalizeHeader(r.LAST_CONTACT.s)] = lastContact;
                filmSubmissionData[normalizeHeader(r.SCORE.s)] = -1;
            }

            setNamedValue(ss, r.CURRENT_FILM_ID.s, currentFilmID);

            setRowsData(filmSheet, rows, filmSheet.getRange(1, minColumn, 1, maxColumn - minColumn + 1), lastUnprocessedIndex);

            //set status color on film submission
            color_CONFIRMED = findStatusColor(ss, r.NO_MEDIA.s, r.CONFIRMED.s, r.NOT_SELECTED.s);
            color_NOT_CONFIRMED = findStatusColor(ss, r.NO_MEDIA.s, r.NOT_CONFIRMED.s, r.NOT_SELECTED.s);


            // send confirmation email
            for (index = 0; index < rows.length; index += 1) {
                filmSubmissionData = rows[index];
                if (filmSubmissionData[normalizeHeader(r.LAST_CONTACT.s)] !== r.NO_CONTACT.s) {
                    mergeAndSend(ss, filmSubmissionData, r.SUBMISSION_CONFIRMATION.s);
                    color = color_CONFIRMED;
                } else {
                    color = color_NOT_CONFIRMED;
                }
                filmSheet.getRange(lastUnprocessedIndex + index, 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);
            }
        } catch (e) {
            log('hProcessSubmission:error:' + catchToString(e));
        }
        lock.releaseLock();
        log('hProcessSubmission end');
    }

    function prettyPrintDate(date) {
        if (!date) {
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
        log('normaliseAndValidateDuration:start:length:"' + length + '"');

        function babylonian(myNumber) {
            return 0 <= myNumber && myNumber <= 59;
        }
        var returnLength, mymatch;
        try {
            var hours = '0';
            minutes = '0', seconds = '0', myregexp = /^[^\d]*([\d]+)[^\d]*$|^[^\d]*([\d]+)[^\d]+([\d]+)[^\d]*$|^[^\d]*([\d]+)[^\d]+([\d]+)[^\d]+([\d]+)[^\d]*$/i;

            mymatch = myregexp.exec(length);
            if (mymatch) {
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
                    hours = mymatch[4];
                    minutes = mymatch[5];
                    seconds = mymatch[6];
                }

                if (babylonian(hours) && babylonian(minutes) && babylonian(seconds)) {
                    returnLength = ('0' + hours).slice(-2) + ':' + ('0' + minutes).slice(-2) + ':' + ('0' + seconds).slice(-2);
                }
            }
        } catch (e) {
            log('normaliseAndValidateDuration:error:' + catchToString(e));
        }
        log('normaliseAndValidateDuration:end');
        return returnLength ? returnLength : length;
    }

    function pad(number) {
        if (number > 0) {
            return ("000000" + number).slice(-r.PAD_NUMBER.n);
        } else {
            return ("\\d\\d\\d\\d\\d\\d").slice(-2 * r.PAD_NUMBER.n);
        }
    }

    // for testing
    function setPadNumber(n) {
        r.PAD_NUMBER.n = n;
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

            if (!value) {
                value = ss.getRangeByName(name).getValue();
                setNamedValueInCache(name, value);
            }
        } catch (e) {
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

            // getRangeByName should not fail by throwing an exception when a named
            // range does not exsist, but it does!!!
            try {
                range = ss.getRangeByName(name);
            } catch (e) {
                log('setNamedValue:probably not important:error:' + catchToString(e));
            }

            if (range) {
                range.setValue(value);
            } else {
                log('setNamedValue: no range with name:' + name + ' found!');
            }
            setNamedValueInCache(name, value);
        } catch (e) {
            log('setNamedValue:name:' + name + ':error:' + catchToString(e));
        }
    }

    // Load data into CACHE so that getNamedRange will later get it from CACHE and
    // not have to make an individual getValue on a spreadsheet cell
    function loadData(ss, name) {
        var items = [];
        try {
            log('loadData:loading data block:' + name);
            name = normalizeHeader(name);
            var loaded = getNamedValueFromCache(name);
            if (!loaded) {
                var item;

                // Named range assumed to have width of 2.
                // Each row is a name value pair.
                var values = ss.getRangeByName(name).getValues();
                for (var i = 0; i < values.length; i++) {
                    item = values[i];
                    items.push(item[0]);
                    CACHE[normalizeHeader(item[0])] = item[1];
                }
                setNamedValueInCache(name, true);
            } else {
                log('loadData:date block has already been loaded!');
            }
        } catch (e) {
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

            for (rowIndex = 0; rowIndex < values.length; rowIndex++) {
                rowValue = values[rowIndex];
                for (columnIndex = 0; columnIndex < rowValue.length; columnIndex++) {
                    value = rowValue[columnIndex];
                    currentCell = startCell.offset(rowIndex, columnIndex);
                    if (Array.isArray(value)) { //expect array to have 2 elements
                        ss.setNamedRange(normalizeHeader(value[0]), currentCell);
                        value = value[1];
                    }
                    currentCell.setValue(value);
                }
            }
        } catch (e) {
            log('saveData:error:' + catchToString(e));
        }
    }

    ////////////////////////////////////////////////////////
    // We expect hReminderConfirmation to do the following:
    // if(Ad Hoc Email is pending or in progress){
    //    Continue with Ad Hoc Email.
    //    Stop when all submissions processed
    //    or daily mail allocation is down to r.MIN_QUOTA.n.
    // } else if (not yet reached Close Of Submissions) {
    //    While daily mail allocation is greater than r.MIN_QUOTA.n for each submission
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
    //    or daily mail allocation is down to r.MIN_QUOTA.n.
    // }
    ///////////////////////////////////////////////////////////////////////////////
    function hReminderConfirmation() {
        try {
            log('hReminderConfirmation start');
            var ss = SpreadsheetApp.getActiveSpreadsheet();

            loadData(ss, r.FESTIVAL_DATA.s);
            loadData(ss, r.TEMPLATE_DATA.s);
            loadData(ss, r.INTERNALS.s);
            loadData(ss, r.COLOR_DATA.s);

            var emailQuotaRemaining = MailApp.getRemainingDailyQuota(),
                enabledReminder = getNamedValue(ss, r.ENABLE_REMINDER.s) === r.ENABLED.s,
                enableConfirmation = getNamedValue(ss, r.ENABLE_CONFIRMATION.s) === r.ENABLED.s,
                filmSheet = ss.getSheetByName(r.FILM_SUBMISSIONS_SHEET.s),
                lastRow = filmSheet.getLastRow(),
                currentDate = new Date(),
                closeOfSubmission = getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s),
                daysBeforeReminder = getNamedValue(ss, r.DAYS_BEFORE_REMINDER.s),
                daysTillClose = diffDays(closeOfSubmission, currentDate);

            log('hReminderConfirmation:enabledReminder: ' + enabledReminder + '\nenableConfirmation: ' + enableConfirmation + '\ncloseOfSubmission: ' + closeOfSubmission + '\ndaysBeforeReminder: ' + daysBeforeReminder + '\ndaysTillClose: ' + daysTillClose + '\nemailQuotaRemaining: ' + emailQuotaRemaining);

            if (lastRow > 1) {
                log('hReminderConfirmation:processing');
                var headersRange = filmSheet.getRange(1, 1, 1, filmSheet.getDataRange().getLastColumn()),
                    headers = normalizeHeaders(headersRange.getValues()[0]),
                    confirmationColumnIndex = headers.indexOf(normalizeHeader(r.CONFIRMATION.s)) + 1,
                    lastContactColumnIndex = headers.indexOf(normalizeHeader(r.LAST_CONTACT.s)) + 1,
                    minMaxColumns = findMinMaxColumns(filmSheet, [r.LAST_CONTACT.s, r.STATUS.s, r.CONFIRMATION.s, r.SELECTION.s, r.FIRST_NAME.s, r.EMAIL.s, r.TITLE.s]),
                    height = filmSheet.getDataRange().getLastRow() - 1,
                    range = filmSheet.getRange(2, minMaxColumns.min, height, minMaxColumns.max - minMaxColumns.min + 1),
                    submissions = getRowsData(filmSheet, range, 1),
                    currentAdHocEmail = getNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s),
                    currentSelectionNotification = getNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s),
                    color = "",
                    i, submission, status, confirmation, selection, filmId, filmIdFormat, processNextSubmission, setToFinishedAtEnd, matches;

                if (currentAdHocEmail !== r.NOT_STARTED.s) {
                    // AD HOC EMAIL
                    log('hReminderConfirmation:processing Ad Hoc Mail');
                    filmIdFormat = r.FILM_ID.s + pad(-1);
                    processNextSubmission = currentAdHocEmail === r.PENDING.s;
                    setToFinishedAtEnd = true;
                    log('hReminderConfirmation:currentAdHocEmail:' + currentAdHocEmail);
                    // Loop through all submissions until we come to the first submission after r.CURRENT_AD_HOC_EMAIL.s.
                    // Start with first if currentAdHocEmail === PENDING. Normally this will be the case.
                    for (i = 0; i < submissions.length; i++) {
                        if (emailQuotaRemaining <= r.MIN_QUOTA.n) {
                            log('hReminderConfirmation:ran out of email quota during ' + r.AD_HOC_EMAIL.s + ', (r.MIN_QUOTA.n, emailQuotaRemaining):(' + r.MIN_QUOTA.n + ',' + emailQuotaRemaining + '). Terminating ' + r.AD_HOC_EMAIL.s + ' early.');
                            setToFinishedAtEnd = false;
                            break; //stop sending emails for the day
                        }

                        submission = submissions[i];
                        status = submission[normalizeHeader(r.STATUS.s)];
                        filmId = submission[normalizeHeader(r.FILM_ID.s)];

                        if (!filmId) {
                            log('hReminderConfirmation:ERROR: exepected film id!');
                            continue;
                        }

                        if (status === r.PROBLEM.s) {
                            // Ingnore submissions with status PROBLEM
                            continue;
                        }

                        matches = filmId.match(filmIdFormat, 'i');
                        if (matches && matches.length > 0) {
                            filmId = matches[0];
                            if (processNextSubmission) {
                                // send Ad Hoc Email
                                emailQuotaRemaining = mergeAndSend(ss, submission, r.AD_HOC_EMAIL.s, emailQuotaRemaining);
                                setNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s, filmId); //Update current
                                log('hReminderConfirmation:Ad Hoc Mail:processed:' + filmId);
                            } else {
                                processNextSubmission = filmId === currentAdHocEmail;
                                log('hReminderConfirmation:found currentAdHocEmail');
                            }
                        } else {
                            log('hReminderConfirmation:ERROR: exepected film id!');
                        }
                    } // for (var i in submissions)
                    if (setToFinishedAtEnd) {
                        setNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s, r.NOT_STARTED.s); // Job done. Reset.
                    }

                    // if (currentAdHocEmail!==NOT_STARTED)
                } else if (currentDate < closeOfSubmission) {
                    for (i = 0; i < submissions.length; i++) {
                        if (emailQuotaRemaining <= r.MIN_QUOTA.n) {
                            log('Ran out of email quota, (r.MIN_QUOTA.n, emailQuotaRemaining):(' + r.MIN_QUOTA.n + ',' + emailQuotaRemaining + ').');
                            log('Terminating daily reminder and confirmation processing early.');
                            break; //stop sending emails for the day
                        }

                        submission = submissions[i];
                        status = submission[normalizeHeader(r.STATUS.s)];
                        confirmation = submission[normalizeHeader(r.CONFIRMATION.s)];
                        selection = submission[normalizeHeader(r.SELECTION.s)];
                        if (status === r.PROBLEM.s) {
                            // Ingnore submissions with status PROBLEM
                            continue;
                        }

                        if (status === r.NO_MEDIA.s) {
                            if (confirmation === r.NOT_CONFIRMED.s) {
                                //Send submission confirmation. This should be a very rare case!
                                log('hReminderConfirmation:no submission confirmation has been sent to ' + submission[normalizeHeader(r.FILM_ID.s)] + '.');

                                emailQuotaRemaining = mergeAndSend(ss, submission, r.SUBMISSION_CONFIRMATION.s, emailQuotaRemaining);

                                filmSheet.getRange(2 + (+i), lastContactColumnIndex, 1, 1).setValue(currentDate);
                                filmSheet.getRange(2 + (+i), confirmationColumnIndex, 1, 1).setValue(r.CONFIRMED.s);
                                color = findStatusColor(ss, status, r.CONFIRMED.s, selection);
                                filmSheet.getRange(2 + (+i), 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);
                                log('hReminderConfirmation:submission confirmation has now been sent to ' + submission[normalizeHeader(r.FILM_ID.s)] + ' and its ' + r.LAST_CONTACT.s + ' updated.');
                            } else if (daysTillClose > daysBeforeReminder && enabledReminder) {
                                // Need to send reminder to send Media and release form
                                var lastContact = submission[normalizeHeader(r.LAST_CONTACT.s)];

                                var daysSinceLastContact = diffDays(currentDate, lastContact);

                                if (daysSinceLastContact >= daysBeforeReminder) {
                                    //send reminder
                                    log('It is ' + daysSinceLastContact + ' days since ' + r.LAST_CONTACT.s + ' with submission ' + submission[normalizeHeader(r.FILM_ID.s)] + ' which has status ' + r.NO_MEDIA.s + '.');

                                    emailQuotaRemaining = mergeAndSend(ss, submission, r.REMINDER.s, emailQuotaRemaining);

                                    filmSheet.getRange(2 + (+i), lastContactColumnIndex, 1, 1).setValue(currentDate);
                                    log('hReminderConfirmation:a reminder has been sent to submission ' + submission[normalizeHeader(r.FILM_ID.s)] + ' and its ' + r.LAST_CONTACT.s + ' updated.');
                                }
                            }

                        } else if (status === r.MEDIA_PRESENT.s && confirmation === r.NOT_CONFIRMED.s && enableConfirmation) {
                            //send confirmation of receipt of media
                            log('hReminderConfirmation:submission ' + submission[normalizeHeader(r.FILM_ID.s)] + ' has status ' + r.MEDIA_PRESENT.s + ', ' + r.NOT_CONFIRMED.s + '.');

                            emailQuotaRemaining = mergeAndSend(ss, submission, r.RECEIPT_CONFIMATION.s, emailQuotaRemaining);

                            filmSheet.getRange(2 + (+i), lastContactColumnIndex, 1, 1).setValue(currentDate);
                            filmSheet.getRange(2 + (+i), confirmationColumnIndex, 1, 1).setValue(r.CONFIRMED.s);
                            color = findStatusColor(ss, status, r.CONFIRMED.s, selection);
                            filmSheet.getRange(2 + (+i), 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);
                            log('hReminderConfirmation:the status of submission ' + submission[normalizeHeader(r.FILM_ID.s)] + ' has been updated to ' + r.MEDIA_PRESENT.s + ', ' + r.CONFIRMED.s + ' and a receipt confirmation email has been sent.');
                        }
                    } // for(var i in submissions)
                    //  else if (currentDate<closeOfSubmission)
                } else if (currentSelectionNotification !== r.NOT_STARTED.s) {
                    // SELECTION NOTIFICATION
                    log('hReminderConfirmation:processing Selection Notification');
                    filmIdFormat = r.FILM_ID.s + pad(-1);
                    processNextSubmission = currentSelectionNotification === r.PENDING.s;
                    setToFinishedAtEnd = true;
                    log('hReminderConfirmation:currentSelectionNotification:' + currentSelectionNotification);
                    // Loop through all submissions until we come to the first submission after r.CURRENT_SELECTION_NOTIFICATION.s.
                    // Start with first if r.CURRENT_SELECTION_NOTIFICATION.s === PENDING. Normally this will be the case.
                    for (i = 0; i < submissions.length; i++) {
                        if (emailQuotaRemaining <= r.MIN_QUOTA.n) {
                            log('hReminderConfirmation:ran out of email quota during ' + r.SELECTION_NOTIFICATION.s + ', (r.MIN_QUOTA.n, emailQuotaRemaining):(' + r.MIN_QUOTA.n + ',' + emailQuotaRemaining + '). Terminating ' + r.SELECTION_NOTIFICATION.s + ' early.');
                            setToFinishedAtEnd = false;
                            break; //stop sending emails for the day
                        }

                        submission = submissions[i];
                        filmId = submission[normalizeHeader(r.FILM_ID.s)];
                        status = submission[normalizeHeader(r.STATUS.s)];
                        selection = submission[normalizeHeader(r.SELECTION.s)];

                        if (!filmId) {
                            log('hReminderConfirmation:ERROR: exepected film id!');
                            continue;
                        }

                        if (status === r.PROBLEM.s) {
                            // Ingnore submissions with status PROBLEM
                            continue;
                        }

                        matches = filmId.match(filmIdFormat, 'i');
                        if (matches && matches.length > 0) {
                            filmId = matches[0];
                            if (processNextSubmission) {
                                //send SELECTION NOTIFICATION
                                if (selection === r.SELECTED.s) {
                                    emailQuotaRemaining = mergeAndSend(ss, submission, r.ACCEPTED.s, emailQuotaRemaining);
                                } else {
                                    emailQuotaRemaining = mergeAndSend(ss, submission, r.NOT_ACCEPTED.s, emailQuotaRemaining);
                                }

                                setNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s, filmId); //Update current
                                log('hReminderConfirmation:processed ' + filmId + ' as ' + selection);
                            } else {
                                processNextSubmission = filmId === currentSelectionNotification;
                                log('hReminderConfirmation:found currentSelectionNotification');
                            }
                        } else {
                            log('hReminderConfirmation:ERROR: exepected film id!');
                        }
                    } //for (var i in submissions)
                    if (setToFinishedAtEnd) {
                        setNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s, r.NOT_STARTED.s); //Job done. Reset.
                    }
                    // else if(currentSelectionNotification!==NOT_STARTED)
                }
                log('hReminderConfirmation end');
            } // if(lastRow>1)
        } catch (e) {
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

        for (var i = 0; i < headersOfInterest.length; i++) {
            var header = headersOfInterest[i],
                index = headers.indexOf(header);

            indices[header] = index + 1;

            if (index !== -1) {
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

    test_sfss_interface = {
        normaliseAndValidateDuration: normaliseAndValidateDuration,
        pad: pad,
        setPadNumber: setPadNumber,
        diffDays: diffDays,
        setNamedValue: setNamedValue,
        getNamedValue: getNamedValue,
        saveData: saveData,
        loadData: loadData,
        findMinMaxColumns: findMinMaxColumns,
        normalizeHeader: normalizeHeader,
        log: log,
        findStatusColor: findStatusColor,
        getProperty: getProperty,
        setProperty: setProperty,
        deleteProperty: deleteProperty,
        setTesting: setTesting,
        setLogDoc: setLogDoc
    };

    sfss_interface = {
        onOpen: onOpen,
        setup: setup,

        // system triggers
        hProcessSubmission: hProcessSubmission,
        hReminderConfirmation: hReminderConfirmation,

        // spreadsheet menu items
        mediaPresentNotConfirmed: mediaPresentNotConfirmed,
        problem: problem,
        selected: selected,
        notSelected: notSelected,
        manual: manual,
        settingsOptions: settingsOptions,
        editAndSaveTemplates: editAndSaveTemplates,
        adHocEmail: adHocEmail,
        selectionNotification: selectionNotification,

        // button actions
        buttonAction: buttonAction,
        templatesButtonAction: templatesButtonAction,
        settingsOptionsButtonAction: settingsOptionsButtonAction,
        selectionNotificationButtonAction: selectionNotificationButtonAction,
        adHocEmailButtonAction: adHocEmailButtonAction,

        // testing
        test_sfss_interface: test_sfss_interface
    };

    return sfss_interface;
}(sfss.r, sfss.smm));

//////////////////////////////////////
// build custom menu for spreadsheet
//////////////////////////////////////
function onOpen() {
    sfss.s.sonOpen();
}

//////////////////////////////////////
// setup system
//////////////////////////////////////
function setup() {
    sfss.s.setup();
}

//////////////////////////////////////
// start of system triggers
//////////////////////////////////////
function hProcessSubmission(e) {
    sfss.s.hProcessSubmission(e);
}

function hReminderConfirmation(e) {
    sfss.s.hReminderConfirmation(e);
}
//////////////////////////////////////
// end of system triggers
//////////////////////////////////////


//////////////////////////////
// spreadsheet menu items
//////////////////////////////
function mediaPresentNotConfirmed() {
    sfss.s.mediaPresentNotConfirmed();
}

function problem() {
    sfss.s.problem();
}

function selected() {
    sfss.s.selected();
}

function notSelected() {
    sfss.s.notSelected();
}

function manual() {
    sfss.s.manual();
}

function settingsOptions() {
    sfss.s.settingsOptions();
}

function editAndSaveTemplates() {
    sfss.s.editAndSaveTemplates();
}

function adHocEmail() {
    sfss.s.adHocEmail();
}

function selectionNotification() {
    sfss.s.selectionNotification();
}
///////////////////////////
// end of spreadsheet menu items
///////////////////////////


///////////////////////////////////
// start of button actions
///////////////////////////////////
function buttonAction(e) {
    return sfss.s.buttonAction(e);
}

function templatesButtonAction(e) {
    return sfss.s.templatesButtonAction(e);
}

function settingsOptionsButtonAction(e) {
    return sfss.s.settingsOptionsButtonAction(e);
}

function selectionNotificationButtonAction(e) {
    return sfss.s.selectionNotificationButtonAction(e);
}

function adHocEmailButtonAction(e) {
    return sfss.s.adHocEmailButtonAction(e);
}
///////////////////////////////////
// end of button actions
///////////////////////////////////