/*global HtmlService, DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file ui');
var sfss = sfss || {};

try {
    sfss.ui = (function (s) {
        'use strict';

        var ui_interface = {},
            r = s.r,
            lg = s.lg,
            u = s.u,
            smm = s.smm,
            normalizeHeaders = smm.normalizeHeaders,
            normalizeHeader = smm.normalizeHeader,
            getRowsData = smm.getRowsData,
            setRowsData = smm.setRowsData,

            log = lg.log,
            catchToString = lg.catchToString,

            findStatusColor = u.findStatusColor,
            loadData = u.loadData,
            prettyPrintDate = u.prettyPrintDate,
            getNamedValue = u.getNamedValue,
            diffDays = u.diffDays,
            findMinMaxColumns = u.findMinMaxColumns,
            setNamedValue = u.setNamedValue;

        // build custom menu for spreadsheet
        function onOpen() {
            // spreadsheet menu
            try {
                var ss = r.SS.d,
                    initialised = r.SCRIPT_PROPERTIES.d.getProperty("initialised");

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
            var value = r.SCRIPT_PROPERTIES.d.getProperty(normalizeHeader(status + ' ' + confirmation + ' ' + selection));
            if (value) {
                setStatus(status, confirmation, selection);
            } else {
                var data = {},
                    template, dialog;
                data.sfss = sfss;
                data.title = title;
                data.height = height;
                data.width = 700;
                data.html = html;
                data[normalizeHeader(r.STATUS.s)] = status;
                data[normalizeHeader(r.CONFIRMATION.s)] = confirmation;
                data[normalizeHeader(r.SELECTION.s)] = selection;

                template = HtmlService.createTemplateFromFile(r.STATUS.s);
                template.data = data;

                dialog = template.evaluate().setHeight(data.height).setWidth(data.width).setSandboxMode(HtmlService.SandboxMode.IFRAME);
                SpreadsheetApp.getUi().showModalDialog(dialog, ' ');
            }
        }

        function mediaPresentNotConfirmed() {
            var ss = r.SS.d;
            statusDialog(r.MEDIA_PRESENT.s, r.NOT_CONFIRMED.s, r.DO_NOT_CHANGE.s, 650, 'Media Present, Not Confirmed', '<p>The state of a submission is given by the value of three fields, <b>' + r.STATUS.s + '</b>, <b>' + r.CONFIRMATION.s + '</b> and <b>' + r.SELECTION.s + '</b>.</p><p>Setting a submission to ' + r.STATUS.s + ': <b>' + r.MEDIA_PRESENT.s + '</b>, ' + r.CONFIRMATION.s + ': <b>' + r.NOT_CONFIRMED.s + '</b> (and leaving ' + r.SELECTION.s + ' unchanged) means the filmmaker will be automatically emailed after midnight to confirm the receipt of the Physical Media and the Permission Slip at the festival office, provided that:<ul><li>close of submission on ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + ' has not been reached at the time that the email confirmation is attempted</li><li>and that confirmations are currently enabled.</li></ul></p><p>Note: receipt confirmations are currently ' + getNamedValue(ss, r.ENABLE_CONFIRMATION.s) + '.</p><p>If you re-set the submission to a new different state before midnight, the receipt confirmation email will not be sent.</p><p>After the receipt confirmation email is sent the state of the submission will be set to ' + r.STATUS.s + ': <b>' + r.MEDIA_PRESENT.s + '</b>, ' + r.CONFIRMATION.s + ': <b>' + r.CONFIRMED.s + '</b>.</p><p>To continue with setting the submission state to ' + r.STATUS.s + ': <b>' + r.MEDIA_PRESENT.s + '</b>, ' + r.CONFIRMATION.s + ': <b>' + r.NOT_CONFIRMED.s + '</b> press OK, to not do this press CANCEL.</p><br/>');
        }

        function problem() {
            var ss = r.SS.d;
            statusDialog(r.PROBLEM.s, r.DO_NOT_CHANGE.s, r.DO_NOT_CHANGE.s, 480, 'Problem', '<p>The state of a submission is given by the value of three fields, <b>' + r.STATUS.s + '</b>, <b>' + r.CONFIRMATION.s + '</b> and <b>' + r.SELECTION.s + '</b>.</p><p>Setting a submission to ' + r.STATUS.s + ': <b>' + r.PROBLEM.s + '</b> (and leaving ' + r.CONFIRMATION.s + ' and ' + r.SELECTION.s + ' unchanged) means the filmmaker will be sent no reminder or receipt confirmation emails from now on. They will however, receive the final selection advisory email after close of submissions(' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + '), when it is sent.</p><p>If you set a submission to ' + r.STATUS.s + ': <b>' + r.PROBLEM.s + '</b>, please update the submission comment field with the details of the problem.</p><p>To continue with setting the submission state to ' + r.STATUS.s + ': <b>' + r.PROBLEM.s + '</b> press OK, to not do this press CANCEL.</p><br/>');
        }

        function selected() {
            var ss = r.SS.d;
            statusDialog(r.DO_NOT_CHANGE.s, r.DO_NOT_CHANGE.s, r.SELECTED.s, 680, 'Selected', '<p>The state of a submission is given by the value of three fields, <b>' + r.STATUS.s + '</b>, <b>' + r.CONFIRMATION.s + '</b> and <b>' + r.SELECTION.s + '</b>.</p><p>A submission set to ' + r.SELECTION.s + ': <b>' + r.SELECTED.s + '</b> (and leaving ' + r.STATUS.s + ' and ' + r.CONFIRMATION.s + ' unchanged) will still receive reminder and receipt confirmation emails dependent on its <b>' + r.STATUS.s + '</b> and <b>' + r.CONFIRMATION.s + '</b> provided that:<ul><li>it is at least ' + getNamedValue(ss, r.DAYS_BEFORE_REMINDER.s) + ' days before the close of submission on ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + ' at the time that the email is attempted</li><li>and that receipt confirmations are currently enabled for receipt confirmation,</li><li>and that reminders are currently enabled for reminders.</li></ul></p><p>Note: receipt confirmations are currently ' + getNamedValue(ss, r.ENABLE_CONFIRMATION.s) + '.</p><p>Note: reminders are currently ' + getNamedValue(ss, r.ENABLE_REMINDER.s) + '.</p><p>Submissions set to ' + r.SELECTION.s + ': <b>' + r.SELECTED.s + '</b> will receive the festival submission acceptance email after close of submissions(' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + '), when it is sent.</p><p>You will usually only set this state after the close of submissions.</p><p>To continue with setting the submission state to ' + r.SELECTION.v + ': <b>' + r.SELECTED.v + '</b> press OK, to not do this press CANCEL.</p><br/>');
        }

        function notSelected() {
            var ss = r.SS.d;
            statusDialog(r.DO_NOT_CHANGE.s, r.DO_NOT_CHANGE.s, r.NOT_SELECTED.s, 700, 'Not Selected', '<p>The state of a submission is given by the value of three fields, <b>' + r.STATUS.s + '</b>, <b>' + r.CONFIRMATION.s + '</b> and <b>' + r.SELECTION.s + '</b>.</p><p>A submission set to ' + r.SELECTION.s + ': <b>' + r.NOT_SELECTED.s + '</b> (and leaving ' + r.STATUS.s + ' and ' + r.CONFIRMATION.s + ' unchanged) will still receive reminder and receipt confirmation emails dependent on its <b>' + r.STATUS.s + '</b> and <b>' + r.CONFIRMATION.s + '</b> provided that:<ul><li>it is at least ' + getNamedValue(ss, r.DAYS_BEFORE_REMINDER.s) + ' days before the close of submission on ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + ' at the time that the email is attempted</li><li>and that receipt confirmations are currently enabled for receipt confirmation,</li><li>and that reminders are currently enabled for reminders.</li></ul></p><p>Note: receipt confirmations are currently ' + getNamedValue(ss, r.ENABLE_CONFIRMATION.s) + '.</p><p>Note: reminders are currently ' + getNamedValue(ss, r.ENABLE_REMINDER.s) + '.</p><p>Submissions set to ' + r.SELECTION.s + ': <b>' + r.NOT_SELECTED.s + '</b> will receive the festival submission not acceptanced email after close of submissions(' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)) + '), when it is sent.</p><p>You will usually only set this state after the close of submissions.</p><p>To continue with setting the submission state to ' + r.SELECTION.s + ': <b>' + r.NOT_SELECTED.s + '</b> press OK, to not do this press CANCEL.</p><br/>');
        }

        function manual() {
        var ss = r.SS.d,
            gridData = [{
                namedValue: r.STATUS.s,
                type: 'list',
                list: [r.DO_NOT_CHANGE.s, r.NO_MEDIA.s, r.MEDIA_PRESENT.s, r.PROBLEM.s]
            }, {
                namedValue: r.CONFIRMATION.s,
                type: 'list',
                list: [r.DO_NOT_CHANGE.s, r.NOT_CONFIRMED.s, r.CONFIRMED.s]
            }, {
                namedValue: r.SELECTION.s,
                type: 'list',
                list: [r.DO_NOT_CHANGE.s, r.NOT_SELECTED.s, r.SELECTED.s]
            }],
            template, dialog;

        gridData.sfss = sfss;
        gridData.title = 'Set State Of Selection';
        gridData.height = 350;
        gridData.width = 600;

        template = HtmlService.createTemplateFromFile(r.MANUAL.s);
        template.gridData = gridData;
        template.sfss = sfss;
        template.ss = ss;

        dialog = template.evaluate().setHeight(gridData.height).setWidth(gridData.width).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        SpreadsheetApp.getUi().showModalDialog(dialog, ' ');    
    }

    function buttonAction(f) {
        try {
            var ss = r.SS.d,
                status = f[normalizeHeader(r.STATUS.s)],
                confirmation = f[normalizeHeader(r.CONFIRMATION.s)],
                selection = f[normalizeHeader(r.SELECTION.s)],
                checked = f[normalizeHeader(r.INHIBIT.s)];

            r.SCRIPT_PROPERTIES.d.setProperty(normalizeHeader(status + ' ' + confirmation + ' ' + selection), checked);

            setStatus(status, confirmation, selection);
        } catch (e) {
            log('There has been an error in buttonAction:' + e);
        }
    }

    function setStatus(status, confirmation, selection) {
        try {
            log('setStatus:(status, confirmation, selection):(' + status + ',' + confirmation + ',' + selection + ')');
            var ss = r.SS.d;
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
                    ss.toast("Update has been made.", "INFORMATION", 5);
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

            var ss = r.SS.d;
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
                ss = r.SS.d,
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

    function selectionNotificationAcceptedTemplate() {
        editAndSaveTemplates(r.ACCEPTED.s);
    }

    function selectionNotificationNotAcceptedTemplate() {
        editAndSaveTemplates(r.NOT_ACCEPTED.s);
    }

    function selectionNotification() {
        try {
            var ss = r.SS.d;

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
            buttonGrid.setWidget(0, 0, getNamedValue(ss, r.LAST_PROCESS_INDEX.s) === r.NOT_STARTED.s ? enable : unenable);
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
            ss = r.SS.d;

        if (e.parameter.source === r.ENABLE.s) {
            log(r.SELECTION_NOTIFICATION.s + ' set to:' + r.PENDING.s);
            setNamedValue(ss, r.CURRENT_PROCESS.s, r.SELECTION_NOTIFICATION_PROCESS.s);
            setNamedValue(ss, r.LAST_PROCESS_INDEX.s, r.NOT_STARTED.s);
            ss.toast("Selection Notification has been queued.", "INFORMATION", 5);
        } else if (e.parameter.source === r.UNENABLE.s) {
            log(r.SELECTION_NOTIFICATION.s + ' set to:' + r.NOT_STARTED.s);
            setNamedValue(ss, r.CURRENT_PROCESS.s, r.DEFAULT_PROCESS.s);
            setNamedValue(ss, r.LAST_PROCESS_INDEX.s, r.NOT_STARTED.s);
            ss.toast("Selection Notification has been unqueued.", "INFORMATION", 5);
        } else {
            ss.toast("Canceling operation.", "WARNING", 5);
        }
        return app.close();
    }

    function adHocEmailTemplate() {
        editAndSaveTemplates(r.AD_HOC_EMAIL.s);
    }

    function adHocEmail() {
        try {
            var ss = r.SS.d;

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

            buttonGrid.setWidget(0, 0, getNamedValue(ss, r.CURRENT_PROCESS.s) === r.AD_HOC_PROCESS.s ? unenable : enable);
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
            app.setHeight('350');
            ss.show(app);
        } catch (e) {
            log('adHocEmail:error:' + catchToString(e));
        }
    }

    function adHocEmailButtonAction(e) {
        var app = UiApp.getActiveApplication(),
            ss = r.SS.d;
        log('adHocEmailButtonAction:e.parameter.source:' + e.parameter.source);
        if (e.parameter.source === normalizeHeader(r.UNENABLE.s)) {
            setNamedValue(ss, r.CURRENT_PROCESS.s, r.DEFAULT_PROCESS.s);
            setNamedValue(ss, r.LAST_PROCESS_INDEX.s, r.NOT_STARTED.s);
            ss.toast("Ad Hoc Email has been unqueued.", "INFORMATION", 5);
            app.close();
        } else if (e.parameter.source === normalizeHeader(r.ENABLE.s)) {
            setNamedValue(ss, r.CURRENT_PROCESS.s, r.AD_HOC_PROCESS.s);
            setNamedValue(ss, r.LAST_PROCESS_INDEX.s, r.NOT_STARTED.s);
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

    function settingsOptions() {
        var ss = r.SS.d,
            festivaDataNames = u.loadData(ss, r.FESTIVAL_DATA.s).join(',');

        setNamedValue(ss, r.FESTIVAL_DATA_NAMES.s, festivaDataNames);

        var gridData = [{
            namedValue: r.FESTIVAL_DATA_NAMES.s,
            type: 'hidden'
        }, {
            namedValue: r.FESTIVAL_NAME.s,
            type: 'text',
            help: 'required, please give the full name of the festival',
            rules: {
                required: true,
                minlength: 1
            },
            messages: {
                required: 'Festival Name required!',
                minlength: 'Festival Name required!'
            }
        }, {
            namedValue: r.FESTIVAL_WEBSITE.s,
            type: 'text',
            help: 'not required',
            rules: {
                required: false
            }
        }, {
            namedValue: r.CLOSE_OF_SUBMISSION.s,
            type: 'date',
            help: 'current value ' + prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s)),
            rules: {
                required: true,
                dateISO: true,
                before: normalizeHeader(r.EVENT_DATE.s)
            },
            messages: {
                required: 'Close of Submission required!',
                dateISO: 'ISO format date required yyyy-mm-dd!',
                before: 'Close of Submission must be before event!'
            }
        }, {
            namedValue: r.EVENT_DATE.s,
            type: 'date',
            help: 'current value ' + prettyPrintDate(getNamedValue(ss, r.EVENT_DATE.s)),
            rules: {
                required: true,
                dateISO: true,
                after: normalizeHeader(r.CLOSE_OF_SUBMISSION.s)
            },
            messages: {
                required: 'Event date required!',
                dateISO: 'ISO format date required yyyy-mm-dd!',
                after: 'Event must be after Close of Submission!'
            }
        }, {
            namedValue: r.RELEASE_LINK.s,
            type: 'text',
            help: 'not required',
            rules: {
                required: false
            }
        }, {
            namedValue: r.DAYS_BEFORE_REMINDER.s,
            type: 'text',
            help: 'please give days before reminder email is sent',
            rules: {
                required: true,
                min: 1
            },
            messages: {
                required: 'Days Before Reminder required',
                min: 'Must be whole positive number'
            }
        }, {
            namedValue: r.ENABLE_REMINDER.s,
            type: 'list',
            help: 'current value ' + getNamedValue(ss, r.ENABLE_REMINDER.s),
            rules: {
                required: true
            },
            list: [r.NOT_ENABLED.s, r.ENABLED.s]
        }, {
            namedValue: r.ENABLE_CONFIRMATION.s,
            type: 'list',
            help: 'current value ' + getNamedValue(ss, r.ENABLE_REMINDER.s),
            rules: {
                required: true
            },
            list: [r.NOT_ENABLED.s, r.ENABLED.s]
        }],
            template, html;

        gridData.formId = normalizeHeader(r.OPTIONS_SETTINGS_SHEET.s);
        gridData.legend = r.OPTIONS_SETTINGS_SHEET.s;
        gridData.height = 700;
        gridData.width = 680;

        template = HtmlService.createTemplateFromFile('Options');
        template.gridData = gridData;
        template.sfss = sfss;
        template.ss = ss;

        html = template.evaluate().setHeight(gridData.height).setWidth(gridData.width).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        SpreadsheetApp.getUi().showModalDialog(html, ' ');
    }

    function settingsOptionsAction(f) {
        var ss = r.SS.d,
            eventDate = new Date(),
            cosDate = new Date(),
            eventString, cosString;

        // The data should be correctly formatted at this point,
        // so throw exception if it is not.

        // At present 2015-04-29  GAS does not understand ISO formatted dates
        eventString = f.eventDate;
        cosString = f.closeOfSubmission;
        [{
            's': eventString,
            'd': eventDate
        }, {
            's': cosString,
            'd': cosDate
        }].forEach(function (item) {
            var s = item.s,
                d = item.d;
            if (!s.match(/\d{4}-\d\d-\d\d/)) {
                throw {
                    message: 'Expected date in ISO format!'
                };
            }
            d.setFullYear(s.slice(0, 4));
            d.setMonth(s.slice(5, 7) - 1);
            d.setDate(s.slice(8, 10));
            d.setHours(0);
            d.setMinutes(0);
            d.setSeconds(0);
            log('d, s:' + d + ',' + s);
        });

        if (isNaN(eventDate.getTime()) || isNaN(cosDate.getTime()) || eventDate < cosDate) {
            throw {
                message: 'Expected Close of Submission to be before the event date!'
            };
        }

        var festivaDataNames = f[normalizeHeader(r.FESTIVAL_DATA_NAMES.s)].split(','),
            range = ss.getRangeByName(normalizeHeader(r.FESTIVAL_DATA.s)),
            values = festivaDataNames.map(function (itemName) {
                return [itemName, f[normalizeHeader(itemName)]];
            });
        range.setValues(values);

        ss.toast("Saved!", "INFORMATION", 5);

    }

    function cancelAction(f) {
        var ss = r.SS.d,
            status = f[normalizeHeader(r.STATUS.s)],
            confirmation = f[normalizeHeader(r.CONFIRMATION.s)],
            selection = f[normalizeHeader(r.SELECTION.s)],
            checked = f[normalizeHeader(r.INHIBIT.s)];
        log('cancelAction:' + status + ' ' + confirmation + ' ' + selection);
        if (status) {
            r.SCRIPT_PROPERTIES.d.setProperty(normalizeHeader(status + ' ' + confirmation + ' ' + selection), checked);
            log('cancelAction: setting checked:' + status + ' ' + confirmation + ' ' + selection);
        }
        ss.toast("Canceling operation.", "WARNING", 5);
        log('cancelAction: canceling action');
    }


    ui_interface = {
        onOpen: onOpen,

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

        selectionNotificationAcceptedTemplate: selectionNotificationAcceptedTemplate,
        selectionNotificationNotAcceptedTemplate: selectionNotificationNotAcceptedTemplate,
        adHocEmailTemplate: adHocEmailTemplate,

        // button actions and server handlers
        buttonAction: buttonAction,
        templatesButtonAction: templatesButtonAction,
        settingsOptionsAction: settingsOptionsAction,
        selectionNotificationButtonAction: selectionNotificationButtonAction,
        adHocEmailButtonAction: adHocEmailButtonAction,
        cancelAction: cancelAction
    };

    return ui_interface;

    }(sfss));
} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}
Logger.log('leaving file ui');