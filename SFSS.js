/*global DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file sfss');
var sfss = sfss || {};
try {
    sfss.s = (function (r, lg, u, ui, smm) {
        'use strict';

        var sfss_interface,
            PROPERTIES = {},

            fillInTemplateFromObject = smm.fillInTemplateFromObject,
            setRowsData = smm.setRowsData,
            getRowsData = smm.getRowsData,
            normalizeHeaders = smm.normalizeHeaders,
            normalizeHeader = smm.normalizeHeader,

            log = lg.log,
            catchToString = lg.catchToString,
            
            setColumnWidth = u.setColumnWidth,
            prettyPrintDate = u.prettyPrintDate,
            normaliseAndValidateDuration = u.normaliseAndValidateDuration,
            pad = u.pad,
            setPadNumber = u.setPadNumber,
            getNamedValue = u.getNamedValue,
            setNamedValue = u.setNamedValue,
            loadData = u.loadData,
            saveData = u.saveData,
            findStatusColor = u.findStatusColor,
            diffDays = u.diffDays,
            findMinMaxColumns = u.findMinMaxColumns,
            mergeAndSend = u.mergeAndSend;


        function setLogDoc(lD) {
            ui.setLogDoc(lD);
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


        function setup() {
            try {
                if (r.TESTING.b || !r.SCRIPT_PROPERTIES.d.getProperty("initialised")) {
                    if (!r.TESTING.b) {
                        r.SCRIPT_PROPERTIES.d.setProperty("initialised", "initialised");
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


        testing = {
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
            setTesting: setTesting
        };

        sfss_interface = {
            setup: setup,

            // system triggers
            hProcessSubmission: hProcessSubmission,
            hReminderConfirmation: hReminderConfirmation,

            // testing
            testing: testing
        };

        return sfss_interface;
    }(sfss.r, sfss.lg, sfss.u, sfss.ui, sfss.smm));
} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}

//////////////////////////////////////
// build custom menu for spreadsheet
//////////////////////////////////////
function onOpen() {
    sfss.ui.onOpen();
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



Logger.log('leaving file sfss');