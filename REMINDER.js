/*global DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file reminder');
var sfss = sfss || {};
try {
    sfss.reminder = (function (r, lg, u, m, smm) {
        'use strict';

        var reminder_interface,
            getRowsData = smm.getRowsData,
            normalizeHeaders = smm.normalizeHeaders,
            normalizeHeader = smm.normalizeHeader,

            log = lg.log,
            catchToString = lg.catchToString,

            pad = u.pad,
            getNamedValue = u.getNamedValue,
            setNamedValue = u.setNamedValue,
            loadData = u.loadData,
            findStatusColor = u.findStatusColor,
            diffDays = u.diffDays,
            findMinMaxColumns = u.findMinMaxColumns,
            touchSpreadsheet = u.touchSpreadsheet,

            mergeAndSend = m.mergeAndSend;

        function hReminderConfirmation() {
            log('hReminderConfirmation start');
            touchSpreadsheet(reminderConfirmation);
            log('hReminderConfirmation end');
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
        function reminderConfirmation() {
            try {
                log('reminderConfirmation start');
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

                log('reminderConfirmation:enabledReminder: ' + enabledReminder + '\nenableConfirmation: ' + enableConfirmation + '\ncloseOfSubmission: ' + closeOfSubmission + '\ndaysBeforeReminder: ' + daysBeforeReminder + '\ndaysTillClose: ' + daysTillClose + '\nemailQuotaRemaining: ' + emailQuotaRemaining);

                if (lastRow > 1) {
                    log('reminderConfirmation:processing');
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
                        log('reminderConfirmation:processing Ad Hoc Mail');
                        filmIdFormat = r.FILM_ID.s + pad(-1);
                        processNextSubmission = currentAdHocEmail === r.PENDING.s;
                        setToFinishedAtEnd = true;
                        log('reminderConfirmation:currentAdHocEmail:' + currentAdHocEmail);
                        // Loop through all submissions until we come to the first submission after r.CURRENT_AD_HOC_EMAIL.s.
                        // Start with first if currentAdHocEmail === PENDING. Normally this will be the case.
                        for (i = 0; i < submissions.length; i++) {
                            if (emailQuotaRemaining <= r.MIN_QUOTA.n) {
                                log('reminderConfirmation:ran out of email quota during ' + r.AD_HOC_EMAIL.s + ', (r.MIN_QUOTA.n, emailQuotaRemaining):(' + r.MIN_QUOTA.n + ',' + emailQuotaRemaining + '). Terminating ' + r.AD_HOC_EMAIL.s + ' early.');
                                setToFinishedAtEnd = false;
                                break; //stop sending emails for the day
                            }

                            submission = submissions[i];
                            status = submission[normalizeHeader(r.STATUS.s)];
                            filmId = submission[normalizeHeader(r.FILM_ID.s)];

                            if (!filmId) {
                                log('reminderConfirmation:ERROR: exepected film id!');
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
                                    log('reminderConfirmation:Ad Hoc Mail:processed:' + filmId);
                                } else {
                                    processNextSubmission = filmId === currentAdHocEmail;
                                    log('reminderConfirmation:found currentAdHocEmail');
                                }
                            } else {
                                log('reminderConfirmation:ERROR: exepected film id!');
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
                                    log('reminderConfirmation:no submission confirmation has been sent to ' + submission[normalizeHeader(r.FILM_ID.s)] + '.');

                                    emailQuotaRemaining = mergeAndSend(ss, submission, r.SUBMISSION_CONFIRMATION.s, emailQuotaRemaining);

                                    filmSheet.getRange(2 + (+i), lastContactColumnIndex, 1, 1).setValue(currentDate);
                                    filmSheet.getRange(2 + (+i), confirmationColumnIndex, 1, 1).setValue(r.CONFIRMED.s);
                                    color = findStatusColor(ss, status, r.CONFIRMED.s, selection);
                                    filmSheet.getRange(2 + (+i), 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);
                                    log('reminderConfirmation:submission confirmation has now been sent to ' + submission[normalizeHeader(r.FILM_ID.s)] + ' and its ' + r.LAST_CONTACT.s + ' updated.');
                                } else if (daysTillClose > daysBeforeReminder && enabledReminder) {
                                    // Need to send reminder to send Media and release form
                                    var lastContact = submission[normalizeHeader(r.LAST_CONTACT.s)];

                                    var daysSinceLastContact = diffDays(currentDate, lastContact);

                                    if (daysSinceLastContact >= daysBeforeReminder) {
                                        //send reminder
                                        log('It is ' + daysSinceLastContact + ' days since ' + r.LAST_CONTACT.s + ' with submission ' + submission[normalizeHeader(r.FILM_ID.s)] + ' which has status ' + r.NO_MEDIA.s + '.');

                                        emailQuotaRemaining = mergeAndSend(ss, submission, r.REMINDER.s, emailQuotaRemaining);

                                        filmSheet.getRange(2 + (+i), lastContactColumnIndex, 1, 1).setValue(currentDate);
                                        log('reminderConfirmation:a reminder has been sent to submission ' + submission[normalizeHeader(r.FILM_ID.s)] + ' and its ' + r.LAST_CONTACT.s + ' updated.');
                                    }
                                }

                            } else if (status === r.MEDIA_PRESENT.s && confirmation === r.NOT_CONFIRMED.s && enableConfirmation) {
                                //send confirmation of receipt of media
                                log('reminderConfirmation:submission ' + submission[normalizeHeader(r.FILM_ID.s)] + ' has status ' + r.MEDIA_PRESENT.s + ', ' + r.NOT_CONFIRMED.s + '.');

                                emailQuotaRemaining = mergeAndSend(ss, submission, r.RECEIPT_CONFIMATION.s, emailQuotaRemaining);

                                filmSheet.getRange(2 + (+i), lastContactColumnIndex, 1, 1).setValue(currentDate);
                                filmSheet.getRange(2 + (+i), confirmationColumnIndex, 1, 1).setValue(r.CONFIRMED.s);
                                color = findStatusColor(ss, status, r.CONFIRMED.s, selection);
                                filmSheet.getRange(2 + (+i), 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);
                                log('reminderConfirmation:the status of submission ' + submission[normalizeHeader(r.FILM_ID.s)] + ' has been updated to ' + r.MEDIA_PRESENT.s + ', ' + r.CONFIRMED.s + ' and a receipt confirmation email has been sent.');
                            }
                        } // for(var i in submissions)
                        //  else if (currentDate<closeOfSubmission)
                    } else if (currentSelectionNotification !== r.NOT_STARTED.s) {
                        // SELECTION NOTIFICATION
                        log('reminderConfirmation:processing Selection Notification');
                        filmIdFormat = r.FILM_ID.s + pad(-1);
                        processNextSubmission = currentSelectionNotification === r.PENDING.s;
                        setToFinishedAtEnd = true;
                        log('reminderConfirmation:currentSelectionNotification:' + currentSelectionNotification);
                        // Loop through all submissions until we come to the first submission after r.CURRENT_SELECTION_NOTIFICATION.s.
                        // Start with first if r.CURRENT_SELECTION_NOTIFICATION.s === PENDING. Normally this will be the case.
                        for (i = 0; i < submissions.length; i++) {
                            if (emailQuotaRemaining <= r.MIN_QUOTA.n) {
                                log('reminderConfirmation:ran out of email quota during ' + r.SELECTION_NOTIFICATION.s + ', (r.MIN_QUOTA.n, emailQuotaRemaining):(' + r.MIN_QUOTA.n + ',' + emailQuotaRemaining + '). Terminating ' + r.SELECTION_NOTIFICATION.s + ' early.');
                                setToFinishedAtEnd = false;
                                break; //stop sending emails for the day
                            }

                            submission = submissions[i];
                            filmId = submission[normalizeHeader(r.FILM_ID.s)];
                            status = submission[normalizeHeader(r.STATUS.s)];
                            selection = submission[normalizeHeader(r.SELECTION.s)];

                            if (!filmId) {
                                log('reminderConfirmation:ERROR: exepected film id!');
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
                                    log('reminderConfirmation:processed ' + filmId + ' as ' + selection);
                                } else {
                                    processNextSubmission = filmId === currentSelectionNotification;
                                    log('reminderConfirmation:found currentSelectionNotification');
                                }
                            } else {
                                log('reminderConfirmation:ERROR: exepected film id!');
                            }
                        } //for (var i in submissions)
                        if (setToFinishedAtEnd) {
                            setNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s, r.NOT_STARTED.s); //Job done. Reset.
                        }
                        // else if(currentSelectionNotification!==NOT_STARTED)
                    }
                    log('reminderConfirmation end');
                } // if(lastRow>1)
            } catch (e) {
                log('ThReminderConfirmation:error:' + catchToString(e));
            }

        }

        reminder_interface = {
            hReminderConfirmation: hReminderConfirmation
        };

        return reminder_interface;
    }(sfss.r, sfss.lg, sfss.u, sfss.merge, sfss.smm));
} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}
Logger.log('leaving file reminder');