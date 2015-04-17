/*global DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file reminder');
var sfss = sfss || {};
try {
    sfss.reminder = (function (s) {
        'use strict';

        var reminder_interface, r = s.r,
            lg = s.lg,
            u = s.u,
            m = s.merge,
            smm = s.smm,

            getRowsData = smm.getRowsData,
            normalizeHeaders = smm.normalizeHeaders,
            normalizeHeader = smm.normalizeHeader,

            log = lg.log,

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
            log('hNewReminderConfirmation start');
            touchSpreadsheet(reminderConfirmation);
            log('hNewReminderConfirmation end');
        }

        function callbackFactory(ss, filmSheet, indices, numberOfSubmissions) {
            var emailQuotaRemaining = MailApp.getRemainingDailyQuota(),
                enabledReminder = getNamedValue(ss, r.ENABLE_REMINDER.s) === r.ENABLED.s,
                enableConfirmation = getNamedValue(ss, r.ENABLE_CONFIRMATION.s) === r.ENABLED.s,
                currentDate = new Date(),
                closeOfSubmission = getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s),
                daysBeforeReminder = getNamedValue(ss, r.DAYS_BEFORE_REMINDER.s),
                daysTillClose = diffDays(closeOfSubmission, currentDate),

                currentProcess = getNamedValue(ss, r.CURRENT_PROCESS.s),
                lastProcessIndex = getNamedValue(ss, r.LAST_PROCESS_INDEX.s),
                filmIdFormat = r.FILM_ID.s + pad(-1),

                confirmationColumnIndex = indices[normalizeHeader(r.CONFIRMATION.s)],
                lastContactColumnIndex = indices[normalizeHeader(r.LAST_CONTACT.s)];

            function defaultProcess(submission, index) {
                var templateName, color, newStates, status, confirmation, selection, lastContact, daysSinceLastContact;

                status = submission[normalizeHeader(r.STATUS.s)];
                confirmation = submission[normalizeHeader(r.CONFIRMATION.s)];
                selection = submission[normalizeHeader(r.SELECTION.s)];

                if (currentDate >= closeOfSubmission) {
                    templateName = r.NO_TEMPLATE.s;
                }

                if (!templateName) {
                    if (status === r.NO_MEDIA.s && confirmation === r.NOT_CONFIRMED.s) {
                        //Send submission confirmation. This should be a very rare case!
                        templateName = r.SUBMISSION_CONFIRMATION.s;
                        color = findStatusColor(ss, status, r.CONFIRMED.s, selection);
                        newStates = [{
                            range: filmSheet.getRange(2 + (+index), lastContactColumnIndex, 1, 1),
                            value: currentDate
                        }, {
                            range: filmSheet.getRange(2 + (+index), confirmationColumnIndex, 1, 1),
                            value: r.CONFIRMED.s
                        }];
                    }
                }

                if (!templateName) {
                    if (status === r.NO_MEDIA.s && daysTillClose > daysBeforeReminder && enabledReminder) {
                        // Need to send reminder to send Media and release form
                        lastContact = submission[normalizeHeader(r.LAST_CONTACT.s)];
                        daysSinceLastContact = diffDays(currentDate, lastContact);
                        if (daysSinceLastContact >= daysBeforeReminder) {
                            templateName = r.REMINDER.s;
                            newStates = [{
                                range: filmSheet.getRange(2 + (+index), lastContactColumnIndex, 1, 1),
                                value: currentDate
                            }];
                        }
                    }
                }

                if (!templateName) {
                    if (status === r.MEDIA_PRESENT.s && confirmation === r.NOT_CONFIRMED.s && enableConfirmation) {
                        //send confirmation of receipt of media
                        templateName = r.RECEIPT_CONFIMATION.s;
                        color = findStatusColor(ss, status, r.CONFIRMED.s, selection);
                        newStates = [{
                            range: filmSheet.getRange(2 + (+index), lastContactColumnIndex, 1, 1),
                            value: currentDate
                        }, {
                            range: filmSheet.getRange(2 + (+index), confirmationColumnIndex, 1, 1),
                            value: r.CONFIRMED.s
                        }];
                    }
                }

                if (!templateName) {
                    templateName = r.NO_TEMPLATE.s;
                }

                return {
                    templateName: templateName,
                    color: color,
                    newStates: newStates
                };
            }

            function adHocEmail(submission, index) {
                var templateName, status;
                status = submission[normalizeHeader(r.STATUS.s)];
                if (status === r.PROBLEM.s) {
                    templateName = r.NO_TEMPLATE.s;
                } else {
                    templateName = r.AD_HOC_EMAIL.s;
                }
                return {
                    templateName: templateName,
                    color: undefined,
                    newStates: undefined
                };
            }

            function notification(submission, index) {
                var templateName, status;
                status = submission[normalizeHeader(r.STATUS.s)];
                selection = submission[normalizeHeader(r.SELECTION.s)];
                if (status === r.PROBLEM.s) {
                    templateName = r.NO_TEMPLATE.s;
                } else if (selection === r.SELECTED.s) {
                    templateName = r.ACCEPTED.s;
                } else {
                    templateName = r.NOT_ACCEPTED.s;
                }
                return {
                    templateName: templateName,
                    color: undefined,
                    newStates: undefined
                };
            }

            return function (submission, index) {
                var filter, templateData, templateName, filmId, filmIndex, matches, stop = false;

                if (r.MIN_QUOTA.n < emailQuotaRemaining) {

                    filmId = submission[normalizeHeader(r.FILM_ID.s)];
                    log('Processing film: ' + filmId);
                    matches = filmId.match(filmIdFormat, 'i');
                    stop = !matches || matches.length !== 1;
                    if (stop) {
                        log('Malformed or no film ID at index:' + index);
                    }

                    if (!stop) {
                        filmId = matches[0];
                        filmIndex = parseInt(filmId.slice(2), 10);
                        stop = lastProcessIndex >= filmIndex;
                    }

                    if (!stop) {
                        if (currentProcess === r.DEFAULT_PROCESS.s) {
                            filter = defaultProcess;
                        } else if (currentProcess === r.AD_HOC_PROCESS.s) {
                            filter = adHocEmail;
                        } else { // assert currentProcess === r.SELECTION_NOTIFICATION_PROCESS.s
                            filter = notification;
                        }
                        templateData = filter(submission, index);
                        templateName = templateData.templateName;
                        stop = templateName === r.NO_TEMPLATE.s;
                    }

                    if (!stop) {
                        emailQuotaRemaining = mergeAndSend(ss, submission, templateName, emailQuotaRemaining);
                        if (templateData.color) {
                            filmSheet.getRange(2 + (+index), 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(templateData.color);
                        }
                        if (templateData.newStates) {
                            templateData.newStates.forEach(function (item) {
                                item.range.setValue(item.value);
                            });
                        }
                    }

                    if (filmIndex) {
                        setNamedValue(ss, r.LAST_PROCESS_INDEX.s, filmIndex);
                    }

                    if (index === numberOfSubmissions - 1) {
                        setNamedValue(ss, r.CURRENT_PROCESS.s, r.DEFAULT_PROCESS.s);
                        setNamedValue(ss, r.LAST_PROCESS_INDEX.s, r.NOT_STARTED.s);
                    }
                }
            };
        }

        function getSubmissionData(ss, filmSheet) {
            var minMaxColumns, height, range, submissions;
            minMaxColumns = findMinMaxColumns(filmSheet, r.HEADERS_OF_INTEREST.d);
            height = filmSheet.getDataRange().getLastRow() - 1;
            range = filmSheet.getRange(2, minMaxColumns.min, height, minMaxColumns.max - minMaxColumns.min + 1);
            submissions = getRowsData(filmSheet, range, 1);

            return {
                indices: minMaxColumns.indices,
                submissions: submissions
            };
        }

        function reminderConfirmation() {
            log('newReminderConfirmation start');
            var ss = SpreadsheetApp.getActiveSpreadsheet(),
                filmSheet = ss.getSheetByName(r.FILM_SUBMISSIONS_SHEET.s),
                callback, submissionData, submissions, indices;

            loadData(ss, r.FESTIVAL_DATA.s);
            loadData(ss, r.TEMPLATE_DATA.s);
            loadData(ss, r.INTERNALS.s);
            loadData(ss, r.COLOR_DATA.s);

            submissionData = getSubmissionData(ss, filmSheet);
            submissions = submissionData.submissions;
            indices = submissionData.indices;

            callback = callbackFactory(ss, filmSheet, indices, submissions.length);

            submissions.forEach(callback);

            log('newReminderConfirmation end');
        }

        reminder_interface = {
            hReminderConfirmation: hReminderConfirmation
        };

        return reminder_interface;
    }(sfss));
} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}
Logger.log('leaving file reminder');