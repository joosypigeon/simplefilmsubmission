/*global DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file process');
var sfss = sfss || {};
try {
    sfss.process = (function (r, lg, u, m, smm) {
        'use strict';

        var process_interface, setRowsData = smm.setRowsData,
            getRowsData = smm.getRowsData,
            normalizeHeader = smm.normalizeHeader,

            log = lg.log,
            catchToString = lg.catchToString,


            normaliseAndValidateDuration = u.normaliseAndValidateDuration,
            pad = u.pad,
            getNamedValue = u.getNamedValue,
            setNamedValue = u.setNamedValue,
            loadData = u.loadData,
            findStatusColor = u.findStatusColor,
            findMinMaxColumns = u.findMinMaxColumns,
            touchSpreadsheet = u.touchSpreadsheet,

            mergeAndSend = m.mergeAndSend;

        // This trigger function processes form submissions. Due to the fact that
        // On Form Submit event occationally does not fire
        function hProcessSubmission(event) {
            log('hProcessSubmission start');
            touchSpreadsheet(processSubmission, event);
            log('hProcessSubmission end');
        }

        function processSubmission(event) {
            log('processSubmission start:' + event);
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
                    headersOfInterest, minMaxColumns, lastContact, rowIndex, maxColumn, minColumn, unprocessedRange;

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
                log('processSubmission:(eventRow,lastUnprocessedIndex):(' + eventRow + ',' + lastUnprocessedIndex + ')');

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
                log('processSubmission:error:' + catchToString(e));
            }
            log('processSubmission end');
        }

        process_interface = {
            hProcessSubmission: hProcessSubmission
        };

        return process_interface;
    }(sfss.r, sfss.lg, sfss.u, sfss.merge, sfss.smm));
} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}
Logger.log('leaving file process');