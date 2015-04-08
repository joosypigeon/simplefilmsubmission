/*global DriveApp, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger, QUnit, sfss, module, test */
Logger.log('entering test file');
///////////////////////////////////////////////////////////////////////////////
// A collection of not very unity Unit Tests
// These tests use a fork of QUnit.
// Google Apps Script library key MxL38OxqIK-B73jyDTvCe-OBao7QLBR4j
// https://github.com/simula-innovation/qunit/tree/gas/gas
///////////////////////////////////////////////////////////////////////////////
try {
    function doGet(e) {
        'use strict';

        QUnit.urlParams(e.parameter);
        QUnit.config({
            title: "Unit tests for my project"
        });

        var s = sfss.s.testing,
            r = sfss.r,
            lg = sfss.lg,
            normalizeHeader = s.normalizeHeader,
            logIdBack = r.SCRIPT_PROPERTIES.d.getProperty(normalizeHeader(r.LOG_FILE.s));

        lg.setLogDoc(undefined);
        s.setTesting(true);

        QUnit.load(myTests);

        r.SCRIPT_PROPERTIES.d.setProperty(normalizeHeader(r.LOG_FILE.s), logIdBack);
        s.setTesting(false);

        return QUnit.getHtml();
    }

    // Imports the following functions:
    // ok, equal, notEqual, deepEqual, notDeepEqual, strictEqual,
    // notStrictEqual, throws, module, test, asyncTest, expect
    QUnit.helpers(this);

    function myTests() {
        'use strict';

        Logger.log('start:myTests');
        var s = sfss.s.testing,
            r = sfss.r,
            u = sfss.u,
            lg = sfss.lg,
            smm = sfss.smm,

            normalizeHeader = smm.normalizeHeader,
            
            log = lg.log,
            
            pad = u.pad,
            setPadNumber = u.setPadNumber,
            diffDays = u.diffDays,
            normaliseAndValidateDuration = u.normaliseAndValidateDuration,
            setNamedValue = u.setNamedValue,
            getNamedValue = u.getNamedValue,
            saveData = u.saveData,
            loadData = u.loadData,
            findMinMaxColumns = u.findMinMaxColumns,
            findStatusColor = u.findStatusColor,
            
            getProperty = s.getProperty,
            setProperty = s.setProperty,
            deleteProperty = s.deleteProperty;


        module("utility function module");

        test("duration", function () {
            var durationTestData = [{
                data: '2 minutes     32 seconds',
                expected: '00:02:32'
            }, {
                data: '   4   minutes    ',
                expected: '00:04:00'
            }, {
                data: '   xxxx 9 xxx minutes xxxx 4 xxx seconds',
                expected: '00:09:04'
            }, {
                data: '  xxxx 8  ',
                expected: '00:08:00'
            }, {
                data: 'xxxx12xxxxx11xxx',
                expected: '00:12:11'
            }, {
                data: '90 seconds',
                expected: '90 seconds'
            }];

            for (var i = 0; i < durationTestData.length; i++) {
                var testItem = durationTestData[i];
                equal(normaliseAndValidateDuration(testItem.data), testItem.expected, "Expect result in form HH:MM:SS.");
            }
        });

        test('pad', function () {
            var padTestDataInput = [2, 27, 271, 2718, 27182, 271828],
                padTestDataOutput = [
                    ['002', '027', '271'],
                    ['0002', '0027', '0271', '2718'],
                    ['00002', '00027', '00271', '02718', '27182'],
                    ['000002', '000027', '000271', '002718', '027182', '271828']
                ];

            for (var i = 0; i < 4; i++) {
                r.PAD_NUMBER.n = i + 3;
                setPadNumber(r.PAD_NUMBER.n);
                for (var j = 0; j < r.PAD_NUMBER.n; j++) {
                    strictEqual(pad(padTestDataInput[j]), padTestDataOutput[i][j], 'Require leading zeros on string rep of number with ' + r.PAD_NUMBER.n + ' places.');
                }
            }

            var testData = ['\\d\\d\\d', '\\d\\d\\d\\d', '\\d\\d\\d\\d\\d', '\\d\\d\\d\\d\\d\\d'];

            for (i = 0; i < testData.length; i++) {
                r.PAD_NUMBER.n = (+i) + 3;
                setPadNumber(r.PAD_NUMBER.n);
                strictEqual(pad(-1), testData[i], 'Film ID suffix returned by pad for ' + r.PAD_NUMBER.n + ' places.');
            }
        });

        test('diffDays', function () {
            var testData = [{
                input: {
                    a: new Date('Jan 1 2010 00:00:00'),
                    b: new Date('Jan 1 2010 00:00:00')
                },
                expected: 0
            }, {
                input: {
                    a: new Date('Jan 1 2010 01:00:00'),
                    b: new Date('Jan 1 2010 00:00:00')
                },
                expected: 0
            }, {
                input: {
                    a: new Date('Jan 1 2010 23:59:59'),
                    b: new Date('Jan 1 2010 00:00:00')
                },
                expected: 0
            }, {
                input: {
                    a: new Date('Jan 2 2010 00:00:00'),
                    b: new Date('Jan 1 2010 00:00:00')
                },
                expected: 1
            }, {
                input: {
                    a: new Date('Jan 29 2010'),
                    b: new Date('Jan 1 2010')
                },
                expected: 28
            }, {
                input: {
                    a: new Date('Feb 12 2010'),
                    b: new Date('Jan 1 2010')
                },
                expected: 42
            }, {
                input: {
                    a: new Date('Dec 8 2010'),
                    b: new Date('Jan 1 2010')
                },
                expected: 341
            }, {
                input: {
                    a: new Date('Jan 1 2020'),
                    b: new Date('Jan 1 2010')
                },
                expected: 3652
            }];

            for (var i = 0; i < testData.length; i++) {
                var item = testData[i];
                equal(diffDays(item.input.a, item.input.b), item.expected, 'Difference in days between two dates.');
            }
        });

        // Create a temporary spreadsheet and put it in the folder Test. Put the
        // test folder in the same folder has the spreadsheet containing this script.
        // Note: ssFile and testFolder should have been created at the root of the drive
        var ss = SpreadsheetApp.create('Testing'),
            ssFile = DriveApp.getFileById(ss.getId()),
            testFolder = DriveApp.createFolder('Test');

        DriveApp.removeFile(ssFile);
        testFolder.addFile(ssFile);

        ss.setSpreadsheetTimeZone('GMT');
        SpreadsheetApp.setActiveSpreadsheet(ss);

        SpreadsheetApp.flush();


        test('get/set', function () {
            var sheet1 = ss.getSheetByName('Sheet1'),
                name1 = 'Hello World',
                name2 = 'Goodbye World',
                value1 = 'Hello World',
                value2 = 'Goodbye World';

            setNamedValue(ss, name1, value1);
            strictEqual(getNamedValue(ss, name1), value1, 'Set and get named value.');
            setNamedValue(ss, name1, value2);
            strictEqual(getNamedValue(ss, name1), value2, 'Set and get named value.');

            ss.setNamedRange(normalizeHeader(name2), sheet1.getRange(1, 1, 1, 1));
            setNamedValue(ss, name2, value1);
            strictEqual(getNamedValue(ss, name2), value1, 'Set and get named value.');
            setNamedValue(ss, name2, value2);
            strictEqual(getNamedValue(ss, name2), value2, 'Set and get named value.');

        });

        test('loadData/saveData', function () {
            var inputData = [],
                i, j, name, data;

            for (i = 0; i < 3; i++) {
                inputData[i] = [];
                for (j = 0; j < 3 + i; j++) {
                    name = 'name' + i + 'x' + j;
                    data = 'data' + i + 'x' + j;
                    inputData[i][j] = [name, [name, data]];
                }
            }

            var sheet2 = ss.insertSheet('Sheet2');
            for (i = 0; i < 3; i++) {
                saveData(ss, sheet2.getRange(1, 2 * (+i) + 1, 1, 1), inputData[i], 'Load Test Data' + i, (+i) % 2 ? 'aliceblue' : 'mistyrose');
                loadData(ss, 'Load Test Data' + i);
                for (j = 0; j < 3 + i; j++) {
                    name = 'name' + i + 'x' + j;
                    data = 'data' + i + 'x' + j;
                    strictEqual(getNamedValue(ss, name), data, 'Must be able to save and load named arrays of named data.');
                }
            }



        });

        test('findMinMaxColumns', function () {
            var sheet3 = ss.insertSheet('Sheet3'),
                numberOfHeaders = 25,
                prefix = 'Header',
                rows = [],
                headers = [],
                i;

            for (i = 0; i < numberOfHeaders; i++) {
                headers[i] = prefix + (i + 1);
            }
            rows[0] = headers;
            sheet3.getRange(1, 1, 1, numberOfHeaders).setValues(rows);

            var testData = [{
                input: [prefix + '10'],
                expected: {
                    "indices": {
                        "header10": 10
                    },
                    "max": 10,
                    "min": 10
                }
            }, {
                input: [prefix + '5', prefix + '7', prefix + '9', prefix + '11'],
                expected: {
                    "indices": {
                        "header11": 11,
                        "header5": 5,
                        "header7": 7,
                        "header9": 9
                    },
                    "max": 11,
                    "min": 5
                }
            }, {
                input: [prefix + 'x', prefix + '7', prefix + '9', prefix + '11', prefix + '14'],
                expected:

                {
                    "indices": {
                        "header11": 11,
                        "header14": 14,
                        "header7": 7,
                        "header9": 9,
                        "headerx": 0
                    },
                    "max": 14,
                    "min": 7
                }
            }, {
                input: [prefix + '1', prefix + '7', prefix + numberOfHeaders],
                expected: {
                    "indices": {
                        "header1": 1,
                        "header25": 25,
                        "header7": 7
                    },
                    min: 1,
                    max: numberOfHeaders
                }
            }];

            for (i = 0; i < testData.length; i++) {
                var item = testData[i];
                deepEqual(findMinMaxColumns(sheet3, item.input), item.expected, 'Need to be able to find the min and max column indices of a subset of column headers.');
            }
        });

        var that = this;
        test('setup', function () {
            ////////////////////////////////////////////////////////////////////////
            // First we setup the system.
            // We then set up a mock MailApp which is only overQuota over r.MIN_QUOTA.n
            // We simulate 10 submissions
            ////////////////////////////////////////////////////////////////////////
            r.PAD_NUMBER.n = 3;
            setPadNumber(r.PAD_NUMBER.n);
            equal(true, true, setup());



            (function (global) {
                var NativeMailApp = global.MailApp;

                global.pushMailApp = function (fakeMailApp) {
                    global.MailApp = fakeMailApp;
                };

                global.popMailApp = function () {
                    global.MailApp = NativeMailApp;
                };

            }(that));

            function FakeMailApp() {
                var remainingDailyQuota;

                this.getRemainingDailyQuota = function () {
                    log('FakeMailApp: returning remainingDailyQuota:' + remainingDailyQuota);
                    return remainingDailyQuota;
                };

                this.sendEmail = function (email, subject, body) {
                    remainingDailyQuota -= 1;
                    log('FakeMailApp: sendMail');
                };

                this.setRemainingDailyQuota = function (x) {
                    remainingDailyQuota = x;
                    log('FakeMailApp: set remainingDailyQuota to ' + x);
                };
            }

            pushMailApp(new FakeMailApp());

            ///////////////////////////////////////////////////////////////////////////////
            // The first set of test is for hProcessSubmission. This is the Spreadsheet
            // Form Submission trigger. Its sets the submissions Last Contact, Status,
            // Confirmation, Selection and Film ID. And if the daily email quota is greater
            // than r.MIN_QUOTA.n it sends a Submission Confirmation email.
            //
            // NOTE: if no Submission Confirmation email is sent, the submission is marked
            // as No Media, Not Confirmed. When hReminderConfirmation runs (after midnight)
            // Submission Confirmations are sent depending on the daily eamil quota at that
            // point.
            ///////////////////////////////////////////////////////////////////////////////

            // build submission data
            var filmSheet = ss.getSheetByName(r.FILM_SUBMISSIONS_SHEET.s),
                datesData = ['Jan 1 2010 01:26:09', 'Jan 2 2010 02:36:10', 'Jan 3 2010 03:10:52', 'Jan 4 2010 04:54:31', 'Jan 5 2010 05:24:15', 'Jan 6 2010 06:57:19', 'Jan 7 2010 07:22:55', 'Jan 8 2010 08:33:22', 'Jan 9 2010 09:42:10', 'Jan 10 2010 10:10:32'],
                columns = filmSheet.getRange(1, 1, 1, filmSheet.getDataRange().getLastColumn()).getValues()[0],
                formColumns = r.FORM_DATA.d.filter(function (x) {
                    return x !== 'section';
                }).map(function (x) {
                    return x.title;
                }),
                formColumnsSpreadsheetOrder = [],
                i, j, column, templatesTesting, templates;

            // Every column given by form must be in spreadsheet
            for (i = 0; i < formColumns.length; i++) {
                column = formColumns[i];
                ok(columns.indexOf(column) > -1, "Expect column '" + column + "' to be in spreadsheet.");
            }

            for (i = 0; i < columns.length; i++) {
                column = columns[i];
                if (formColumns.indexOf(column) > -1) {
                    formColumnsSpreadsheetOrder.push(column);
                }
            }

            var overQuota = 5;
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota);

            var spreadsheetValues = [],
                formSubmissionValues = [];

            function cellValue(x) {
                return o[x] ? o[x] : '';
            }

            function submissionValue(x) {
                return o[x];
            }
            for (i = 0; i < 10; i++) {
                var n = ('0' + i).slice(-2),
                    o = {};

                for (j = 0; j < formColumns.length; j++) {
                    column = formColumns[j];
                    o[column] = column + n;

                }
                o[r.TIMESTAMP.s] = new Date(datesData[i] + ' GMT+0000 (GMT)');
                o[r.LENGTH.s] = '3:2';
                o[r.YEAR.s] = (1900 + (10 * (+i))).toString();
                o[r.CONFIRM.s] = 'Confirm';
                spreadsheetValues[i] = columns.map(cellValue);
                formSubmissionValues[i] = formColumnsSpreadsheetOrder.map(submissionValue);

            }
            var width = spreadsheetValues[0].length;
            // Do the film submissions
            for (i = 0; i < spreadsheetValues.length; i++) {
                filmSheet.insertRowAfter((+i) + 1); //just to move the row that the form will populate, one futher on.
                var submission = spreadsheetValues[i],
                    row = [];
                row[0] = submission;
                var range = filmSheet.getRange(2 + (+i), 1, 1, width);
                // Set the submission data in the spreadsheet as though inserted by an attached form
                range.setValues(row);

                //Build the associated Spreadsheet Form Submission object.
                var spreadsheetFormSubmitEvent = {};
                spreadsheetFormSubmitEvent.range = range;
                spreadsheetFormSubmitEvent.values = formSubmissionValues[i];
                spreadsheetFormSubmitEvent.namedValues = {};
                for (j = 0; j < submission.length; j++) {
                    if (formColumns.indexOf(columns[j]) > -1 || columns[j] === r.TIMESTAMP.s) {
                        spreadsheetFormSubmitEvent.namedValues[columns[j]] = [submission[j]];
                    }
                }

                // Clean up previous tests.
                deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

                hProcessSubmission(spreadsheetFormSubmitEvent); //Call the spreadsheet form submission trigget with the event object
                ////Utilities.sleep(3000);

                // Check that the emails we expected to be sent have been sent.
                templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
                if ((+i) < overQuota) {
                    if (templatesTesting) {
                        templates = templatesTesting.split(',');
                    }
                    ok(templatesTesting && templates.length === 1 && templates[0] === r.SUBMISSION_CONFIRMATION.s, 'One Submisson Confirmation email has been sent.');
                } else {
                    ok(!templatesTesting, 'No email sent.');
                }
            }

            // After the film submissions, check that the Status, Confirmation, Selection and Film ID columns are as we expect
            var expectedStatus = [
                [r.NO_MEDIA.s],
                [r.NO_MEDIA.s],
                [r.NO_MEDIA.s],
                [r.NO_MEDIA.s],
                [r.NO_MEDIA.s],
                [r.NO_MEDIA.s],
                [r.NO_MEDIA.s],
                [r.NO_MEDIA.s],
                [r.NO_MEDIA.s],
                [r.NO_MEDIA.s]
            ];
            var statusColumn = columns.indexOf(r.STATUS.s);
            ok(statusColumn > -1, 'There is a ' + r.STATUS.s + ' column.');
            statusColumn += 1;
            deepEqual(filmSheet.getRange(2, statusColumn, 10, 1).getValues(), expectedStatus, 'Expect status ' + r.NO_MEDIA.s + '.');

            var expectedConfirmation = [
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.NOT_CONFIRMED.s],
                [r.NOT_CONFIRMED.s],
                [r.NOT_CONFIRMED.s],
                [r.NOT_CONFIRMED.s],
                [r.NOT_CONFIRMED.s]
            ];
            var confirmationColumn = columns.indexOf(r.CONFIRMATION.s);
            ok(confirmationColumn > -1, 'There is a ' + r.CONFIRMATION.s + ' column.');
            confirmationColumn += 1;
            deepEqual(filmSheet.getRange(2, confirmationColumn, 10, 1).getValues(), expectedConfirmation, 'Expect first 5 submissions to be ' + r.CONFIRMED.s + ', second 5 submissions to be ' + r.NOT_CONFIRMED.s);

            var expectedSelection = [
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s]
            ];
            var selectionColumn = columns.indexOf(r.SELECTION.s);
            ok(selectionColumn > -1, 'There is a ' + r.SELECTION.s + ' column.');
            selectionColumn += 1;
            deepEqual(filmSheet.getRange(2, selectionColumn, 10, 1).getValues(), expectedSelection, 'Expect selection to be ' + r.NOT_SELECTED.s);

            var expectedFilmId = [
                ['ID002'],
                ['ID003'],
                ['ID004'],
                ['ID005'],
                ['ID006'],
                ['ID007'],
                ['ID008'],
                ['ID009'],
                ['ID010'],
                ['ID011']
            ];
            var filmIdColumn = columns.indexOf(r.FILM_ID.s);
            ok(filmIdColumn > -1, 'There is a ' + r.FILM_ID.s + ' column.');
            filmIdColumn += 1;
            deepEqual(filmSheet.getRange(2, filmIdColumn, 10, 1).getValues(), expectedFilmId, 'Expect correctly formated and sequenced ' + r.FILM_ID.s);

            //Utilities.sleep(1000);

            ///////////////////////////////////////////////////////////////////////////////
            // Now we will test hReminderConfirmation which is intended to run once a night
            // after midnight. We will use the 10 submissions already in the spreadsheet.
            // 
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

            // Set the Cose of Submission to something consistant with the Timestamps on the ten film submission
            setNamedValue(ss, r.CLOSE_OF_SUBMISSION.s, new Date('Jun 1 2010 GMT+0000 (GMT)'));

            // Need to Mock the Date constructor
            (function (global) {
                global.NativeDate = global.Date;

                global.pushDate = function (fakeDate) {
                    global.Date = fakeDate;
                };

                global.popDate = function () {
                    global.Date = global.NativeDate;
                };

            }(that));

            function FakeDate1(d) {
                return new NativeDate(d ? d : 'Jan 11 2010 GMT+0000 (GMT)');
            }
            FakeDate1.prototype = Date.prototype;

            ///////////////////////////////////////////////////////////////////////////////
            // Now we test hReminderConfirmation.
            //
            // The next two tests check that hReminderConfirmation can send submission
            // confirmations. That it stops when its daily email quota drops to r.MIN_QUOTA.n
            // and that it can resume sending submissions confirmations when called
            // again
            ///////////////////////////////////////////////////////////////////////////////

            // Set the remaining daily email quote to only 3 emails left
            overQuota = 3;
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate1); //Mock date
            hReminderConfirmation(); //first Reminder Confirmation cycle
            popDate(); // reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === overQuota && templates[0] === r.SUBMISSION_CONFIRMATION.s && templates[1] === r.SUBMISSION_CONFIRMATION.s && templates[2] === r.SUBMISSION_CONFIRMATION.s, 'Three ' + r.SUBMISSION_CONFIRMATION.s + ' emails sent:' + templates.length);
            for (var ii = 0; ii < templates.length; ii++) {
                log('templates[' + i + ']:' + templates[i]);
            }

            expectedConfirmation = [
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.NOT_CONFIRMED.s],
                [r.NOT_CONFIRMED.s]
            ];
            deepEqual(filmSheet.getRange(2, confirmationColumn, 10, 1).getValues(), expectedConfirmation, 'Expect first 7 submissions to be ' + r.CONFIRMED.s + ', second 3 submissions to be ' + r.NOT_CONFIRMED.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

            //Utilities.sleep(3000);
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota to only 3 emails left
            pushDate(FakeDate1); //Mock date
            hReminderConfirmation(); //Second Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 2 && templates[0] === r.SUBMISSION_CONFIRMATION.s && templates[1] === r.SUBMISSION_CONFIRMATION.s, 'Two ' + r.SUBMISSION_CONFIRMATION.s + ' emails sent');
            expectedConfirmation = [
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s]
            ];
            deepEqual(filmSheet.getRange(2, confirmationColumn, 10, 1).getValues(), expectedConfirmation, 'Expect all submissions to be ' + r.CONFIRMED.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));


            function FakeDate2(d) {
                return new NativeDate(d ? d : 'Jan 20 2010 GMT+0000 (GMT)');
            }
            FakeDate2.prototype = Date;

            ///////////////////////////////////////////////////////////////////////////////
            // The next three tests check that hReminderConfirmation can send Reminders
            // when the media is not present for a submissions and a sufficent number of
            // days have passed since the last contact. It is checked that reminders are
            // sent until the daily email quota drops to r.MIN_QUOTA.n, in which case it stops.
            // It is checked that hReminderConfirmation can resume sending Reminders when
            // it stopped in the previous call.
            ///////////////////////////////////////////////////////////////////////////////

            // Expect a sufficent amount of time has not elapsed for reminders to
            // have been sent
            overQuota = 3;
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            setNamedValue(ss, r.DAYS_BEFORE_REMINDER.s, 50); //High r.DAYS_BEFORE_REMINDER.s so no reminder emails need to be sent yet
            pushDate(FakeDate2); //Mock date
            hReminderConfirmation(); //thid Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            ok(!templatesTesting, 'Expect no emails sent i.e. no reminder emails need to be sent yet.');
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

            //Utilities.sleep(3000);

            // Expect that 3 reminders will be sent.
            setNamedValue(ss, r.DAYS_BEFORE_REMINDER.s, 14); // lower r.DAYS_BEFORE_REMINDER.s so that reminder emails will need to be sent
            pushDate(FakeDate2); //Mock date
            hReminderConfirmation(); //fourth Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 3 && templates[0] === r.REMINDER.s && templates[1] === r.REMINDER.s && templates[2] === r.REMINDER.s, 'Expect 3 ' + r.REMINDER.s + ' emails sent');
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

            //Utilities.sleep(3000);

            // Expect that 2 reminders will be sent
            overQuota = 3;
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate2); //Mock date
            hReminderConfirmation(); //fourth Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 2 && templates[0] === r.REMINDER.s && templates[1] === r.REMINDER.s, 'Expect 2 ' + r.REMINDER.s + ' emails sent');
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

            //Utilities.sleep(3000);

            ///////////////////////////////////////////////////////////////////////////////
            // The next 4 tests check that hReminderConfirmation can send Receipt
            // Confirmations. That it stops when its daily email quota drops to r.MIN_QUOTA.n
            // and that it can resume sending submissions confirmations when called
            // again
            ///////////////////////////////////////////////////////////////////////////////
            var statusColumnData = [
                [r.MEDIA_PRESENT.s],
                [r.MEDIA_PRESENT.s],
                [r.NO_MEDIA.s],
                [r.NO_MEDIA.s],
                [r.MEDIA_PRESENT.s],
                [r.MEDIA_PRESENT.s],
                [r.NO_MEDIA.s],
                [r.MEDIA_PRESENT.s],
                [r.NO_MEDIA.s],
                [r.MEDIA_PRESENT.s]
            ];
            filmSheet.getRange(2, statusColumn, 10, 1).setValues(statusColumnData);
            var confirmationColumnData = [
                [r.CONFIRMED.s],
                [r.NOT_CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.NOT_CONFIRMED.s],
                [r.NOT_CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.NOT_CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.NOT_CONFIRMED.s]
            ];
            filmSheet.getRange(2, confirmationColumn, 10, 1).setValues(confirmationColumnData);
            var color;
            for (i = 0; i < confirmationColumnData.length; i++) {
                color = findStatusColor(ss, statusColumnData[i][0], confirmationColumnData[i][0], r.NOT_SELECTED.s);
                filmSheet.getRange(2 + (+i), 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);
            }

            // Expect 3 Receipt Confirmations to be sent
            overQuota = 3;
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate2); //Mock date
            hReminderConfirmation(); //fourth Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 3 && templates[0] === r.RECEIPT_CONFIMATION.s && templates[1] === r.RECEIPT_CONFIMATION.s && templates[2] === r.RECEIPT_CONFIMATION.s, 'Expect 3 ' + r.RECEIPT_CONFIMATION.s + ' emails sent');
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

            expectedConfirmation = [
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.NOT_CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.NOT_CONFIRMED.s]
            ];
            deepEqual(filmSheet.getRange(2, confirmationColumn, 10, 1).getValues(), expectedConfirmation, 'Expect first 7 submissions to be ' + r.CONFIRMED.s + ' and the 9th submission to be ' + r.CONFIRMED.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

            // Expect 2 Receipt Confirmations to be sent
            overQuota = 3;
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate2); //Mock date
            hReminderConfirmation(); //fourth Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 2 && templates[0] === r.RECEIPT_CONFIMATION.s && templates[1] === r.RECEIPT_CONFIMATION.s, 'Expect 2 ' + r.RECEIPT_CONFIMATION.s + ' emails sent');
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

            expectedConfirmation = [
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s],
                [r.CONFIRMED.s]
            ];
            deepEqual(filmSheet.getRange(2, confirmationColumn, 10, 1).getValues(), expectedConfirmation, 'Expect all submissions to be ' + r.CONFIRMED.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

            //Utilities.sleep(1000);

            ///////////////////////////////////////////////////////////////////////////////
            // Test Ad Hoc Email sent before close of submission
            ///////////////////////////////////////////////////////////////////////////////
            // Set Status of submission ID008 to Problem
            // Submissions with Status Problem are ingnored by Ad Hoc Email
            filmSheet.getRange(8, statusColumn, 1, 1).setValues([
                [r.PROBLEM.s]
            ]);
            color = findStatusColor(ss, r.PROBLEM.s, r.CONFIRMED.s, r.NOT_SELECTED.s);
            filmSheet.getRange(8, 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);

            //Set Ad Hoc Email to trigger with the next hReminderConfirmation call
            setNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s, r.PENDING.s);

            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate2); //Mock date
            hReminderConfirmation(); //5th Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            log('templates.length:' + templates.length);
            for (var ii = 0; ii < templates.length; ii++) {
                log('templates[' + ii + ']:' + templates[ii]);
            }
            ok(templatesTesting && templates.length === 3 && templates[0] === r.AD_HOC_EMAIL.s && templates[1] === r.AD_HOC_EMAIL.s && templates[2] === r.AD_HOC_EMAIL.s, 'Expect 3 ' + r.AD_HOC_EMAIL.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            equal(getNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s), r.FILM_ID.s + pad(4), 'Expect ' + r.CURRENT_AD_HOC_EMAIL.s + ' to be ' + r.FILM_ID.s + pad(4));

            //Utilities.sleep(1000);
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate2); //Mock date
            hReminderConfirmation(); //6th Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 3 && templates[0] === r.AD_HOC_EMAIL.s && templates[1] === r.AD_HOC_EMAIL.s && templates[2] === r.AD_HOC_EMAIL.s, 'Expect 3 ' + r.AD_HOC_EMAIL.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            equal(getNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s), r.FILM_ID.s + pad(7), 'Expect ' + r.CURRENT_AD_HOC_EMAIL.s + ' to be ' + r.FILM_ID.s + pad(7));

            //Utilities.sleep(1000);
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate2); //Mock date
            hReminderConfirmation(); //6th Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 3 && templates[0] === r.AD_HOC_EMAIL.s && templates[1] === r.AD_HOC_EMAIL.s && templates[2] === r.AD_HOC_EMAIL.s, 'Expect 3 ' + r.AD_HOC_EMAIL.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            equal(getNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s), r.NOT_STARTED.s, 'Expect ' + r.CURRENT_AD_HOC_EMAIL.s + ' to be ' + r.NOT_STARTED.s);

            //Utilities.sleep(1000);

            ///////////////////////////////////////////////////////////////////////////////
            // Test Ad Hoc Email sent after close of submission
            ///////////////////////////////////////////////////////////////////////////////
            //Set the current date to be one month after the Close of Submission and try Ad Hoc Email again
            function FakeDate3(d) {
                return new NativeDate(d ? d : 'Jul 1 2010 GMT+0000 (GMT)');
            }
            FakeDate3.prototype = Date.prototype;

            //Set Ad Hoc Email to trigger with the next hReminderConfirmation call
            setNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s, r.PENDING.s);

            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate3); //Mock date
            hReminderConfirmation(); //5th Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            log('templates.length:' + templates.length);
            for (var ii = 0; ii < templates.length; ii++) {
                log('templates[' + ii + ']:' + templates[ii]);
            }
            ok(templatesTesting && templates.length === 3 && templates[0] === r.AD_HOC_EMAIL.s && templates[1] === r.AD_HOC_EMAIL.s && templates[2] === r.AD_HOC_EMAIL.s, 'Expect 3 ' + r.AD_HOC_EMAIL.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            equal(getNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s), r.FILM_ID.s + pad(4), 'Expect ' + r.CURRENT_AD_HOC_EMAIL.s + ' to be ' + r.FILM_ID.s + pad(4));

            //Utilities.sleep(1000);
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate3); //Mock date
            hReminderConfirmation(); //6th Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 3 && templates[0] === r.AD_HOC_EMAIL.s && templates[1] === r.AD_HOC_EMAIL.s && templates[2] === r.AD_HOC_EMAIL.s, 'Expect 3 ' + r.AD_HOC_EMAIL.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            equal(getNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s), r.FILM_ID.s + pad(7), 'Expect ' + r.CURRENT_AD_HOC_EMAIL.s + ' to be ' + r.FILM_ID.s + pad(7));

            //Utilities.sleep(1000);
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate3); //Mock date
            hReminderConfirmation(); //6th Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 3 && templates[0] === r.AD_HOC_EMAIL.s && templates[1] === r.AD_HOC_EMAIL.s && templates[2] === r.AD_HOC_EMAIL.s, 'Expect 3 ' + r.AD_HOC_EMAIL.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            equal(getNamedValue(ss, r.CURRENT_AD_HOC_EMAIL.s), r.NOT_STARTED.s, 'Expect ' + r.CURRENT_AD_HOC_EMAIL.s + ' to be ' + r.NOT_STARTED.s);

            //Utilities.sleep(1000);

            ///////////////////////////////////////////////////////////////////////////////
            // Test Selection Notification Email
            ///////////////////////////////////////////////////////////////////////////////
            selectionColumnData = [
                [r.SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.SELECTED.s],
                [r.NOT_SELECTED.s],
                [r.SELECTED.s]
            ];
            filmSheet.getRange(2, selectionColumn, 10, 1).setValues(selectionColumnData);

            //Set Ad Hoc Email to trigger with the next hReminderConfirmation call
            setNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s, r.PENDING.s);

            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate3); //Mock date
            hReminderConfirmation(); //5th Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            log('templatesTesting:' + templatesTesting);
            log(templates);
            log('templates.length:' + templates.length);
            ok(templatesTesting && templates.length === 3 && templates[0] === r.ACCEPTED.s && templates[1] === r.NOT_ACCEPTED.s && templates[2] === r.NOT_ACCEPTED.s, 'Expect ' + r.ACCEPTED.s + ' followed by 2 ' + r.NOT_ACCEPTED.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            equal(getNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s), r.FILM_ID.s + pad(4), 'Expect ' + r.CURRENT_SELECTION_NOTIFICATION.s + ' to be ' + r.FILM_ID.s + pad(4));

            //Utilities.sleep(1000);
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate3); //Mock date
            hReminderConfirmation(); //6th Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 3 && templates[0] === r.NOT_ACCEPTED.s && templates[1] === r.ACCEPTED.s && templates[2] === r.NOT_ACCEPTED.s, 'Expect ' + r.NOT_ACCEPTED.s + ' followed by ' + r.ACCEPTED.s + ' followed by ' + r.NOT_ACCEPTED.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            equal(getNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s), r.FILM_ID.s + pad(7), 'Expect ' + r.CURRENT_SELECTION_NOTIFICATION.s + ' to be ' + r.FILM_ID.s + pad(7));

            //Utilities.sleep(1000);
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            pushDate(FakeDate3); //Mock date
            hReminderConfirmation(); //6th Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 3 && templates[0] === r.ACCEPTED.s && templates[1] === r.NOT_ACCEPTED.s && templates[2] === r.ACCEPTED.s, 'Expect ' + r.ACCEPTED.s + ' follwed by ' + r.NOT_ACCEPTED.s + ' follwed by ' + r.ACCEPTED.s);
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            equal(getNamedValue(ss, r.CURRENT_SELECTION_NOTIFICATION.s), r.NOT_STARTED.s, 'Expect ' + r.CURRENT_SELECTION_NOTIFICATION.s + ' to be ' + r.NOT_STARTED.s);


            ///////////////////////////////////////////////////////////////////////
            // Very occasionally the On Form Submit trigger does not fire!
            // There are *numerous* continuing reports of this open issue.
            //
            // Now we test that hProcessSubmission can pick up that this and
            // process the previous unprocessed submissions created by a failed
            // submission trigger.
            // 
            // Also we test if hReminderConfirmation can pick up NO_CONTACT
            // results made by hProcessSubmission when it runs out of email quota
            // 
            // We append 4 unprocessed submissions into the spreadsheet and
            // follow that by appending another submission with an associated call
            // to hProcessSubmission with email quota set to 2 over r.MIN_QUOTA.n
            // 
            // We expect hProcessSubmission to pick up the previous unprocessed
            // submission and confirm the first 2 of the 5 in total submissions.
            // 
            // We then make a call to hReminderConfirmation again with email quota
            // set to 2 over r.MIN_QUOTA.n, which we expect to pick up the first 2 of
            // the remaining NO_CONTACT submission
            ///////////////////////////////////////////////////////////////////////
            var numberOfPreviousSubmissions = 10,
                unprocessedEntries = 4;


            overQuota = 2; // we will only be sending 3 fake emails
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota);


            for (i = 0; i < unprocessedEntries + 1; i++) {
                n = ('0' + (i + numberOfPreviousSubmissions)).slice(-2);
                o = {};

                for (j = 0; j < formColumns.length; j++) {
                    column = formColumns[j];
                    o[column] = column + n;

                }
                o[r.TIMESTAMP.s] = new Date(datesData[i] + ' GMT+0000 (GMT)');
                o[r.LENGTH.s] = '3:2';
                o[r.YEAR.s] = (1900 + (10 * (+i))).toString();
                o[r.CONFIRM.s] = 'Confirm';
                spreadsheetValues[i] = columns.map(cellValue);
                formSubmissionValues[i] = formColumnsSpreadsheetOrder.map(submissionValue);
                log('formSubmissionValues[' + i + ']:' + formSubmissionValues[i]);
            }
            var width = spreadsheetValues[0].length,
                submission, row;
            // Do the film submissions
            for (i = 0; i < unprocessedEntries + 1; i++) {
                filmSheet.insertRowAfter((+i) + 1 + numberOfPreviousSubmissions); //just to move the row that the form will populate, one futher on.
                submission = spreadsheetValues[i];
                row = [];
                row[0] = submission;
                var range = filmSheet.getRange(2 + (+i) + numberOfPreviousSubmissions, 1, 1, width);
                // Set the submission data in the spreadsheet as though inserted by an attached form
                range.setValues(row);

                if (i === unprocessedEntries) {
                    //Build the associated Spreadsheet Form Submission object.
                    var spreadsheetFormSubmitEvent = {};
                    spreadsheetFormSubmitEvent.range = range;
                    spreadsheetFormSubmitEvent.values = formSubmissionValues[i];
                    spreadsheetFormSubmitEvent.namedValues = {};
                    for (j = 0; j < submission.length; j++) {
                        if (formColumns.indexOf(columns[j]) > -1 || columns[j] === r.TIMESTAMP.s) {
                            spreadsheetFormSubmitEvent.namedValues[columns[j]] = [submission[j]];
                        }
                    }

                    // Clean up previous tests.
                    deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

                    hProcessSubmission(spreadsheetFormSubmitEvent); //Call the spreadsheet form submission trigget with the event object

                    // Check that the emails we expected to be sent have been sent.
                    templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
                    if (templatesTesting) {
                        templates = templatesTesting.split(',');
                    }
                    ok(templatesTesting && templates.length === 2 && templates[0] === r.SUBMISSION_CONFIRMATION.s && templates[1] === r.SUBMISSION_CONFIRMATION.s, 'Two Submisson Confirmation email has been sent.');
                    deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));

                }
            }

            overQuota = 2;
            pushDate(FakeDate1); //Mock date
            MailApp.setRemainingDailyQuota(r.MIN_QUOTA.n + overQuota); //reset email quota
            hReminderConfirmation(); //6th Reminder Confirmation cycle
            popDate(); //reset date
            templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
            templates = templatesTesting ? templatesTesting.split(',') : null;
            ok(templatesTesting && templates.length === 2 && templates[0] === r.SUBMISSION_CONFIRMATION.s && templates[1] === r.SUBMISSION_CONFIRMATION.s, 'Two Submisson Confirmation email has been sent.');
            deleteProperty(normalizeHeader(r.TEMPLATES_TESTING.s));




        });


    }
} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}
Logger.log('leaving test file');