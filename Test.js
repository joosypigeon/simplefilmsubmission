/*global DriveApp, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger, QUnit, sfss, module, test */

///////////////////////////////////////////////////////////////////////////////
// A collection of not very unity Unit Tests
// These tests use a fork of QUnit.
// Google Apps Script library key MxL38OxqIK-B73jyDTvCe-OBao7QLBR4j
// https://github.com/simula-innovation/qunit/tree/gas/gas
///////////////////////////////////////////////////////////////////////////////
function doGet(e) {
    QUnit.urlParams(e.parameter);
    QUnit.config({
        title: "Unit tests for my project"
    });
    QUnit.load(myTests);
    return QUnit.getHtml();
}

// Imports the following functions:
// ok, equal, notEqual, deepEqual, notDeepEqual, strictEqual,
// notStrictEqual, throws, module, test, asyncTest, expect
QUnit.helpers(this);

function myTests() {
    Logger.log('start:myTests');
    var pad = sfss.test_sfss_interface.pad,
        setPadNumber = sfss.test_sfss_interface.setPadNumber,
        diffDays = sfss.test_sfss_interface.diffDays,
        normaliseAndValidateDuration = sfss.test_sfss_interface.normaliseAndValidateDuration,
        normalizeHeader = sfss.test_sfss_interface.normalizeHeader,
        setNamedValue = sfss.test_sfss_interface.setNamedValue,
        getNamedValue = sfss.test_sfss_interface.getNamedValue,
        saveData = sfss.test_sfss_interface.saveData,
        loadData = sfss.test_sfss_interface.loadData,
        findMinMaxColumns = sfss.test_sfss_interface.findMinMaxColumns,
        log = sfss.test_sfss_interface.log,
        findStatusColor = sfss.test_sfss_interface.findStatusColor,
        getProperty = sfss.test_sfss_interface.getProperty,
        setProperty = sfss.test_sfss_interface.setProperty,
        deleteProperty = sfss.test_sfss_interface.deleteProperty,
        setTesting = sfss.test_sfss_interface.setTesting,

        // the 3 spreadsheet sheets
        FILM_SUBMISSIONS_SHEET = sfss.test_sfss_interface.FILM_SUBMISSIONS_SHEET,

        CLOSE_OF_SUBMISSION = sfss.test_sfss_interface.CLOSE_OF_SUBMISSION,
        DAYS_BEFORE_REMINDER = sfss.test_sfss_interface.DAYS_BEFORE_REMINDER,


        SUBMISSION_CONFIRMATION = sfss.test_sfss_interface.SUBMISSION_CONFIRMATION,
        RECEIPT_CONFIMATION = sfss.test_sfss_interface.RECEIPT_CONFIMATION,
        REMINDER = sfss.test_sfss_interface.REMINDER,
        NOT_ACCEPTED = sfss.test_sfss_interface.NOT_ACCEPTED,
        ACCEPTED = sfss.test_sfss_interface.ACCEPTED,
        AD_HOC_EMAIL = sfss.test_sfss_interface.AD_HOC_EMAIL,

        CURRENT_AD_HOC_EMAIL = sfss.test_sfss_interface.CURRENT_AD_HOC_EMAIL,
        CURRENT_SELECTION_NOTIFICATION = sfss.test_sfss_interface.CURRENT_SELECTION_NOTIFICATION,

        //values for CURRENT_SELECTION_NOTIFICATION, CURRENT_AD_HOC_EMAIL
        NOT_STARTED = sfss.test_sfss_interface.NOT_STARTED,
        PENDING = sfss.test_sfss_interface.PENDING,

        //states
        NO_MEDIA = sfss.test_sfss_interface.NO_MEDIA,
        MEDIA_PRESENT = sfss.test_sfss_interface.MEDIA_PRESENT,
        PROBLEM = sfss.test_sfss_interface.PROBLEM,
        SELECTED = sfss.test_sfss_interface.SELECTED,
        NOT_SELECTED = sfss.test_sfss_interface.NOT_SELECTED,
        CONFIRMED = sfss.test_sfss_interface.CONFIRMED,
        NOT_CONFIRMED = sfss.test_sfss_interface.NOT_CONFIRMED,

        LENGTH = sfss.test_sfss_interface.LENGTH,
        YEAR = sfss.test_sfss_interface.YEAR,

        CONFIRM = sfss.test_sfss_interface.CONFIRM,

        TIMESTAMP = sfss.test_sfss_interface.TIMESTAMP,

        //COMMENTS = sfss.test_sfss_interface.COMMENTS,
        FILM_ID = sfss.test_sfss_interface.FILM_ID,
        SCORE = sfss.test_sfss_interface.SCORE,
        CONFIRMATION = sfss.test_sfss_interface.CONFIRMATION,
        SELECTION = sfss.test_sfss_interface.SELECTION,
        STATUS = sfss.test_sfss_interface.STATUS,
        LAST_CONTACT = sfss.test_sfss_interface.LAST_CONTACT,
        FORM_DATA = sfss.test_sfss_interface.FORM_DATA,

        // Sets number of numerical places on film ID so an example of a film ID were the PAD_NUMBER were set to 3 would be ID029.
        // Legal values are 3,4,5,6
        // NOTE: PAD_NUMBER puts an upper bound on the number of film submissions the system can cope with
        PAD_NUMBER = sfss.test_sfss_interface.PAD_NUMBER,

        // Will stop sending emails for the day when email quota is reported as less than or equal to MIN_QUOTA.
        // Will attempt to send unsent emails the next day.
        MIN_QUOTA = sfss.test_sfss_interface.MIN_QUOTA,

        // name of log file
        LOG_FILE = sfss.test_sfss_interface.LOG_FILE,

        TEMPLATES_TESTING = sfss.test_sfss_interface.TEMPLATES_TESTING;

    setTesting(true);

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
            PAD_NUMBER = i + 3;
            setPadNumber(PAD_NUMBER);
            for (var j = 0; j < PAD_NUMBER; j++) {
                strictEqual(pad(padTestDataInput[j]), padTestDataOutput[i][j], 'Require leading zeros on string rep of number with ' + PAD_NUMBER + ' places.');
            }
        }

        var testData = ['\\d\\d\\d', '\\d\\d\\d\\d', '\\d\\d\\d\\d\\d', '\\d\\d\\d\\d\\d\\d'];

        for (i = 0; i < testData.length; i++) {
            PAD_NUMBER = (+i) + 3;
            setPadNumber(PAD_NUMBER);
            strictEqual(pad(-1), testData[i], 'Film ID suffix returned by pad for ' + PAD_NUMBER + ' places.');
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
        // We then set up a mock MailApp which is only overQuota over MIN_QUOTA
        // We simulate 10 submissions
        ////////////////////////////////////////////////////////////////////////
        PAD_NUMBER = 3;
        setPadNumber(PAD_NUMBER);
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
        // than MIN_QUOTA it sends a Submission Confirmation email.
        //
        // NOTE: if no Submission Confirmation email is sent, the submission is marked
        // as No Media, Not Confirmed. When hReminderConfirmation runs (after midnight)
        // Submission Confirmations are sent depending on the daily eamil quota at that
        // point.
        ///////////////////////////////////////////////////////////////////////////////

        // build submission data
        var filmSheet = ss.getSheetByName(FILM_SUBMISSIONS_SHEET),
            datesData = ['Jan 1 2010 01:26:09', 'Jan 2 2010 02:36:10', 'Jan 3 2010 03:10:52', 'Jan 4 2010 04:54:31', 'Jan 5 2010 05:24:15', 'Jan 6 2010 06:57:19', 'Jan 7 2010 07:22:55', 'Jan 8 2010 08:33:22', 'Jan 9 2010 09:42:10', 'Jan 10 2010 10:10:32'],
            columns = filmSheet.getRange(1, 1, 1, filmSheet.getDataRange().getLastColumn()).getValues()[0],
            formColumns = FORM_DATA.filter(function (x) {
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
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota);

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
            o[TIMESTAMP] = new Date(datesData[i] + ' GMT+0000 (GMT)');
            o[LENGTH] = '3:2';
            o[YEAR] = (1900 + (10 * (+i))).toString();
            o[CONFIRM] = 'Confirm';
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
                if (formColumns.indexOf(columns[j]) > -1 || columns[j] === TIMESTAMP) {
                    spreadsheetFormSubmitEvent.namedValues[columns[j]] = [submission[j]];
                }
            }

            // Clean up previous tests.
            deleteProperty(normalizeHeader(TEMPLATES_TESTING));

            hProcessSubmission(spreadsheetFormSubmitEvent); //Call the spreadsheet form submission trigget with the event object
            ////Utilities.sleep(3000);

            // Check that the emails we expected to be sent have been sent.
            templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
            if ((+i) < overQuota) {
                if (templatesTesting) {
                    templates = templatesTesting.split(',');
                }
                ok(templatesTesting && templates.length === 1 && templates[0] === SUBMISSION_CONFIRMATION, 'One Submisson Confirmation email has been sent.');
            } else {
                ok(!templatesTesting, 'No email sent.');
            }
        }

        // After the film submissions, check that the Status, Confirmation, Selection and Film ID columns are as we expect
        var expectedStatus = [
            [NO_MEDIA],
            [NO_MEDIA],
            [NO_MEDIA],
            [NO_MEDIA],
            [NO_MEDIA],
            [NO_MEDIA],
            [NO_MEDIA],
            [NO_MEDIA],
            [NO_MEDIA],
            [NO_MEDIA]
        ];
        var statusColumn = columns.indexOf(STATUS);
        ok(statusColumn > -1, 'There is a ' + STATUS + ' column.');
        statusColumn += 1;
        deepEqual(filmSheet.getRange(2, statusColumn, 10, 1).getValues(), expectedStatus, 'Expect status ' + NO_MEDIA + '.');

        var expectedConfirmation = [
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [NOT_CONFIRMED],
            [NOT_CONFIRMED],
            [NOT_CONFIRMED],
            [NOT_CONFIRMED],
            [NOT_CONFIRMED]
        ];
        var confirmationColumn = columns.indexOf(CONFIRMATION);
        ok(confirmationColumn > -1, 'There is a ' + CONFIRMATION + ' column.');
        confirmationColumn += 1;
        deepEqual(filmSheet.getRange(2, confirmationColumn, 10, 1).getValues(), expectedConfirmation, 'Expect first 5 submissions to be ' + CONFIRMED + ', second 5 submissions to be ' + NOT_CONFIRMED);

        var expectedSelection = [
            [NOT_SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED]
        ];
        var selectionColumn = columns.indexOf(SELECTION);
        ok(selectionColumn > -1, 'There is a ' + SELECTION + ' column.');
        selectionColumn += 1;
        deepEqual(filmSheet.getRange(2, selectionColumn, 10, 1).getValues(), expectedSelection, 'Expect selection to be ' + NOT_SELECTED);

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
        var filmIdColumn = columns.indexOf(FILM_ID);
        ok(filmIdColumn > -1, 'There is a ' + FILM_ID + ' column.');
        filmIdColumn += 1;
        deepEqual(filmSheet.getRange(2, filmIdColumn, 10, 1).getValues(), expectedFilmId, 'Expect correctly formated and sequenced ' + FILM_ID);

        //Utilities.sleep(1000);

        ///////////////////////////////////////////////////////////////////////////////
        // Now we will test hReminderConfirmation which is intended to run once a night
        // after midnight. We will use the 10 submissions already in the spreadsheet.
        // 
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

        // Set the Cose of Submission to something consistant with the Timestamps on the ten film submission
        setNamedValue(ss, CLOSE_OF_SUBMISSION, new Date('Jun 1 2010 GMT+0000 (GMT)'));

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
        // confirmations. That it stops when its daily email quota drops to MIN_QUOTA
        // and that it can resume sending submissions confirmations when called
        // again
        ///////////////////////////////////////////////////////////////////////////////

        // Set the remaining daily email quote to only 3 emails left
        overQuota = 3;
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate1); //Mock date
        hReminderConfirmation(); //first Reminder Confirmation cycle
        popDate(); // reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === overQuota && templates[0] === SUBMISSION_CONFIRMATION && templates[1] === SUBMISSION_CONFIRMATION && templates[2] === SUBMISSION_CONFIRMATION, 'Three ' + SUBMISSION_CONFIRMATION + ' emails sent:' + templates.length);
        for (var ii = 0; ii < templates.length; ii++) {
            log('templates[' + i + ']:' + templates[i]);
        }

        expectedConfirmation = [
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [NOT_CONFIRMED],
            [NOT_CONFIRMED]
        ];
        deepEqual(filmSheet.getRange(2, confirmationColumn, 10, 1).getValues(), expectedConfirmation, 'Expect first 7 submissions to be ' + CONFIRMED + ', second 3 submissions to be ' + NOT_CONFIRMED);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));

        //Utilities.sleep(3000);
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota to only 3 emails left
        pushDate(FakeDate1); //Mock date
        hReminderConfirmation(); //Second Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 2 && templates[0] === SUBMISSION_CONFIRMATION && templates[1] === SUBMISSION_CONFIRMATION, 'Two ' + SUBMISSION_CONFIRMATION + ' emails sent');
        expectedConfirmation = [
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED]
        ];
        deepEqual(filmSheet.getRange(2, confirmationColumn, 10, 1).getValues(), expectedConfirmation, 'Expect all submissions to be ' + CONFIRMED);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));


        function FakeDate2(d) {
            return new NativeDate(d ? d : 'Jan 20 2010 GMT+0000 (GMT)');
        }
        FakeDate2.prototype = Date;

        ///////////////////////////////////////////////////////////////////////////////
        // The next three tests check that hReminderConfirmation can send Reminders
        // when the media is not present for a submissions and a sufficent number of
        // days have passed since the last contact. It is checked that reminders are
        // sent until the daily email quota drops to MIN_QUOTA, in which case it stops.
        // It is checked that hReminderConfirmation can resume sending Reminders when
        // it stopped in the previous call.
        ///////////////////////////////////////////////////////////////////////////////

        // Expect a sufficent amount of time has not elapsed for reminders to
        // have been sent
        overQuota = 3;
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        setNamedValue(ss, DAYS_BEFORE_REMINDER, 50); //High DAYS_BEFORE_REMINDER so no reminder emails need to be sent yet
        pushDate(FakeDate2); //Mock date
        hReminderConfirmation(); //thid Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        ok(!templatesTesting, 'Expect no emails sent i.e. no reminder emails need to be sent yet.');
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));

        //Utilities.sleep(3000);

        // Expect that 3 reminders will be sent.
        setNamedValue(ss, DAYS_BEFORE_REMINDER, 14); // lower DAYS_BEFORE_REMINDER so that reminder emails will need to be sent
        pushDate(FakeDate2); //Mock date
        hReminderConfirmation(); //fourth Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 3 && templates[0] === REMINDER && templates[1] === REMINDER && templates[2] === REMINDER, 'Expect 3 ' + REMINDER + ' emails sent');
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));

        //Utilities.sleep(3000);

        // Expect that 2 reminders will be sent
        overQuota = 3;
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate2); //Mock date
        hReminderConfirmation(); //fourth Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 2 && templates[0] === REMINDER && templates[1] === REMINDER, 'Expect 2 ' + REMINDER + ' emails sent');
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));

        //Utilities.sleep(3000);

        ///////////////////////////////////////////////////////////////////////////////
        // The next 4 tests check that hReminderConfirmation can send Receipt
        // Confirmations. That it stops when its daily email quota drops to MIN_QUOTA
        // and that it can resume sending submissions confirmations when called
        // again
        ///////////////////////////////////////////////////////////////////////////////
        var statusColumnData = [
            [MEDIA_PRESENT],
            [MEDIA_PRESENT],
            [NO_MEDIA],
            [NO_MEDIA],
            [MEDIA_PRESENT],
            [MEDIA_PRESENT],
            [NO_MEDIA],
            [MEDIA_PRESENT],
            [NO_MEDIA],
            [MEDIA_PRESENT]
        ];
        filmSheet.getRange(2, statusColumn, 10, 1).setValues(statusColumnData);
        var confirmationColumnData = [
            [CONFIRMED],
            [NOT_CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [NOT_CONFIRMED],
            [NOT_CONFIRMED],
            [CONFIRMED],
            [NOT_CONFIRMED],
            [CONFIRMED],
            [NOT_CONFIRMED]
        ];
        filmSheet.getRange(2, confirmationColumn, 10, 1).setValues(confirmationColumnData);
        var color;
        for (i = 0; i < confirmationColumnData.length; i++) {
            color = findStatusColor(ss, statusColumnData[i][0], confirmationColumnData[i][0], NOT_SELECTED);
            filmSheet.getRange(2 + (+i), 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);
        }

        // Expect 3 Receipt Confirmations to be sent
        overQuota = 3;
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate2); //Mock date
        hReminderConfirmation(); //fourth Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 3 && templates[0] === RECEIPT_CONFIMATION && templates[1] === RECEIPT_CONFIMATION && templates[2] === RECEIPT_CONFIMATION, 'Expect 3 ' + RECEIPT_CONFIMATION + ' emails sent');
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));

        expectedConfirmation = [
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [NOT_CONFIRMED],
            [CONFIRMED],
            [NOT_CONFIRMED]
        ];
        deepEqual(filmSheet.getRange(2, confirmationColumn, 10, 1).getValues(), expectedConfirmation, 'Expect first 7 submissions to be ' + CONFIRMED + ' and the 9th submission to be ' + CONFIRMED);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));

        // Expect 2 Receipt Confirmations to be sent
        overQuota = 3;
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate2); //Mock date
        hReminderConfirmation(); //fourth Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 2 && templates[0] === RECEIPT_CONFIMATION && templates[1] === RECEIPT_CONFIMATION, 'Expect 2 ' + RECEIPT_CONFIMATION + ' emails sent');
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));

        expectedConfirmation = [
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED],
            [CONFIRMED]
        ];
        deepEqual(filmSheet.getRange(2, confirmationColumn, 10, 1).getValues(), expectedConfirmation, 'Expect all submissions to be ' + CONFIRMED);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));

        //Utilities.sleep(1000);

        ///////////////////////////////////////////////////////////////////////////////
        // Test Ad Hoc Email sent before close of submission
        ///////////////////////////////////////////////////////////////////////////////
        // Set Status of submission ID008 to Problem
        // Submissions with Status Problem are ingnored by Ad Hoc Email
        filmSheet.getRange(8, statusColumn, 1, 1).setValues([
            [PROBLEM]
        ]);
        color = findStatusColor(ss, PROBLEM, CONFIRMED, NOT_SELECTED);
        filmSheet.getRange(8, 1, 1, filmSheet.getDataRange().getLastColumn()).setBackground(color);

        //Set Ad Hoc Email to trigger with the next hReminderConfirmation call
        setNamedValue(ss, CURRENT_AD_HOC_EMAIL, PENDING);

        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate2); //Mock date
        hReminderConfirmation(); //5th Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        log('templates.length:' + templates.length);
        for (var ii = 0; ii < templates.length; ii++) {
            log('templates[' + ii + ']:' + templates[ii]);
        }
        ok(templatesTesting && templates.length === 3 && templates[0] === AD_HOC_EMAIL && templates[1] === AD_HOC_EMAIL && templates[2] === AD_HOC_EMAIL, 'Expect 3 ' + AD_HOC_EMAIL);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));
        equal(getNamedValue(ss, CURRENT_AD_HOC_EMAIL), FILM_ID + pad(4), 'Expect ' + CURRENT_AD_HOC_EMAIL + ' to be ' + FILM_ID + pad(4));

        //Utilities.sleep(1000);
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate2); //Mock date
        hReminderConfirmation(); //6th Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 3 && templates[0] === AD_HOC_EMAIL && templates[1] === AD_HOC_EMAIL && templates[2] === AD_HOC_EMAIL, 'Expect 3 ' + AD_HOC_EMAIL);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));
        equal(getNamedValue(ss, CURRENT_AD_HOC_EMAIL), FILM_ID + pad(7), 'Expect ' + CURRENT_AD_HOC_EMAIL + ' to be ' + FILM_ID + pad(7));

        //Utilities.sleep(1000);
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate2); //Mock date
        hReminderConfirmation(); //6th Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 3 && templates[0] === AD_HOC_EMAIL && templates[1] === AD_HOC_EMAIL && templates[2] === AD_HOC_EMAIL, 'Expect 3 ' + AD_HOC_EMAIL);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));
        equal(getNamedValue(ss, CURRENT_AD_HOC_EMAIL), NOT_STARTED, 'Expect ' + CURRENT_AD_HOC_EMAIL + ' to be ' + NOT_STARTED);

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
        setNamedValue(ss, CURRENT_AD_HOC_EMAIL, PENDING);

        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate3); //Mock date
        hReminderConfirmation(); //5th Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        log('templates.length:' + templates.length);
        for (var ii = 0; ii < templates.length; ii++) {
            log('templates[' + ii + ']:' + templates[ii]);
        }
        ok(templatesTesting && templates.length === 3 && templates[0] === AD_HOC_EMAIL && templates[1] === AD_HOC_EMAIL && templates[2] === AD_HOC_EMAIL, 'Expect 3 ' + AD_HOC_EMAIL);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));
        equal(getNamedValue(ss, CURRENT_AD_HOC_EMAIL), FILM_ID + pad(4), 'Expect ' + CURRENT_AD_HOC_EMAIL + ' to be ' + FILM_ID + pad(4));

        //Utilities.sleep(1000);
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate3); //Mock date
        hReminderConfirmation(); //6th Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 3 && templates[0] === AD_HOC_EMAIL && templates[1] === AD_HOC_EMAIL && templates[2] === AD_HOC_EMAIL, 'Expect 3 ' + AD_HOC_EMAIL);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));
        equal(getNamedValue(ss, CURRENT_AD_HOC_EMAIL), FILM_ID + pad(7), 'Expect ' + CURRENT_AD_HOC_EMAIL + ' to be ' + FILM_ID + pad(7));

        //Utilities.sleep(1000);
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate3); //Mock date
        hReminderConfirmation(); //6th Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 3 && templates[0] === AD_HOC_EMAIL && templates[1] === AD_HOC_EMAIL && templates[2] === AD_HOC_EMAIL, 'Expect 3 ' + AD_HOC_EMAIL);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));
        equal(getNamedValue(ss, CURRENT_AD_HOC_EMAIL), NOT_STARTED, 'Expect ' + CURRENT_AD_HOC_EMAIL + ' to be ' + NOT_STARTED);

        //Utilities.sleep(1000);

        ///////////////////////////////////////////////////////////////////////////////
        // Test Selection Notification Email
        ///////////////////////////////////////////////////////////////////////////////
        selectionColumnData = [
            [SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED],
            [SELECTED],
            [NOT_SELECTED],
            [NOT_SELECTED],
            [SELECTED],
            [NOT_SELECTED],
            [SELECTED]
        ];
        filmSheet.getRange(2, selectionColumn, 10, 1).setValues(selectionColumnData);

        //Set Ad Hoc Email to trigger with the next hReminderConfirmation call
        setNamedValue(ss, CURRENT_SELECTION_NOTIFICATION, PENDING);

        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate3); //Mock date
        hReminderConfirmation(); //5th Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        log('templatesTesting:' + templatesTesting);
        log(templates);
        log('templates.length:' + templates.length);
        ok(templatesTesting && templates.length === 3 && templates[0] === ACCEPTED && templates[1] === NOT_ACCEPTED && templates[2] === NOT_ACCEPTED, 'Expect ' + ACCEPTED + ' followed by 2 ' + NOT_ACCEPTED);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));
        equal(getNamedValue(ss, CURRENT_SELECTION_NOTIFICATION), FILM_ID + pad(4), 'Expect ' + CURRENT_SELECTION_NOTIFICATION + ' to be ' + FILM_ID + pad(4));

        //Utilities.sleep(1000);
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate3); //Mock date
        hReminderConfirmation(); //6th Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 3 && templates[0] === NOT_ACCEPTED && templates[1] === ACCEPTED && templates[2] === NOT_ACCEPTED, 'Expect ' + NOT_ACCEPTED + ' followed by ' + ACCEPTED + ' followed by ' + NOT_ACCEPTED);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));
        equal(getNamedValue(ss, CURRENT_SELECTION_NOTIFICATION), FILM_ID + pad(7), 'Expect ' + CURRENT_SELECTION_NOTIFICATION + ' to be ' + FILM_ID + pad(7));

        //Utilities.sleep(1000);
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        pushDate(FakeDate3); //Mock date
        hReminderConfirmation(); //6th Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 3 && templates[0] === ACCEPTED && templates[1] === NOT_ACCEPTED && templates[2] === ACCEPTED, 'Expect ' + ACCEPTED + ' follwed by ' + NOT_ACCEPTED + ' follwed by ' + ACCEPTED);
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));
        equal(getNamedValue(ss, CURRENT_SELECTION_NOTIFICATION), NOT_STARTED, 'Expect ' + CURRENT_SELECTION_NOTIFICATION + ' to be ' + NOT_STARTED);


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
        // to hProcessSubmission with email quota set to 2 over MIN_QUOTA
        // 
        // We expect hProcessSubmission to pick up the previous unprocessed
        // submission and confirm the first 2 of the 5 in total submissions.
        // 
        // We then make a call to hReminderConfirmation again with email quota
        // set to 2 over MIN_QUOTA, which we expect to pick up the first 2 of
        // the remaining NO_CONTACT submission
        ///////////////////////////////////////////////////////////////////////
        var numberOfPreviousSubmissions = 10,
            unprocessedEntries = 4;


        overQuota = 2; // we will only be sending 3 fake emails
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota);


        for (i = 0; i < unprocessedEntries + 1; i++) {
            n = ('0' + (i + numberOfPreviousSubmissions)).slice(-2);
            o = {};

            for (j = 0; j < formColumns.length; j++) {
                column = formColumns[j];
                o[column] = column + n;

            }
            o[TIMESTAMP] = new Date(datesData[i] + ' GMT+0000 (GMT)');
            o[LENGTH] = '3:2';
            o[YEAR] = (1900 + (10 * (+i))).toString();
            o[CONFIRM] = 'Confirm';
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
                    if (formColumns.indexOf(columns[j]) > -1 || columns[j] === TIMESTAMP) {
                        spreadsheetFormSubmitEvent.namedValues[columns[j]] = [submission[j]];
                    }
                }

                // Clean up previous tests.
                deleteProperty(normalizeHeader(TEMPLATES_TESTING));

                hProcessSubmission(spreadsheetFormSubmitEvent); //Call the spreadsheet form submission trigget with the event object

                // Check that the emails we expected to be sent have been sent.
                templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
                if (templatesTesting) {
                    templates = templatesTesting.split(',');
                }
                ok(templatesTesting && templates.length === 2 && templates[0] === SUBMISSION_CONFIRMATION && templates[1] === SUBMISSION_CONFIRMATION, 'Two Submisson Confirmation email has been sent.');
                deleteProperty(normalizeHeader(TEMPLATES_TESTING));

            }
        }

        overQuota = 2;
        pushDate(FakeDate1); //Mock date
        MailApp.setRemainingDailyQuota(MIN_QUOTA + overQuota); //reset email quota
        hReminderConfirmation(); //6th Reminder Confirmation cycle
        popDate(); //reset date
        templatesTesting = getProperty(normalizeHeader(TEMPLATES_TESTING));
        templates = templatesTesting ? templatesTesting.split(',') : null;
        ok(templatesTesting && templates.length === 2 && templates[0] === SUBMISSION_CONFIRMATION && templates[1] === SUBMISSION_CONFIRMATION, 'Two Submisson Confirmation email has been sent.');
        deleteProperty(normalizeHeader(TEMPLATES_TESTING));
    });

    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteAllProperties();

}