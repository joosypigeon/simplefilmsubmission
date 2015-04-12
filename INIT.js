/*global sfss, DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file init');
var sfss = sfss || {};
try {
    sfss.init = (function (r, lg, u, f, smm) {
        'use strict';
        var init_interface, log = lg.log,
            catchToString = lg.catchToString,
            buildFormAndSpreadsheet = f.buildFormAndSpreadsheet,
            buildLogFile = lg.buildLogFile,
            settingsOptions = f.settingsOptions;

        function setup() {
            try {
                if (r.TESTING.b || !r.SCRIPT_PROPERTIES.d.getProperty("initialised")) {
                    if (!r.TESTING.b) {
                        r.SCRIPT_PROPERTIES.d.setProperty("initialised", "initialised");
                    }

                    log("setup start.");
                    var ss = SpreadsheetApp.getActiveSpreadsheet();
                    if (!r.TESTING.b) {
                        buildLogFile(ss);
                    }
                    buildFormAndSpreadsheet(ss);

                    // update menu
                    if (!r.TESTING.b) {
                        ss.updateMenu(r.FILM_SUBMISSION.s, r.MENU_ENTRIES.d);

                        settingsOptions(); // let user set options and settings

                        // enable film submission processing on form submission
                        ScriptApp.newTrigger("hProcessSubmission").forSpreadsheet(ss).onFormSubmit().create();

                        // enable reminders and confirmations, sent out after midnight each day
                        ScriptApp.newTrigger("hReminderConfirmation").timeBased().everyDays(1).atHour(3).create();

                        log('Added triggers.');
                        log("setup end.");
                    }
                }
            } catch (e) {
                Logger.log('setup:' + catchToString(e));
            }
        }

        init_interface = {
            setup: setup
        };

        return init_interface;
    }(sfss.r, sfss.lg, sfss.u, sfss.form, sfss.smm));
} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}
Logger.log('leaving file init');