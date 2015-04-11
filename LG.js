/*global DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file log');
var sfss = sfss || {};
try {
    sfss.lg = (function (r, smm) {
        'use strict';

        var log_interface, normalizeHeader = smm.normalizeHeader,
            logFile, logDoc, logId = r.SCRIPT_PROPERTIES.d.getProperty(normalizeHeader(r.LOG_FILE.s)),
            logDoc;

        if (logId && !r.TESTING.b) {
            logDoc = DocumentApp.openById(logId);
        }

        function buildLogFile(ss) {
            logDoc = DocumentApp.create(r.LOG_FILE.s);
            logId = logDoc.getId();
            logFile = DriveApp.getFileById(logId);
            DriveApp.removeFile(logFile); // remove from root
            var ssParentFolder = DriveApp.getFileById(ss.getId()).getParents().next(),
                logFolder = DriveApp.createFolder(r.LOG_FOLDER.s);
            Logger.log('ssParentFolder: ' + ssParentFolder);
            Logger.log('logFolder: ' + logFolder);
            DriveApp.removeFolder(logFolder); // remove from root
            ssParentFolder.addFolder(logFolder);
            logFolder.addFile(logFile);

            logDoc.getBody().appendParagraph((new Date()) + ":Created log file!");

            if (!r.TESTING.b) {
                r.SCRIPT_PROPERTIES.d.setProperty(normalizeHeader(r.LOG_FILE.s), logId);
            }
        }

        function catchToString(err) {
            var errInfo = "ERROR:\n";
            for (var prop in err) {
                if (err.hasOwnProperty(prop)) {
                    errInfo += "(" + prop + ", " + err[prop] + ")\n";
                }
            }
            errInfo += "  toString(): " + " value: [" + err.toString() + "]";
            return errInfo;
        }

        function log(info) {
            try {
                logDoc.getBody().appendParagraph((new Date()) + ": " + info);
            } catch (e) {
                Logger.log('log:' + catchToString(e));
            }
        }



        log_interface = {
            log: log,
            catchToString: catchToString,
            buildLogFile: buildLogFile
        };

        return log_interface;

    }(sfss.r, sfss.smm));

} catch (e) {
    Logger.log(e);
}
Logger.log('leaving file log');