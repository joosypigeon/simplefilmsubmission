/*global DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file log');
var sfss = sfss || {};
sfss.lg = (function (r, smm) {
    'use strict';

    var log_interface = {},
        normalizeHeader = smm.normalizeHeader,
        logFile, logDoc, logId = r.SCRIPT_PROPERTIES.d.getProperty(normalizeHeader(r.LOG_FILE.s));


    if (logId) {
        logDoc = DocumentApp.openById(logId);
    }

    function setLogDoc(lD) {
        logDoc = lD;
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
            if (!logDoc) {
                logDoc = DocumentApp.create(r.LOG_FILE.s);
                logId = logDoc.getId();
                logFile = DriveApp.getFileById(logId);
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

                r.SCRIPT_PROPERTIES.d.setProperty(normalizeHeader(r.LOG_FILE.s), logId);

            }
            logDoc.getBody().appendParagraph((new Date()) + ": " + info);
        } catch (e) {
            Logger.log('log:error:' + catchToString(e));
        }
    }



    log_interface = {
        log: log,
        catchToString: catchToString,
        setLogDoc: setLogDoc
    };

    return log_interface;

}(sfss.r, sfss.smm));
Logger.log('leaving file log');