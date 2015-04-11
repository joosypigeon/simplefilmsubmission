/*global DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file merge');
var sfss = sfss || {};
try {
    sfss.merge = (function (r, lg, u, smm) {
        'use strict';


        var merge_interface, PROPERTIES = {},
            log = lg.log,
            catchToString = lg.catchToString,
           
            normalizeHeader = smm.normalizeHeader,
            fillInTemplateFromObject = smm.fillInTemplateFromObject,
            
            getNamedValue = u.getNamedValue,
            prettyPrintDate = u.prettyPrintDate;

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

        function mergeAndSend(ss, submission, templateName, emailQuotaRemaining) {
            try {
                emailQuotaRemaining = emailQuotaRemaining ? emailQuotaRemaining : 0; // Set to 0 if not supplied i.e. it dosnt matter in this case
                submission[normalizeHeader(r.FESTIVAL_NAME.s)] = getNamedValue(ss, r.FESTIVAL_NAME.s);
                submission[normalizeHeader(r.CLOSE_OF_SUBMISSION.s)] = prettyPrintDate(getNamedValue(ss, r.CLOSE_OF_SUBMISSION.s));
                submission[normalizeHeader(r.EVENT_DATE.s)] = prettyPrintDate(getNamedValue(ss, r.EVENT_DATE.s));
                submission[normalizeHeader(r.FESTIVAL_WEBSITE.s)] = getNamedValue(ss, r.FESTIVAL_WEBSITE.s);
                submission[normalizeHeader(r.RELEASE_LINK.s)] = getNamedValue(ss, r.RELEASE_LINK.s);

                var subjectTemplate = getNamedValue(ss, templateName + ' ' + r.SUBJECT_LINE.s),
                    bodyTemplate = getNamedValue(ss, templateName + ' ' + r.BODY.s);

                if (submission[normalizeHeader(r.EMAIL.s)]) {
                    MailApp.sendEmail(submission[normalizeHeader(r.EMAIL.s)], fillInTemplateFromObject(subjectTemplate, submission), fillInTemplateFromObject(bodyTemplate, submission));
                    emailQuotaRemaining -= 1;
                    if (r.TESTING.b) {
                        var templatesTesting = getProperty(normalizeHeader(r.TEMPLATES_TESTING.s));
                        if (templatesTesting) {
                            templatesTesting += ',' + templateName;
                        } else {
                            templatesTesting = templateName;
                        }
                        setProperty(normalizeHeader(r.TEMPLATES_TESTING.s), templatesTesting);
                    }
                    log('mergeAndSend:email sent');
                } else {
                    log('mergeAndSend:error:no email found!');
                }


            } catch (e) {
                log('mergeAndSend:' + catchToString(e));
            }
            return emailQuotaRemaining;
        }


        merge_interface = {
            mergeAndSend: mergeAndSend,
            getProperty: getProperty,
            setProperty: setProperty,
            deleteProperty: deleteProperty
        };

        return merge_interface;

    }(sfss.r, sfss.lg, sfss.u, sfss.smm));

} catch (e) {
    Logger.log(e);
}
Logger.log('leaving file merge');