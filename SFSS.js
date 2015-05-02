/*global DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file sfss');
var sfss = sfss || {};
try {
    sfss.s = (function (s) {
        'use strict';

        var sfss_interface, testing, lg = s.lg,
            u = s.u,
            ui = s.ui,
            i = s.init,
            m = s.merge,
            p = s.process,
            re = s.reminder,
            smm = s.smm,

            normalizeHeader = smm.normalizeHeader,

            log = lg.log,

            normaliseAndValidateDuration = u.normaliseAndValidateDuration,
            pad = u.pad,
            setPadNumber = u.setPadNumber,
            getNamedValue = u.getNamedValue,
            setNamedValue = u.setNamedValue,
            saveData = u.saveData,
            loadData = u.loadData,
            findStatusColor = u.findStatusColor,
            diffDays = u.diffDays,
            findMinMaxColumns = u.findMinMaxColumns,
            setTesting = u.setTesting,

            onOpen = ui.onOpen,

            getProperty = m.getProperty,
            setProperty = m.setProperty,
            deleteProperty = m.deleteProperty,

            setup = i.setup,
            hProcessSubmission = p.hProcessSubmission,
            hReminderConfirmation = re.hReminderConfirmation,

            // spreadsheet menu items
            mediaPresentNotConfirmed = ui.mediaPresentNotConfirmed,
            problem = ui.problem,
            selected = ui.selected,
            notSelected = ui.notSelected,
            manual = ui.manual,
            settingsOptions = ui.settingsOptions,
            editAndSaveTemplates = ui.editAndSaveTemplates,
            adHocEmail = ui.adHocEmail,
            selectionNotification = ui.selectionNotification,

            selectionNotificationAcceptedTemplate = ui.selectionNotificationAcceptedTemplate,
            selectionNotificationNotAcceptedTemplate = ui.selectionNotificationNotAcceptedTemplate,
            adHocEmailTemplate = ui.adHocEmailTemplate,

            // button actions and server handlers
            bbuttonAction = ui.bbuttonAction,
            templatesButtonAction = ui.templatesButtonAction,
            settingsOptionsAction = ui.settingsOptionsAction,
            selectionNotificationButtonAction = ui.selectionNotificationButtonAction,
            adHocEmailButtonAction = ui.adHocEmailButtonAction,
            cancelAction = ui.cancelAction;
            

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
            onOpen: onOpen,

            setup: setup,

            // system triggers
            hProcessSubmission: hProcessSubmission,
            hReminderConfirmation: hReminderConfirmation,

            // spreadsheet menu items
            mediaPresentNotConfirmed: mediaPresentNotConfirmed,
            problem: problem,
            selected: selected,
            notSelected: notSelected,
            manual: manual,
            settingsOptions: settingsOptions,
            editAndSaveTemplates: editAndSaveTemplates,
            adHocEmail: adHocEmail,
            selectionNotification: selectionNotification,

            selectionNotificationAcceptedTemplate: selectionNotificationAcceptedTemplate,
            selectionNotificationNotAcceptedTemplate: selectionNotificationNotAcceptedTemplate,
            adHocEmailTemplate: adHocEmailTemplate,

            // button actions and server handlers
            bbuttonAction: bbuttonAction,
            templatesButtonAction: templatesButtonAction,
            settingsOptionsAction: settingsOptionsAction,
            selectionNotificationButtonAction: selectionNotificationButtonAction,
            adHocEmailButtonAction: adHocEmailButtonAction,
            cancelAction: cancelAction, 
            
            // testing
            testing: testing
        };

        return sfss_interface;
    }(sfss));
} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}

//////////////////////////////////////
// build custom menu for spreadsheet
//////////////////////////////////////
function onOpen() {
    sfss.s.onOpen();
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

////////////////////////////////////////////////////////////////////////////
// spreadsheet menu items
////////////////////////////////////////////////////////////////////////////
function mediaPresentNotConfirmed() {
    sfss.s.mediaPresentNotConfirmed();
}

function problem() {
    sfss.s.problem();
}

function selected() {
    sfss.s.selected();
}

function notSelected() {
    sfss.s.notSelected();
}

function manual() {
    sfss.s.manual();
}

function settingsOptions() {
    sfss.s.settingsOptions();
}

function editAndSaveTemplates() {
    sfss.s.editAndSaveTemplates();
}

function adHocEmail() {
    sfss.s.adHocEmail();
}

function selectionNotification() {
    sfss.s.selectionNotification();
}


function selectionNotificationAcceptedTemplate() {
    sfss.s.selectionNotificationAcceptedTemplate();
}

function selectionNotificationNotAcceptedTemplate() {
    sfss.s.selectionNotificationNotAcceptedTemplate();
}

function adHocEmailTemplate() {
    sfss.s.adHocEmailTemplate();
}

////////////////////////////////////////////////////////////////////////////
// end of spreadsheet menu items
////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////
// start of button actions and server handlers
////////////////////////////////////////////////////////////////////////////
function bbuttonAction(e) {
    return sfss.s.bbuttonAction(e);
}

function templatesButtonAction(e) {
    return sfss.s.templatesButtonAction(e);
}

function settingsOptionsAction(e) {
    return sfss.s.settingsOptionsAction(e);
}

function selectionNotificationButtonAction(e) {
    return sfss.s.selectionNotificationButtonAction(e);
}

function adHocEmailButtonAction(e) {
    return sfss.s.adHocEmailButtonAction(e);
}

function cancelAction(f) {
    return sfss.s.cancelAction(f);
}
////////////////////////////////////////////////////////////////////////////
// end of button actions and server handlers
////////////////////////////////////////////////////////////////////////////
Logger.log('leaving file sfss');