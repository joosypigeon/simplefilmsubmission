/*global sfss, DocumentApp, ScriptApp, FormApp, DriveApp, Session, LockService, PropertiesService, MailApp, UiApp, SpreadsheetApp, Logger */
Logger.log('entering file form');
var sfss = sfss || {};
try {
    sfss.form = (function (r, lg, u, smm) {
        'use strict';
        var form_interface, log = lg.log,
            catchToString = lg.catchToString,
            saveData = u.saveData;

        function setColumnWidth(ss, sheetName, columnWidths) {
            try {
                var sheet = ss.getSheetByName(sheetName),
                    headers = sheet.getRange(1, 1, 1, sheet.getDataRange().getLastColumn()).getValues()[0],
                    width = {};
                for (var columnIndex = 0; columnIndex < columnWidths.length; columnIndex++) {
                    width[columnWidths[columnIndex].header] = columnWidths[columnIndex].width;
                }

                for (columnIndex in headers) {
                    if (width[headers[columnIndex]]) {
                        sheet.setColumnWidth(1 + (+columnIndex), width[headers[columnIndex]]);
                    }
                }
                SpreadsheetApp.flush();
            } catch (e) {
                log('setColumnWidth:error:' + catchToString(e));
            }
        }

        function buildFormAndSpreadsheet(ss) {
            var itemData, sheet;

            // build online submission form
            var form = FormApp.create(r.FORM_TITLE.s).setConfirmationMessage(r.FORM_RESPONSE.s);
            for (var itemIndex = 0; itemIndex < r.FORM_DATA.d.length; itemIndex++) {
                itemData = r.FORM_DATA.d[itemIndex];
                if (itemData === "section") {
                    form.addSectionHeaderItem();
                } else {
                    var formItem = form[itemData.type]().setTitle(itemData.title).setHelpText(itemData.help).setRequired(itemData.required ? true : false);
                    if (itemData.type === "addCheckboxItem") {
                        formItem.setChoices([formItem.createChoice(itemData.choice)]);
                    }
                    if (itemData.type === "addMultipleChoiceItem") {
                        formItem.setChoiceValues(itemData.choices);
                    }
                }
            }

            var ssId = ss.getId(),
                formId = form.getId();

            // find parent folder of spreadsheet and move form to that folder
            var ssParentFolder = DriveApp.getFileById(ssId).getParents().next(),
                formFile = DriveApp.getFileById(formId),
                formFolder = DriveApp.createFolder(r.FORM_FOLDER.s);

            DriveApp.removeFolder(formFolder); // remove from root
            DriveApp.removeFile(formFile); // remove from root
            ssParentFolder.addFolder(formFolder);
            formFolder.addFile(formFile);

            log('Built online submission form.');

            // initialise spreadsheet and set spreadsheet as the destination of the form
            form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

            // Need this for testing.
            // Otherwise 'Form Responses 1' not present when looked for!!!
            // Why does this only break testing!!!
            SpreadsheetApp.flush();

            // create the r.FILM_SUBMISSIONS_SHEET.s, r.OPTIONS_SETTINGS_SHEET.s and r.TEMPLATE_SHEET.s sheet
            ss.getSheetByName('Form Responses 1').setName(r.FILM_SUBMISSIONS_SHEET.s);
            ss.insertSheet(r.OPTIONS_SETTINGS_SHEET.s, 2);
            ss.insertSheet(r.TEMPLATE_SHEET.s, 3);
            ss.deleteSheet(ss.getSheetByName('Sheet1'));

            // add extra columns to r.FILM_SUBMISSIONS_SHEET.s
            sheet = ss.getSheetByName(r.FILM_SUBMISSIONS_SHEET.s);
            for (var i = 0; i < r.ADDITIONAL_COLUMNS.d.length; i++) {
                columnIndex = r.ADDITIONAL_COLUMNS.d[i].columnIndex;
                if (columnIndex === 'end') {
                    columnIndex = sheet.getLastColumn();
                }
                sheet.insertColumnAfter(columnIndex);
                sheet.getRange(1, columnIndex + 1, 1, 1).setValue(r.ADDITIONAL_COLUMNS.d[i].title);
            }

            //add templates to r.TEMPLATE_SHEET.s
            saveData(ss, ss.getSheetByName(r.TEMPLATE_SHEET.s).getRange(1, 1, 1, 1), r.TEMPLATE_DATA.d, r.TEMPLATE_DATA.s, 'mistyrose');

            //initialise r.OPTIONS_SETTINGS_SHEET.s
            for (i = 0; i < r.OPTION_SHEET_DATA.d.length; i++) {
                var dataItem = r.OPTION_SHEET_DATA.d[i];
                saveData(ss, ss.getSheetByName(r.OPTIONS_SETTINGS_SHEET.s).getRange(1, 2 * (+i) + 1, 1, 1), dataItem.data, dataItem.name, (+i) % 2 ? 'aliceblue' : 'mistyrose');

            }

            setColumnWidth(ss, r.FILM_SUBMISSIONS_SHEET.s, r.filmSheetColumnWidth.d);
            setColumnWidth(ss, r.TEMPLATE_SHEET.s, r.templateSheetColumnWidth.d);

            log('Built spreadsheet.');
        }

        form_interface = {
            buildFormAndSpreadsheet: buildFormAndSpreadsheet
        };

        return form_interface;
    }(sfss.r, sfss.lg, sfss.u, sfss.smm));
} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}
Logger.log('leaving file form');