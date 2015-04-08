/*global smm, Session, PropertiesService */
Logger.log('entering file r');
var sfss = sfss || {};
sfss.r = (function () {
    var r = {}; // Resource storage.

    // The resource constructor
    function R(name) {
        this.name = name;
        this.values = {};
    }

    // getters & setters for resource object
    var types = ['n', 'b', 's', 'd'];
    types.forEach(function (t) {
        Object.defineProperty(R.prototype, t, {
            get: function () {
                var value = this.values[t];
                if (value === undefined) {
                    throw {
                        message: 'Attempt to access unset data!',
                        name: this.name,
                        type: t
                    };
                }
                return value;
            },
            set: function (v) {
                this.values[t] = v;
            }
        });
    });

    // class/static method to register a resource
    R.addResource = function (item) {
        var name = item[0],
            value = item[1],
            type = typeof value,
            o = r[name]; // Reasource storage from closure
        if (!o) {
            o = new R(name);
            r[name] = o;
        }

        if (type === 'number') {
            o.n = value;
        } else if (type === 'boolean') {
            o.b = value;
        } else if (type === 'string') {
            o.s = value;
        } else { // data
            o.d = value;
        }
    };

    (function () {
        var resourceData00 = [
            ["FILM_SUBMISSION", "Film Submission"],
            ["ID", "ID"],

            // the 3 spreadsheet sheets
            ["FILM_SUBMISSIONS_SHEET", "Film Submissions"],
            ["TEMPLATE_SHEET", "Templates"],
            ["OPTIONS_SETTINGS_SHEET", "Options & Settings"],

            //named ranges
            ["FESTIVAL_NAME", "Festival Name"],
            ["FESTIVAL_WEBSITE", "Festival Website"],
            ["CLOSE_OF_SUBMISSION", "Close Of Submission"],
            ["EVENT_DATE", "Event Date"],
            ["RELEASE_LINK", "Release Link"],
            ["FIRST_FILM_ID", "First ID"],
            ["CURRENT_FILM_ID", "Current Film ID"],
            ["DAYS_BEFORE_REMINDER", "Days Before Reminder"],
            ["TEST_FIRST_NAME", "Test First Name"],
            ["TEST_LAST_NAME", "Test Last Name"],
            ["TEST_TITLE", "Test Title"],
            ["TEST_EMAIL", "Test Email"],
            ["COLOR_DATA", "Color Data"],
            ["FESTIVAL_DATA", "Festival Data"],
            ["TEMPLATE_DATA", "Template Data"],
            ["TEST_DATA", "Test Data"],
            ["INTERNALS", "Internals"],
            ["SUBMISSION_CONFIRMATION", "Submission Confirmation"],
            ["SUBMISSION_CONFIRMATION_SUBJECT_LINE", "Submission Confirmation Subject Line"],
            ["SUBMISSION_CONFIRMATION_BODY", "Submission Confirmation Body"],
            ["RECEIPT_CONFIMATION", "Receipt Confimation"],
            ["RECEIPT_CONFIMATION_SUBJECT_LINE", "Receipt Confimation Subject Line"],
            ["RECEIPT_CONFIMATION_BODY", "Receipt Confimation Body"],
            ["REMINDER", "Reminder"],
            ["REMINDER_SUBJECT_LINE", "Reminder Subject Line"],
            ["REMINDER_BODY", "Reminder Body"],
            ["NOT_ACCEPTED", "Not Accepted"],
            ["NOT_ACCEPTED_SUBJECT_LINE", "Not Accepted Subject Line"],
            ["NOT_ACCEPTED_BODY", "Not Accepted Body"],
            ["ACCEPTED", "Accepted"],
            ["ACCEPTED_SUBJECT_LINE", "Accepted Subject Line"],
            ["ACCEPTED_BODY", "Accepted Body"],
            ["AD_HOC_EMAIL", "Ad Hoc Email"],
            ["AD_HOC_EMAIL_SUBJECT_LINE", "Ad Hoc Email Subject Line"],
            ["AD_HOC_EMAIL_BODY", "Ad Hoc Email Body"],
            ["SUBJECT_LINE", 'Subject Line'],
            ["ENABLE_REMINDER", "Enable Reminder"],
            ["ENABLE_CONFIRMATION", "Enable Confirmation"],
            ["CURRENT_AD_HOC_EMAIL", "Current Ad Hoc Email"],
            ["CURRENT_SELECTION_NOTIFICATION", "Current Selection Notification"],
            ["SUBJECT", "Subject"],
            ["BODY", "Body"],
            ["AVAILABLE_TAGS", "Available Tags"],
            ["SEND_TEST_EMAIL", "Send Test Email"],

            //sufix and prefix for ID
            ["HELP", "Help"],
            ["LABEL", "Label"],
            ["TEST", "Test"],
            ["BUTTON", "Button"],
            ["HIDDEN", "Hidden"],
            ["PLABEL", "PLable"],

            //IDs
            ["NEXT", "Next"],
            ["PREVIOUS", "Previous"],
            ["SAVE", "Save"],
            ["OK", "OK"],
            ["ENABLE", "Enable"],
            ["UNENABLE", "Unenable"],
            ["CANCEL", "Cancel"],
            ["WAIT", "Wait"],
            ["STATUS_HTML", "Status HTML"],
            ["STATUS_WARNING_HTML", "Status Warning HTML"],
            ["TEST_PROCESSING_LABEL", "Test Processing Label"],
            ["DATE_DIFF", "Date Diff"],
            ["FESTIVAL_DATA_NAMES", "Festival Data Names"],
            ["TEMPLATE_DATA_NAMES", "Template Data Names"],
            ["TEST_DATA_NAMES", "Test Data Names"],
            ["TAGS", "Tags"],

            //value for ENABLE_REMINDER and ENABLE_CONFIRMATION
            ["ENABLED", "Enabled"],
            ["NOT_ENABLED", "Not Enabled"],

            //values for CURRENT_SELECTION_NOTIFICATION, CURRENT_AD_HOC_EMAIL
            ["NOT_STARTED", "Not Started"],
            ["PENDING", "Pending"],

            //states
            ["NO_MEDIA", "No Media"],
            ["MEDIA_PRESENT", "Media Present"],
            ["PROBLEM", "Problem"],
            ["SELECTED", "Selected"],
            ["NOT_SELECTED", "Not Selected"],
            ["CONFIRMED", "Confirmed"],
            ["NOT_CONFIRMED", "Not Confirmed"],

            ["DO_NOT_CHANGE", "Do Not Change"],

            // process names
            ["SELECTION_NOTIFICATION", "Selection Notification"],

            // form data
            ["FIRST_NAME", "First Name"],
            ["LAST_NAME", "Last Name"],
            ["EMAIL", "Email"],
            ["TITLE", "Title"],
            ["LENGTH", "Length"],
            ["COUNTRY", "Country"],
            ["YEAR", "Year"],
            ["GENRE", "Genre"],
            ["WEBSITE", "Website"],
            ["SYNOPSIS", "Synopsis"],
            ["CAST_AND_CREW", "Cast & Crew"],
            ["FESTIVAL_SELECTION_AND_AWARDS", "Festival Selection & Awards"],

            ["CONFIRM", "Confirm"],

            // Timestamp column
            ["TIMESTAMP", "Timestamp"],

            // Additional FILM_SUBMISSIONS_SHEET Columns
            ["COMMENTS", "Comments"],
            ["FILM_ID", "ID"],
            ["SCORE", "Score"],
            ["CONFIRMATION", "Confirmation"],
            ["SELECTION", "Selection"],
            ["STATUS", "Status"],
            ["LAST_CONTACT", "Last Contact"],
            ["FORM_TITLE", "Online Film Submissions Form"],
            ["FORM_RESPONSE", "You now need to download, print off and sign the permission slip. A link to the permission slip has been emailed to you. The permission slip must be mailed together with a DVD of your film, to the address given on the slip. We cannot screen your film without a signed permission slip. If you do not receive this email, please check your spam folder. If you still cannot find the email please email us at " + Session.getActiveUser().getEmail() + " ."],
            //columns for TEMPLATE_SHEET
            ["TEMPLATE_NAME", "Name"],
            ["TEMPLATE", "Template"],

            // Sets number of numerical places on film ID so an example of a film ID were the PAD_NUMBER were set to 3 would be ID029.
            // Legal values are 3,4,5,6
            // NOTE: PAD_NUMBER puts an upper bound on the number of film submissions the system can cope with
            ["PAD_NUMBER", 3],

            // Will stop sending emails for the day when email quota is reported as less than or equal to MIN_QUOTA.
            // Will attempt to send unsent emails the next day.
            ["MIN_QUOTA", 50],
            ["REMAINING_EMAIL_QUOTA", "Remaining Email Quota"],

            // one contact value
            ["NO_CONTACT", "No Contact"],

            // name of log file
            ["LOG_FILE", "Submissions Processing Log File"],
            ["LOG_FOLDER", "Log"],
            ["FORM_FOLDER", "Form"],

            // String resources
            ["PROCESSING", "Processing, please wait . . ."],
            ["SAVE_AND_RETURN_TO", "Save and return To:"],

            // For Tesing
            ["TEMPLATES_TESTING", "Templates Testing"],

            ["INITIALISE", "Initialise"],

            ['currentDate', new Date()],

            ['TESTING', false],
            ['TEMPLATES_TESTING', 'Templates Testing'],
            
            ['SCRIPT_PROPERTIES', PropertiesService.getScriptProperties()]
        ];

        resourceData00.forEach(R.addResource);
    }());

    (function () {
        var resourceData01 = [

            ['MENU_ENTRIES', [{
                name: "Set selected range to: '" + r.MEDIA_PRESENT.s + ", " + r.NOT_CONFIRMED.s + "'",
                functionName: "mediaPresentNotConfirmed"
            }, {
                name: "Set selected range to: '" + r.PROBLEM.s + "'",
                functionName: "problem"
            }, {
                name: "Set selected range to: '" + r.SELECTED.s + "'",
                functionName: "selected"
            }, {
                name: "Set selected range to: '" + r.NOT_SELECTED.s + "'",
                functionName: "notSelected"
            }, {
                name: "Manually set state of selected range",
                functionName: "manual"
            },
            null,
            {
                name: r.OPTIONS_SETTINGS_SHEET.s,
                functionName: "settingsOptions"
            }, {
                name: r.TEMPLATE_SHEET.s,
                functionName: "editAndSaveTemplates"
            },
            null,
            {
                name: r.AD_HOC_EMAIL.s,
                functionName: "adHocEmail"
            }, {
                name: r.SELECTION_NOTIFICATION.s,
                functionName: "selectionNotification"
            }]],


            ['FORM_DATA', [{
                type: "addTextItem",
                required: {},
                title: r.FIRST_NAME.s,
                help: "Please enter your first name"
            }, {
                type: "addTextItem",
                required: {},
                title: r.LAST_NAME.s,
                help: "Please enter your last name"
            }, {
                type: "addTextItem",
                required: {},
                title: r.EMAIL.s,
                help: "Please enter your email address"
            }, "section",
            {
                type: "addTextItem",
                required: {},
                title: r.TITLE.s,
                help: "Please give the title of your film"
            }, {
                type: "addTextItem",
                required: {},
                title: r.LENGTH.s,
                help: "Please enter the length of the film in format MM:SS"
            }, {
                type: "addTextItem",
                required: {},
                title: r.COUNTRY.s,
                help: "Please enter the films country/countries of origin"
            }, {
                type: "addTextItem",
                required: {},
                title: r.YEAR.s,
                help: "Please enter the year of the films completion in format YYYY"
            }, {
                type: "addTextItem",
                required: {},
                title: r.GENRE.s,
                help: "Please enter the genre of the film"
            }, {
                type: "addTextItem",
                title: r.WEBSITE.s,
                help: "Please give film website, if any"
            }, {
                type: "addParagraphTextItem",
                required: {},
                title: r.SYNOPSIS.s,
                help: "Please give synopsis of film"
            }, {
                type: "addParagraphTextItem",
                title: r.CAST_AND_CREW.s,
                help: "Please give principal cast and crew"
            }, {
                type: "addParagraphTextItem",
                title: r.FESTIVAL_SELECTION_AND_AWARDS.s,
                help: "Please give the festivals for which your film has been selected. Also give any awards your film has received"
            }, {
                type: "addCheckboxItem",
                required: {},
                title: r.CONFIRM.s,
                help: "I confirm that I have the right to submit this film and that, should it be accepted, I give permission for this film to be screened at the cambridge strawberry shorts film festival 2015.",
                choice: "Confirm"
            }], 'Form Data'],


            ['ADDITIONAL_COLUMNS', [{
                title: r.FILM_ID.s,
                columnIndex: 1
            }, {
                title: r.SCORE.s,
                columnIndex: 1
            }, {
                title: r.SELECTION.s,
                columnIndex: 1
            }, {
                title: r.CONFIRMATION.s,
                columnIndex: 1
            }, {
                title: r.STATUS.s,
                columnIndex: 1
            }, {
                title: r.LAST_CONTACT.s,
                columnIndex: 1
            }, {
                title: r.COMMENTS.s,
                columnIndex: "end"
            }], 'Additional Columns Data'],

            ['filmSheetColumnWidth', [{
                header: r.TIMESTAMP.s,
                width: 120
            }, {
                header: r.LAST_CONTACT.s,
                width: 120
            }, {
                header: r.STATUS.s,
                width: 93
            }, {
                header: r.CONFIRMATION.s,
                width: 92
            }, {
                header: r.SELECTION.s,
                width: 85
            }, {
                header: r.SCORE.s,
                width: 48
            }, {
                header: r.FILM_ID.s,
                width: 40
            }, {
                header: r.FIRST_NAME.s,
                width: 120
            }, {
                header: r.LAST_NAME.s,
                width: 120
            }, {
                header: r.EMAIL.s,
                width: 52
            }, {
                header: r.TITLE.s,
                width: 166
            }, {
                header: r.LENGTH.s,
                width: 52
            }, {
                header: r.COUNTRY.s,
                width: 120
            }, {
                header: r.YEAR.s,
                width: 37
            }, {
                header: r.GENRE.s,
                width: 120
            }, {
                header: r.WEBSITE.s,
                width: 120
            }, {
                header: r.SYNOPSIS.s,
                width: 778
            }, {
                header: r.CAST_AND_CREW.s,
                width: 463
            }, {
                header: r.FESTIVAL_SELECTION_AND_AWARDS.s,
                width: 503
            }, {
                header: r.CONFIRM.s,
                width: 120
            }, {
                header: r.COMMENTS.s,
                width: 120
            }]],

            ['templateSheetColumnWidth', [{
                header: r.TEMPLATE_NAME.s,
                width: 235
            }, {
                header: r.TEMPLATE.s,
                width: 410
            }]],

            // data for Options & Settings sheet
            ['STATES', [
                [r.NO_MEDIA.s + " " + r.NOT_CONFIRMED.s, [r.NO_MEDIA.s + " " + r.NOT_CONFIRMED.s, "IndianRed"]],
                [r.NO_MEDIA.s + " " + r.CONFIRMED.s, [r.NO_MEDIA.s + " " + r.CONFIRMED.s, "LightPink"]],
                [r.MEDIA_PRESENT.s + " " + r.NOT_CONFIRMED.s, [r.MEDIA_PRESENT.s + " " + r.NOT_CONFIRMED.s, "DarkSeaGreen"]],
                [r.MEDIA_PRESENT.s + " " + r.CONFIRMED.s, [r.MEDIA_PRESENT.s + " " + r.CONFIRMED.s, "LightGreen"]],
                [r.PROBLEM.s, [r.PROBLEM.s, "Orange"]],
                [r.SELECTED.s, [r.SELECTED.s, "Gold"]],
                [r.NOT_SELECTED.s, [r.NOT_SELECTED.s, "LightSteelBlue"]]
            ]],

            ['closeDate', new Date(r.currentDate.d.getFullYear(), r.currentDate.d.getMonth() + 7, r.currentDate.d.getDate())],
            ['eventDate', new Date(r.currentDate.d.getFullYear(), r.currentDate.d.getMonth() + 9, r.currentDate.d.getDate())],

            ['TEST_DATA', [
                [r.TEST_FIRST_NAME.s, [r.TEST_FIRST_NAME.s, 'Maya']],
                [r.TEST_LAST_NAME.s, [r.TEST_LAST_NAME.s, 'Deren']],
                [r.TEST_TITLE.s, [r.TEST_TITLE.s, 'Meshes of the Afternoon']],
                [r.TEST_EMAIL.s, [r.TEST_EMAIL.s, Session.getActiveUser().getEmail()]]
            ]],

            ['INTERNALS', [
                [r.FIRST_FILM_ID.s, [r.FIRST_FILM_ID.s, 2]],
                [r.CURRENT_FILM_ID.s, [r.CURRENT_FILM_ID.s, -1]],
                [r.CURRENT_SELECTION_NOTIFICATION.s, [r.CURRENT_SELECTION_NOTIFICATION.s, r.NOT_STARTED.s]],
                [r.CURRENT_AD_HOC_EMAIL.s, [r.CURRENT_AD_HOC_EMAIL.s, r.NOT_STARTED.s]]
            ]],

            // data for Template sheet
            ['TEMPLATE_DATA', [
                [r.TEMPLATE_NAME.s, r.TEMPLATE.s],
                [r.SUBMISSION_CONFIRMATION_SUBJECT_LINE.s, [r.SUBMISSION_CONFIRMATION_SUBJECT_LINE.s, '${"Festival Name"} Submissions Confirmation: ${"Title"}']],
                [r.SUBMISSION_CONFIRMATION_BODY.s, [r.SUBMISSION_CONFIRMATION_BODY.s, 'Dear ${"First Name"},\n\nthank you for submitting your film "${"Title"}" to ${"Festival Name"}.\n\nNow you need to download and sign the permission slip that can be found on our website via the link below.\n${"Release Link"}\n\nThe signed permission slip must be mailed together with a DVD of your film to our festival office. If we do not receive these by the close of submission on ${"Close Of Submission"}, we will not be able to consider your film for selection. The address of our festival office is given on the permission slip.\n\nYou will receive an email confirmation of the receipt of your film and after the close of submissions you will be advised by emailed whether or not your film has been selected.\n\nMany Thanks\n${"Festival Name"}\n${"Festival Website"}']],
                [r.RECEIPT_CONFIMATION_SUBJECT_LINE.s, [r.RECEIPT_CONFIMATION_SUBJECT_LINE.s, '${"Festival Name"} Receipt Confimation: ${"Title"}']],
                [r.RECEIPT_CONFIMATION_BODY.s, [r.RECEIPT_CONFIMATION_BODY.s, 'Dear ${"First Name"}\n\nthis is to confirm the receipt of the DVD of your film "${"Title"}" by ${"Festival Name"}.\n\nSelection will take place after the close of submission on ${"Close Of Submission"}. You will be informed after this date if your film has been selected.\n\n${"Festival Name"} runs on the evening of ${"Event Date"}.\n\nMany thanks for your film submission.\n${"Festival Name"}\n${"Festival Website"}']],
                [r.REMINDER_SUBJECT_LINE.s, [r.REMINDER_SUBJECT_LINE.s, '${"Festival Name"} Submission Reminder: ${"Title"}']],
                [r.REMINDER_BODY.s, [r.REMINDER_BODY.s, 'Dear ${"First Name"}\n\nthank you for submitting your film "${"Title"}" to ${"Festival Name"}.\n\nWe are still awaiting a DVD of your film together with a signed permission slip. If we do not receive these at our festival office by the close of submission on ${"Close Of Submission"}, we will not be able to consider your film for selection.\n\nYou can download the permission slip from our website via the link below.\n${"Release Link"}\nThe address of our festival office is given on the permission slip.\n\nWe will confirm the receipt of your film by email and after the close of submissions you will be advised by emailed whether or not your film has been selected.\n\n${"Festival Name"} runs on the evening of ${"Event Date"}.\n\nMany Thanks\n${"Festival Name"}\n${"Festival Website"}']],
                [r.NOT_ACCEPTED_SUBJECT_LINE.s, [r.NOT_ACCEPTED_SUBJECT_LINE.s, '${"Festival Name"}: ${"Title"}']],
                [r.NOT_ACCEPTED_BODY.s, [r.NOT_ACCEPTED_BODY.s, 'Dear ${"First Name"},\n\nit has been very difficult to make the selection for ${"Festival Name"} this year. We have received an enormous number of accomplished short films from all over the world and we have had to make some difficult choices. We are sorry to tell you that your film ${"Title"} has not been selected for ${"Festival Name"}.\n\nIt is extremely important to point out that this years film selection is intended to be a broad representative sample from the huge number of submissions we have received. Many film merited being in the selection, but could not be include due to the limited screening time of the festival.\n\nWe wish to thank you for your film submission and to encourage you to submit to ${"Festival Name"} next year.\n\nMany Thanks\n${"Festival Name"}\n${"Festival Website"}']],
                [r.ACCEPTED_SUBJECT_LINE.s, [r.ACCEPTED_SUBJECT_LINE.s, '${"Festival Name"}: ${"Title"}']],
                [r.ACCEPTED_BODY.s, [r.ACCEPTED_BODY.s, 'Dear ${"First Name"},\n\nit has been very difficult to make the selection for ${"Festival Name"} this year. We have received an enormous number of accomplished short films from all over the world and we have had to make some difficult choices. We are very happy to tell you that your film ${"Title"} has been selected for ${"Festival Name"}.\n\nDue to limit seating at our festival venue, we can only give two free tickets to each attending filmmaker.\n\nWe will be in touch shortly to confirm the details of the screening.\n\nMany Thanks\n${"Festival Name"}\n${"Festival Website"}']],
                [r.AD_HOC_EMAIL_SUBJECT_LINE.s, [r.AD_HOC_EMAIL_SUBJECT_LINE.s, '${"Festival Name"}: ${"Title"}']],
                [r.AD_HOC_EMAIL_BODY.s, [r.AD_HOC_EMAIL_BODY.s, 'Dear ${"First Name"},\n\njust to thank you for submitting your film "${"Title"}" to ${"Festival Name"}.\n\nAnd to let you know that ${"Festival Name"} is still completely awesome.\n\nMany Thanks\n${"Festival Name"}\n${"Festival Website"}']]
            ]]
        ];


        resourceData01.forEach(R.addResource);
    }());

    (function () {
        var resourceData02 = [
            ['FESTIVAL_DATA', [
                [r.FESTIVAL_NAME.s, [r.FESTIVAL_NAME.s, 'The Awesome Short Film Festival']],
                [r.FESTIVAL_WEBSITE.s, [r.FESTIVAL_WEBSITE.s, 'http://festivalwebsite']],
                [r.CLOSE_OF_SUBMISSION.s, [r.CLOSE_OF_SUBMISSION.s, r.closeDate.d]],
                [r.RELEASE_LINK.s, [r.RELEASE_LINK.s, 'http://festivalwebsite/releaseform']],
                [r.EVENT_DATE.s, [r.EVENT_DATE.s, r.eventDate.d]],
                [r.DAYS_BEFORE_REMINDER.s, [r.DAYS_BEFORE_REMINDER.s, 28]],
                [r.ENABLE_REMINDER.s, [r.ENABLE_REMINDER.s, r.ENABLED.s]],
                [r.ENABLE_CONFIRMATION.s, [r.ENABLE_CONFIRMATION.s, r.ENABLED.s]]
            ]]
        ];
        resourceData02.forEach(R.addResource);
    }());

    (function () {
        var resourceData03 = [
            ['OPTION_SHEET_DATA', [{
                name: r.FESTIVAL_DATA.s,
                data: r.FESTIVAL_DATA.d
            }, {
                name: r.TEST_DATA.s,
                data: r.TEST_DATA.d
            }, {
                name: r.COLOR_DATA.s,
                data: r.STATES.d
            }, {
                name: r.INTERNALS.s,
                data: r.INTERNALS.d
            }]]
        ];
        resourceData03.forEach(R.addResource);
    }());

    Object.freeze(r);

    return r; // Pass out resource storage
}());
Logger.log('leaving file r');