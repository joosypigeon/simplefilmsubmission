/*global Logger, SpreadsheetApp, MailApp */

Logger.log('entering file util');
var sfss = sfss || {};
try {
    sfss.u = (function (r, lg, smm) {
        'use strict';

        var utils_interface,

            CACHE = {},

            log = lg.log,
            catchToString = lg.catchToString,
            
            fillInTemplateFromObject = smm.fillInTemplateFromObject,
            normalizeHeaders = smm.normalizeHeaders,
            normalizeHeader = smm.normalizeHeader;




        function prettyPrintDate(date) {
            if (!date) {
                return null;
            } else {
                return ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][date.getDay()] + " the " + (function (d) {
                    var s = d.toString(),
                        l = s[s.length - 1];
                    return s + (['th', 'th', 'th'][d - 11] || ['st', 'nd', 'rd'][l - 1] || 'th');
                })(date.getDate()) + " of " + ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][date.getMonth()] + ", " + date.getFullYear();
            }
        }

        function normaliseAndValidateDuration(length) {
            log('normaliseAndValidateDuration:start:length:"' + length + '"');

            function babylonian(myNumber) {
                return 0 <= myNumber && myNumber <= 59;
            }
            var returnLength, mymatch;
            try {
                var hours = '0';
                minutes = '0', seconds = '0', myregexp = /^[^\d]*([\d]+)[^\d]*$|^[^\d]*([\d]+)[^\d]+([\d]+)[^\d]*$|^[^\d]*([\d]+)[^\d]+([\d]+)[^\d]+([\d]+)[^\d]*$/i;

                mymatch = myregexp.exec(length);
                if (mymatch) {
                    if (mymatch[1]) {
                        // One number found, assume it is minutes
                        minutes = mymatch[1];
                    } else if (mymatch[2]) {
                        // NOTE: due to regexp mymatch[2] is true ifandonlyif mymatch[3] is true
                        // Found two numbers, assume they are minutes and seconds
                        minutes = mymatch[2];
                        seconds = mymatch[3];
                    } else if (mymatch[4]) {
                        // NOTE: due to regexp mymatch[4] is true ifandonlyif mymatch[5] is true ifandonlyif mymatch[6] is true
                        // Found three numbers, assume they are hours, minutes and seconds
                        hours = mymatch[4];
                        minutes = mymatch[5];
                        seconds = mymatch[6];
                    }

                    if (babylonian(hours) && babylonian(minutes) && babylonian(seconds)) {
                        returnLength = ('0' + hours).slice(-2) + ':' + ('0' + minutes).slice(-2) + ':' + ('0' + seconds).slice(-2);
                    }
                }
            } catch (e) {
                log('normaliseAndValidateDuration:error:' + catchToString(e));
            }
            log('normaliseAndValidateDuration:end');
            return returnLength ? returnLength : length;
        }

        function pad(number) {
            if (number > 0) {
                return ("000000" + number).slice(-r.PAD_NUMBER.n);
            } else {
                return ("\\d\\d\\d\\d\\d\\d").slice(-2 * r.PAD_NUMBER.n);
            }
        }

        // for testing
        function setPadNumber(n) {
            r.PAD_NUMBER.n = n;
        }

        function setNamedValueInCache(name, value) {
            CACHE[name] = value;
        }

        function getNamedValueFromCache(name) {
            return CACHE[name];
        }

        function getNamedValue(ss, name) {
            var value;
            try {
                name = normalizeHeader(name);

                value = getNamedValueFromCache(name);

                if (!value) {
                    value = ss.getRangeByName(name).getValue();
                    setNamedValueInCache(name, value);
                }
            } catch (e) {
                log('getNamedValue:name:' + name + ':error:' + catchToString(e));
            }
            return value;
        }

        //NOTE: possible to set a value for a named range that does not exist.
        //It will be set in the cache just for this single run of the script.
        function setNamedValue(ss, name, value) {
            try {
                name = normalizeHeader(name);
                var range = null;

                // getRangeByName should not fail by throwing an exception when a named
                // range does not exsist, but it does!!!
                try {
                    range = ss.getRangeByName(name);
                } catch (e) {
                    log('setNamedValue:probably not important:error:' + catchToString(e));
                }

                if (range) {
                    range.setValue(value);
                } else {
                    log('setNamedValue: no range with name:' + name + ' found!');
                }
                setNamedValueInCache(name, value);
            } catch (e) {
                log('setNamedValue:name:' + name + ':error:' + catchToString(e));
            }
        }

        // Load data into CACHE so that getNamedRange will later get it from CACHE and
        // not have to make an individual getValue on a spreadsheet cell
        function loadData(ss, name) {
            var items = [];
            try {
                log('loadData:loading data block:' + name);
                name = normalizeHeader(name);
                var loaded = getNamedValueFromCache(name);
                if (!loaded) {
                    var item;

                    // Named range assumed to have width of 2.
                    // Each row is a name value pair.
                    var values = ss.getRangeByName(name).getValues();
                    for (var i = 0; i < values.length; i++) {
                        item = values[i];
                        items.push(item[0]);
                        CACHE[normalizeHeader(item[0])] = item[1];
                    }
                    setNamedValueInCache(name, true);
                } else {
                    log('loadData:date block has already been loaded!');
                }
            } catch (e) {
                log('loadData:error:' + catchToString(e));
            }
            return items;
        }

        // sets a range with array values
        // if an entry in values is a pair, the first member will be the name of a cell and the second member will be its value
        function saveData(ss, startCell, values, name, color) {
            try {
                var rowIndex, columnIndex, rowValue, value, currentCell;

                var range = startCell.offset(0, 0, values.length, 2);
                range.setBackground(color);
                ss.setNamedRange(normalizeHeader(name), range);

                for (rowIndex = 0; rowIndex < values.length; rowIndex++) {
                    rowValue = values[rowIndex];
                    for (columnIndex = 0; columnIndex < rowValue.length; columnIndex++) {
                        value = rowValue[columnIndex];
                        currentCell = startCell.offset(rowIndex, columnIndex);
                        if (Array.isArray(value)) { //expect array to have 2 elements
                            ss.setNamedRange(normalizeHeader(value[0]), currentCell);
                            value = value[1];
                        }
                        currentCell.setValue(value);
                    }
                }
            } catch (e) {
                log('saveData:error:' + catchToString(e));
            }
        }


        function findStatusColor(ss, status, confirmation, selection) {
            var color = r.NOT_SELECTED.s;

            if (selection === r.SELECTED.s) {
                color = r.SELECTED.s;
            } else if (status === r.PROBLEM.s) {
                color = r.PROBLEM.s;
            } else {
                color = status + ' ' + confirmation;
            }
            return getNamedValue(ss, color);
        }

        //returns the difference in days of two dates
        function diffDays(a, b) {
            return Math.floor((a.getTime() - b.getTime()) / (1000 * 3600 * 24));
        }


        function findMinMaxColumns(filmSheet, headersOfInterest) {
            var headersRange = filmSheet.getRange(1, 1, 1, filmSheet.getDataRange().getLastColumn()),
                headers = normalizeHeaders(headersRange.getValues()[0]),
                headersOfInterestIndices = [],
                minColumn, maxColumn, indices = {};

            headersOfInterest = normalizeHeaders(headersOfInterest);

            for (var i = 0; i < headersOfInterest.length; i++) {
                var header = headersOfInterest[i],
                    index = headers.indexOf(header);

                indices[header] = index + 1;

                if (index !== -1) {
                    headersOfInterestIndices.push(index);
                }
            }

            minColumn = Math.min.apply(Math, headersOfInterestIndices) + 1;
            maxColumn = Math.max.apply(Math, headersOfInterestIndices) + 1;
            return {
                min: minColumn,
                max: maxColumn,
                indices: indices
            };
        }

        utils_interface = {
            prettyPrintDate: prettyPrintDate,
            normaliseAndValidateDuration: normaliseAndValidateDuration,
            pad: pad,
            setPadNumber: setPadNumber,
            getNamedValue: getNamedValue,
            setNamedValue: setNamedValue,
            loadData: loadData,
            saveData: saveData,
            findStatusColor: findStatusColor,
            diffDays: diffDays,
            findMinMaxColumns: findMinMaxColumns
        };

        return utils_interface;

    }(sfss.r, sfss.lg, sfss.smm));

} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}
Logger.log('leaving file util');