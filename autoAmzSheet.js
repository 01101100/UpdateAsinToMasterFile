/**
 * Copyright Long Trinh Thien
 *
 */

/**
 * get list asins in master file to check duplicate
 * @return {array} array of asins
 */
var memberFileIds = [{
    'id': '',
    'name': ''
}, {
    "id": '',
    'name': ''
}, {
    "id": "",
    "name": ""
}, {
    "id": "",
    "name": ""
}, {
    "id": "",
    "name": ""
}, {
    "id": "",
    "name": ""
}, {
    "id": "",
    "name": ""
}];

var masterSpreadsheetId = '';

function getMasterFileAsin() {
    var masterData = []; // to store list of asins
    var values = [];
    values.push(Sheets.Spreadsheets.Values.get(masterSpreadsheetId, "2018!B2:B").values);
    values.push(Sheets.Spreadsheets.Values.get(masterSpreadsheetId, "2017!B2:B").values);
    for (var i = 0; i < values.length; i++) {
        for (var row = 0; row < values[i].length; row++) {
            if (values[i][row][0] == null) continue;
            masterData.push(values[i][row][0]);
        }
    }
    return masterData;
}

/**
 * get all member data by sheet ID
 * @param  {string} SheetId ID of google Spreadsheet, get in URL.
 * @return {array[][]}         data of sheet
 */
function getMemberData(SheetId) {
    var rangeName = 'Now!C2:M'; //TODO
    var values = Sheets.Spreadsheets.Values.get(SheetId, rangeName).values;
    return values;
}

function getLastestRowHasAsin(SheetId) {
    var res = 0;
    var rangeName = 'Now!C2:C'; //TODO
    var values = Sheets.Spreadsheets.Values.get(SheetId, rangeName).values;
    Logger.log("getLastestRowHasAsin");
    return values.length + 1;
}

/**
 * after scanning is all done, it will mark the lastest row with color: #ff00ff
 * @return {none}
 */
function markLastest() {
    memberFileIds.forEach(function(member) {
        var sheet = SpreadsheetApp.openById(member['id']).getSheets()[0];
        var lastRow = getLastestRowHasAsin(member['id']);
        var lastRange = sheet.getRange(lastRow, 1);
        Logger.log(lastRow);
        lastRange.setBackground('#ff00ff');
    });
}

/**
 * find the lastest checked row to continue scanning.
 * @param  {string} sheetId ID of spreadsheet.
 * @return {int}         position of lastest checked row.
 */
function findLastestCheck(sheetId) {
    var res = 0;
    var sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
    var lastRow = sheet.getLastRow();
    for (var row = 1; row <= lastRow; row++) {
        if (sheet.getRange(row, 1).getBackground() == '#ff00ff') res = row;
    }
    if (res == lastRow) res = -1; // nothing new.
    return res;
}

function getLastRowNo() {
    var sheet = SpreadsheetApp.openById(masterSpreadsheetId).getSheets()[0];
    var lastRow = sheet.getLastRow();
    var lastRowValue = sheet.getRange(lastRow, 1).getValue();
    return lastRowValue;
}

function main() {
    var masterSheet = SpreadsheetApp.openById(masterSpreadsheetId).getSheetByName('2018'); // open to append row
    var masterAsins = getMasterFileAsin();
    var lastNo = getLastRowNo();

    memberFileIds.forEach(function(member) {
        memberData = getMemberData(member['id']);

        var lastCheck = findLastestCheck(member['id']);
        for (var row = 0; row < memberData.length; row++) {
            if (lastCheck == -1) break; // Member chua co them ASIN moi
            if (row < lastCheck - 1) continue; // Bo qua cac ASIN da check
            Logger.log("first data: " + memberData[row][0]);
            Logger.log("len: " + memberData.length);
            if (!memberData[row][0]) {
                continue;
            } else if (masterAsins.indexOf(memberData[row][0]) >= 0) { // ASIN da ton tai
                Logger.log("duplicate at Row: " + row + "data: " + memberData[row][0]);
                var duplicateRow = SpreadsheetApp.openById(member['id'])
                    .getSheets()[0]
                    .getRange(row + 2, 3);
                duplicateRow.setBackground('red');
            } else {
                memberData[row].unshift(++lastNo);
                masterSheet.appendRow(memberData[row])
            }
        }
    });
    markLastest();
}

function onOpen(e) {
    SpreadsheetApp.getUi()
        .createMenu('-_- Tool Auto -_-')
        .addItem('Update ASIN to master File', 'main')
        .addToUi();
}