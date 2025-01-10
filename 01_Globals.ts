/**
 * Actual globals
 */

declare namespace GAS {
    export import Spreadsheet = GoogleAppsScript.Spreadsheet;
    export import Script = GoogleAppsScript.Script;
    export import Utilities = GoogleAppsScript.Utilities;
}

const months: Array<String> = [
    'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'
    ];

const masterHeaderLabels: Array<Array<String>> = [
    ["E", "F"], ["G", "H"], ["I", "J"], ["K", "L"], ["M", "N"], ["O", "P"], ["Q", "R"]
    ];

const spreadsheet: GAS.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

var master: MasterSheet;
var outgoing: OutgoingSheet;
var incoming: IncomingSheet;