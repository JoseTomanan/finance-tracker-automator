/**
 * Alias GAS for GoogleAppsScript
 */
declare namespace GAS {
    export import Spreadsheet = GoogleAppsScript.Spreadsheet;
    export import Script = GoogleAppsScript.Script;
    export import Utilities = GoogleAppsScript.Utilities;
}

/**
 * Actual globals
 */
const months: Array<String> = [
    'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'
    ];
const masterHeaderLabels: Array<Array<String>> = [
    ["E", "F"], ["G", "H"], ["I", "J"], ["K", "L"], ["M", "N"], ["O", "P"], ["Q", "R"]
    ];

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const scriptProperties = PropertiesService.getScriptProperties();

var master: MasterSheet;
var outgoing: OutgoingSheet;
var incoming: IncomingSheet;

/**
 * Project types
 */
type RowNumber = number;
enum Column {
    A = 1,
    B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z
}

interface SheetProtocol {
    getLastRow() : RowNumber
}
