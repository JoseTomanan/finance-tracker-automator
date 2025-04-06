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