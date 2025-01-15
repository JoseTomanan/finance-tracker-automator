/**
 * Project types
 */

declare namespace GAS {
    export import Spreadsheet = GoogleAppsScript.Spreadsheet;
    export import Script = GoogleAppsScript.Script;
    export import Utilities = GoogleAppsScript.Utilities;
}

type RowNumber = number;

enum Column {
    null = 0, A = 1,
    B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z
}

enum ExpenseType {
    NULL = "",
    DORM = "Dorm",
    FOOD = "Food",
    TRANSPO = "Transpo",
    LAUNDRY = "Laundry",
    HEALTH = "Health",
    MISC = "Misc",
    SELF = "Self"
}

class ExpenseEntry {
    type: ExpenseType = ExpenseType.NULL;
    cost: number = 0;
    entry: string = "";
}

interface SheetProtocol {
    sheet: GAS.Spreadsheet.Sheet;
    getLastRow() : RowNumber;
}