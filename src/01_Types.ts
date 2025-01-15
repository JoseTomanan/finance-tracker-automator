/**
 * Alias for brevity: GoogleAppsScript -> GAS
 */
declare namespace GAS {
    export import Spreadsheet = GoogleAppsScript.Spreadsheet;
    export import Script = GoogleAppsScript.Script;
    export import Utilities = GoogleAppsScript.Utilities;
}

/**
 * Enum to connote row number
 */
type RowNumber = number;

/**
 * Enum to simplify column use
 */
enum Column {
    null = 0, A = 1,
    B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z
}

/**
 * Enum to limit possible entries in Type row
 */
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

/**
 * Frozen dataclass for ExpenseEntry
 */
class ExpenseEntry {
    readonly type: ExpenseType;
    readonly cost: number;
    readonly entry: string;

    constructor(type: ExpenseType = ExpenseType.NULL, cost: number = 0, entry: string = "") {
        this.type = type;
        this.cost = cost;
        this.entry = entry;
    }
}

/**
 * Interface for sheet-related classes
 */
interface SheetProtocol {
    sheet: GAS.Spreadsheet.Sheet;
    getLastRow() : RowNumber;
}