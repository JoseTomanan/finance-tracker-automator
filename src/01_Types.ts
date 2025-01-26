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
enum Tag {
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
    readonly tag: Tag;
    readonly cost: number;
    readonly entry: string;
    readonly isIncoming: boolean;

    constructor(tag: Tag = Tag.NULL, cost: number = 0, entry: string = "", isIncoming: boolean = false) {
        this.tag = tag;
        this.cost = cost;
        this.entry = entry;
        this.isIncoming = isIncoming;
    }
}

/**
 * Interface for sheet-related classes
 */
abstract class Sheet {
    abstract sheet: GAS.Spreadsheet.Sheet;

    getLastRow() : RowNumber
    {
        return this.sheet.getLastRow();
    }

    setCellValue(r: RowNumber, c: Column, val: string) : void
    {
        const cell = this.sheet.getRange(r,c).setValue(val);

        // if (isCenter)
        //     cell.setHorizontalAlignment("center");

        // if (isItalic)
        //     cell.setFontStyle("italic");
    }

    clearCell(row: RowNumber, col : Column)
    {
        this.sheet.getRange(row, col).clear();
    }

    addRow()
    {
        this.sheet.insertRowAfter(this.getLastRow());
    }

    /**
     * High level abstraction for Sheet.hideRows
     */
    hideRowSpan(start: number, end: number) : void
    {
        this.sheet.hideRows(start, end-start-1);
    }
}