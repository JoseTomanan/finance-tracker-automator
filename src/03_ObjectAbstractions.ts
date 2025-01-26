/**
 * High-level abstraction for setting, getting script properties
 */
class PropertyFetcher {
    /**
     * Return CURRENT_WEEK_FIRST_ENTRY
     */
    getWeekVal() {
        scriptProperties.getProperty("CURRENT_WEEK_FIRST_ENTRY");
    }

    /**
     * Set CURRENT_WEEK_FIRST_ENTRY to new value
     */
    setWeekVal(row : RowNumber) {
        scriptProperties.setProperty("CURRENT_WEEK_FIRST_ENTRY", `${ row }`);
    }
}

/**
 * Abstraction for current month sheet
 */
class OutgoingSheet extends Sheet {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheets()[1];
    datesRowOffset: RowNumber = 4;

    getMostRecentDate() : Date
    {
        const columnData = this.sheet.getRange("B:B").getValues();
        const lastRow = columnData.filter(String).length + 1;
        
        return new Date( columnData[lastRow][0] );
    }
    
    setSheet(newSheet: GAS.Spreadsheet.Sheet) : void
    {
        this.sheet = newSheet;
    }

    /** 
     * High level abstraction for Sheet.active -> Sheet.hideSheet
     */
    archiveSheet() : void
    {
        this.sheet.activate();
        spreadsheet.moveActiveSheet(3);
        this.sheet.hideSheet();
    }

    /**
     * Hooked to addToday();
     * Add new line with today's date -- optional: expense entry, type
     */
    addNewExpense(e: ExpenseEntry = new ExpenseEntry()) : void
    {
        this.#formatNewExpense();
        this.addRow();

        const currentRow = this.getLastRow();
            
        this.setCellValue(currentRow, Column.B,
            Utilities.formatDate(new Date(), "GMT+8", "MM/dd/yyyy")
            );

        if (e.tag != Tag.NULL) {
            this.setCellValue(currentRow, Column.D, e.entry);
            this.setCellValue(currentRow, Column.E, e.tag);
            this.setCellValue(currentRow, Column.F, `${e.cost}`);
        }

    }

    #formatNewExpense(offset: number = 0) : void
    {
        const entry = this.getLastRow() + 1 - offset;
        this.setCellValue(entry, Column.C,
            `= TEXT(weekday(B${ entry }), "ddd")`
            );
    }

    /**
     * Return whether or not recent entry corresponds to date today
     */
    isNewDay() : boolean
    {
        return this.getMostRecentDate().getDate() != new Date().getDate();
    }

    /**
     * Return whether or not new week is entered (i.e., new Sunday)
     */
    isNewWeek() : boolean
    {
        const latestEntryWeek = Utilities.formatDate(
            this.getMostRecentDate(),
            spreadsheet.getSpreadsheetTimeZone(), 'w'
            );
        const currentWeek = Utilities.formatDate(
            new Date(),
            spreadsheet.getSpreadsheetTimeZone(), 'w'
            );

        return latestEntryWeek < currentWeek;
    }
}

/**
 * Abstraction for MASTER SHEET
 */
class MasterSheet extends Sheet {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheetByName("MASTER SHEET")!;

    /**
     * Set formulas for the new month row (cost, alloted, total, on hand)
     */
    makeFormulas(newMonthName: string, incomingNewRow: RowNumber) : void
    {
        const lastRow = this.getLastRow();
        
        this.#makeCostCols(newMonthName);
        this.#makeAllottedCols(incomingNewRow);
        this.#makeTotalCols();

        this.setCellValue(lastRow, Column.O, `= C${ lastRow } - B${ lastRow }`);
    }

    /**
     * Add COST columns for each category in new month
     */
    #makeCostCols(newMonthName: string)
    {
        const lastRow = this.getLastRow();

        for (var i = 0; i < masterHeaderLabels.length; i++) {
            this.sheet.getRange(lastRow, i*2)
                .setFormula(
                    `= SUMIF(${ newMonthName }!$E4: $E, ${ masterHeaderLabels[i][0] }$1, ${ newMonthName }!$F4:$F)`
                );
        }
    }

    /**
     * Add ALLOTTED columns for each category in new month
     */
    #makeAllottedCols(incomingNewRow: RowNumber)
    {
        const lastRow = this.getLastRow();

        for (var i = 0; i < masterHeaderLabels.length; i++) {
            this.sheet.getRange(lastRow, i*2 + 1)
                .setFormula(
                    `= SUMIF(INCOMING! $C${ incomingNewRow }:$C, ${ masterHeaderLabels[i][0] }$1, INCOMING! $B${ incomingNewRow }:$B)`
                );
        }
    }

    /**
     * Add TOTAL columns (cost, allotted)
     */
    #makeTotalCols()
    {
        const lastRow = this.getLastRow();

        var costString: string = "";
        var allottedString: string = "";

        for (var i = 0; i < masterHeaderLabels.length; i++) {
            costString += `${ masterHeaderLabels[i][0] }${ lastRow } ,`;
            allottedString += `${ masterHeaderLabels[i][1] }${ lastRow } ,`;
        }

        this.sheet.getRange(lastRow, Column.A)
            .setFormula(`= SUM(${ costString })`);
        this.sheet.getRange(lastRow, Column.B)
            .setFormula(`= SUM(${ allottedString })`);
    }
}

/**
 * Abstraction for INCOMING sheet (i.e., incoming funds)
 */
class IncomingSheet extends Sheet {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheetByName("INCOMING")!;
    newRowOffset: number = 4;

    getTotalRow() : RowNumber
    {
        const dataFirstRow: Object[][] = this.sheet.getDataRange().getValues();
        
        for (var i = this.getLastRow() ; i >= 0 ; i--) {
            if (dataFirstRow[i][0] == "TOTAL")
                return i + 1;
        }

        return -1;
    }

    capOffAndReturnTotal(totalRow : RowNumber = this.getTotalRow()) : number
    {
        const traversable = this.sheet.getRange(totalRow, Column.B);
        
        traversable.setFormula(`= SUM(B${ totalRow + 1 } : B${ this.getLastRow() })`);

        return +traversable.getDisplayValue();
    }

    hidePrevMonth(startHideable : RowNumber = this.getTotalRow()) : void
    {
        const numRows = this.getLastRow() - startHideable;
        this.sheet.hideRows(startHideable, numRows);
    }

    initiateNewMonth(newMonthName: string) : void
    {
        const totalRow = this.getTotalRow();
        const startingRow = this.getLastRow() + 2;
        const copyDest = this.sheet.getRange(startingRow+1, 1, 1, 3);

        this.sheet.getRange(totalRow, 1, 1, 3)
            .copyTo(copyDest);

        this.setCellValue(startingRow, Column.A, newMonthName);
        this.setCellValue(startingRow+1, Column.B, `= SUM(B${ startingRow + 2 }:B)`);
    }

    addFundsEntry(e: ExpenseEntry) : void
    {
        if ( e.isIncoming ) {
            const newRow = this.getLastRow()+1;

            this.setCellValue(newRow, Column.A, e.entry);
            this.setCellValue(newRow, Column.B, `${e.cost}`);
            this.setCellValue(newRow, Column.C, e.tag);
        }
    }
}