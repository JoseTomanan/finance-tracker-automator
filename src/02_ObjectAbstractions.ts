/**
 * Dataclass for current month sheet
 */
class OutgoingSheet implements SheetProtocol {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheets()[1];
    datesRowOffset: RowNumber = 4;

    getLastRow() : RowNumber
    {
        return this.sheet.getLastRow();
    }

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
     * Archive current sheet (i.e. hide in Google Sheets) 
     */
    archiveSheet() : void
    {
        this.sheet.activate();
        spreadsheet.moveActiveSheet(3);
        this.sheet.hideSheet();
    }

    /**
     * Copy format of last row -- can override with offset from last row
     */
    copyLastRowFormat(offset: number = 0)
    {
        const currentEntry = this.sheet.getRange(this.getLastRow() - offset, Column.C);
        const nextEntry = this.sheet.getRange(this.getLastRow()+1, Column.C);
        
        // this.sheet.insertRowAfter(this.getLastRow());
        currentEntry.copyTo(nextEntry);
    }

    /**
     * Hooked to addToday();
     * Add new line with today's date -- can override with which row to put in date
     */
    addNewEntry(row: number = this.getLastRow()) : void
    {
        this.sheet.getRange(row, 2)
            .setValue(
                Utilities.formatDate(new Date(), "GMT+8", "MM/dd/yyyy")
                );
    }

    /**
     * Return whether or not recent entry corresponds to date today
     */
    isNeedsNewDay() : boolean
    {
        return this.getMostRecentDate().getDate() != new Date().getDate();
    }

    /**
     * Return whether or not new week is entered (i.e., new Sunday)
     */
    isNeedsNewWeek() : boolean
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

    /**
     * Hide previous weeks;
     * Unstable -- under construction!
     */
    hideLastWeek(startHideable: string | null) : void
    {
        const numRows = this.getLastRow() - this.datesRowOffset - 1;
        
        if ( startHideable !== null ) {
            this.sheet.hideRows(+startHideable, numRows);
        }
        else {
            this.sheet.hideRows(this.datesRowOffset, numRows);
        }
    }

    /**
     * Add new week label
     */
    labelNewWeek()
    {
        this.sheet.getRange(this.getLastRow()+1, Column.D)
            .setValue("<~~ NEW WEEK ~~>")
            .setHorizontalAlignment("center")
            .setFontStyle("italic");

        this.sheet.getRange(this.getLastRow(), Column.B)
            .setValue("--")
            .setHorizontalAlignment("center");

        this.sheet.getRange(this.getLastRow(), Column.E)
            .clear();
    }
}

/**
 * Dataclass for MASTER SHEET
 */
class MasterSheet implements SheetProtocol {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheetByName("MASTER SHEET")!;

    getLastRow() : RowNumber
    {
        return this.sheet.getLastRow();
    }

    capPrevAllotted(lastRow: RowNumber, totalRow: RowNumber) : void
    {
        const hereLastRow = this.getLastRow();

        for (var i = 0; i < masterHeaderLabels.length; i++) {
            this.sheet.getRange(hereLastRow, i*2 + 1)
                .setFormula(
                    `= SUMIF(INCOMING! $C${ totalRow }: $C${ lastRow }, ${ masterHeaderLabels[i][0] }$1, INCOMING! $B${ totalRow }: $B${ lastRow })`
                );
        }
    }

    addNewRow(newMonthName: string) : void
    {
        this.sheet.insertRowAfter(this.getLastRow());
        this.sheet.getRange(this.getLastRow()+1, Column.A)
            .setValue(newMonthName);
    }

    /**
     * Set formulas for the new month row (cost, alloted, total, on hand)
     */
    makeFormulas(newMonthName: string, incomingNewRow: RowNumber) : void
    {
        const lastRow = this.getLastRow();
        
        this.#makeCostCols(newMonthName);
        this.#makeAllottedCols(incomingNewRow);
        this.#makeTotalCols();

        this.sheet.getRange(lastRow, Column.O)
            .setFormula(`= C${ lastRow } - B${ lastRow }`);
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
                    `= SUMIF(${ newMonthName }!$E4: $E, ${ masterHeaderLabels[i][0] }$1, ${ newMonthName }!$F4: $F)`
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
                    `= SUMIF(INCOMING! $C${ incomingNewRow }: $C, ${ masterHeaderLabels[i][0] }$1, INCOMING! $B${ incomingNewRow }: $B)`
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
 * Dataclass for INCOMING sheet (i.e., incoming funds)
 */
class IncomingSheet implements SheetProtocol {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheetByName("INCOMING")!;
    newRowOffset: number = 4;
    
    getLastRow() : RowNumber
    {
        return this.sheet.getLastRow();
    }

    getTotalRow() : RowNumber
    {
        const dataFirstRow: Object[][] = this.sheet.getDataRange().getValues();
        
        for (var i = this.getLastRow() ; i >= 0 ; i--) {
            if (dataFirstRow[i][0] === "TOTAL")
                return i + 1;
        }

        return -1;
    }

    addNewMonth(newMonthName: string) : void
    {
        const totalRow = this.getTotalRow();
        const startingRow = this.getLastRow() + 2;
        const copyDest = this.sheet.getRange(startingRow+1, Column.A, 1, 3);

        this.sheet.insertRows(startingRow - 1, 5);
        this.sheet.getRange(totalRow, Column.A, 1, 3)
            .copyTo(copyDest);
        this.sheet.getRange(startingRow, Column.A)
            .setValue(newMonthName);
        this.sheet.getRange(startingRow + 1, Column.B)
            .setFormula(`= SUM(B${ startingRow + 2 } : B)`);
    }

    capOffTotalRow() : void
    {
        const totalRow = this.getTotalRow();

        this.sheet.getRange(totalRow, Column.B)
            .setFormula(`= SUM(B${ totalRow + 1 } : B${ this.getLastRow() })`);
    }
}