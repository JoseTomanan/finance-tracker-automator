/**
 * Dataclass for current month sheet (i.e. outgoing funds)
 */
class OutgoingSheet {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheets()[1];
    datesRowOffset: number = 4;

    getDates() : Object[][]
    {
        return this.sheet.getRange("B:B").getValues();
    }

    getLastRow() : number
    {
        return this.sheet.getLastRow();
    }

    getCurrentEntry() : GAS.Spreadsheet.Range
    {
        return this.sheet.getRange(this.getLastRow(), 2, 1, 2);
    }

    getNextEntry() : GAS.Spreadsheet.Range
    {
        return this.sheet.getRange(this.getLastRow()+1, 2, 1, 2);
    }

    setSheet(newSheet: GAS.Spreadsheet.Sheet) : void
    {
        this.sheet = newSheet;
    }

    /**
     * Hooked to addToday();
     * Add new line with today's date
     */
    addNewEntry() : void
    {
        this.sheet.insertRowAfter(this.getLastRow());
        this.getCurrentEntry().copyTo(this.getNextEntry());
        
        const sameDay = Utilities.formatDate(new Date(), "GMT+8", "MM/dd/yyyy");

        outgoing.sheet.getRange( this.getLastRow(),2,1,1 ).setValue(sameDay);
    }

    /**
     * Return whether or not recent entry corresponds to date today
     */
    isNeedsNewDay() : boolean
    {
        const columnData = outgoing.sheet.getRange("B:B").getValues();
        const lastRow = columnData.filter(String).length + 1;
        
        const mostRecentDate = new Date( columnData[lastRow][0] );
        
        return (mostRecentDate.getDate() != new Date().getDate());
    }

    /**
     * Hooked to compareWeek();
     * Evaluate if new week is entered
     */
    evaluateWeek() : void
    {
        if (this.#isNeedsNewWeek() === true) {
            const lastRow = this.getLastRow();
            const lastEntryFormat = outgoing.sheet.getRange(
                lastRow, 1, 1, outgoing.sheet.getLastColumn()
                );
            const newEntry = outgoing.sheet.getRange(
                lastRow+1, 1, 1, outgoing.sheet.getLastColumn()
                );

            lastEntryFormat.copyTo(
                newEntry, {formatOnly: true}
                );

            this.#hideLastWeek();
        }
    }

    /**
     * Return whether or not new week is entered (i.e., new Sunday)
     */
    #isNeedsNewWeek() : boolean
    {
        const columnData: any[][] = outgoing.sheet.getRange("B4:B").getValues();
        const latestEntryDate: Date = new Date(
            columnData[this.getLastRow() - this.datesRowOffset][0]
            );
        const latestEntryWeek: string = Utilities.formatDate(
            latestEntryDate,
            spreadsheet.getSpreadsheetTimeZone(), 'w'
            );
        const currentWeek: string = Utilities.formatDate(
            new Date(),
            spreadsheet.getSpreadsheetTimeZone(), 'w'
            );
        
        return (latestEntryWeek < currentWeek);
    }

    /**
     * Hide previous weeks. Under construction!
     */
    #hideLastWeek() : void
    {
        outgoing.sheet.hideRows(4, this.getLastRow() - this.datesRowOffset+1);
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
}

/**
 * Dataclass for master sheet
 */
class MasterSheet {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheetByName("MASTER SHEET")!;

    getLastRow() : number
    {
        return this.sheet.getLastRow();
    }

    capPrevAllotted(lastRow: number, totalRow: number) : void
    {
        const hereLastRow = this.getLastRow();

        for (var i = 1; i <= masterHeaderLabels.length; i++) {
            master.sheet
                .getRange(hereLastRow, i*2 + 1)
                .setFormula(
                    `= SUMIF(INCOMING! $C${ totalRow }: $C${ lastRow }, ${ masterHeaderLabels[i-1][0] }$1, INCOMING! $B${ totalRow }: $B${ lastRow })`
                );
        }
    }

    addNewRow(newMonthName: string) : void
    {
        this.sheet.insertRowAfter(this.getLastRow());
        this.sheet.getRange(this.getLastRow()+1, 1).setValue(newMonthName);
    }

    /**
     * Set formulas for the new month row (cost, alloted, total, on hand)
     */
    makeFormulas(newMonthName: string, incomingNewRow: number) : void
    {
        const lastRow = this.getLastRow();
        
        this.#makeCostCols(newMonthName);
        this.#makeAllottedCols(incomingNewRow);
        this.#makeTotalCols();

        master.sheet
            .getRange(lastRow, 16)
            .setFormula(
                `= C${ lastRow } - B${ lastRow }`
                );
    }

    /**
     * Add COST columns for each category in new month
     */
    #makeCostCols(newMonthName: string)
    {
        const lastRow = this.getLastRow();

        for (var i = 0; i < masterHeaderLabels.length; i++) {
            master.sheet
                .getRange(lastRow, i*2)
                .setFormula(
                    `= SUMIF(${ newMonthName }!$E4: $E, ${ masterHeaderLabels[i][0] }$1, ${ newMonthName }!$F4: $F)`
                );
        }
    }

    /**
     * Add ALLOTTED columns for each category in new month
     */
    #makeAllottedCols(incomingNewRow: number)
    {
        const lastRow = this.getLastRow();

        for (var i = 0; i < masterHeaderLabels.length; i++) {
            master.sheet
                .getRange(lastRow, i*2 + 1)
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

        master.sheet.getRange(lastRow, 1)
            .setFormula(`= SUM(${ costString })`);
        
        master.sheet.getRange(lastRow, 2)
            .setFormula(`= SUM(${ allottedString })`);
    }
}

/**
 * Dataclass for INCOMING sheet (i.e., incoming funds)
 */
class IncomingSheet {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheetByName("INCOMING")!;
    newRowOffset: number = 4;
    
    getLastRow() : number
    {
        return this.sheet.getLastRow();
    }

    getTotalRow() : number
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

        this.sheet.getRange(totalRow, 2)
            .setFormula(`= SUM(B${ totalRow + 1 } : B${ this.getLastRow() })`);

        const startingRow = this.getLastRow() + 2;
        const copyDest = incoming.sheet.getRange(startingRow+1, 1, 1, 3);

        this.sheet.insertRows(startingRow - 1, 5);
        this.sheet.getRange(totalRow, 1, 1, 3).copyTo(copyDest);
        this.sheet.getRange(startingRow, 1).setValue(newMonthName);
        this.sheet.getRange(startingRow + 1, 2)
            .setFormula(`= SUM(B${ startingRow + 2 } : B)`);
    }

    #capOffTotal() : void
    {
        const totalRow = this.getTotalRow();
    }
}