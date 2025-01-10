/**
 * Dataclass for current month sheet (i.e. outgoing funds)
 */
class OutgoingSheet {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheets()[1];
    dates: Object[][] = this.sheet.getRange("B:B").getValues();

    setSheet(newSheet: GAS.Spreadsheet.Sheet) : void
    {
        this.sheet = newSheet;
    }

    getLastRow() : number
    {
        return this.sheet.getLastRow();
    }

    getCurrentEntry() : GAS.Spreadsheet.Range
    {
        return this.sheet.getRange(this.sheet.getLastRow(), 2, 1, 2);
    }

    getNextEntry() : GAS.Spreadsheet.Range
    {
        return this.sheet.getRange(this.sheet.getLastRow()+1, 2, 1, 2);
    }

    addNewEntry() : void
    {
        this.sheet.insertRowAfter(this.getLastRow());
        this.getCurrentEntry().copyTo(this.getNextEntry());
        
        const sameDay = Utilities.formatDate(new Date(), "GMT+8", "MM/dd/yyyy");

        outgoing.sheet.getRange(this.getLastRow(), 2, 1, 1)
            .setValue(sameDay);
    }

    isNeedsNewDay() : boolean
    {
        const columnData = outgoing.sheet.getRange("B:B").getValues();
        const lastRow = columnData.filter(String).length + 1;
        
        const mostRecentDate = new Date( columnData[lastRow][0] );
        
        return (mostRecentDate.getDate() != new Date().getDate());
    }

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

    #isNeedsNewWeek() : boolean
    {
        const columnData: any = outgoing.sheet.getRange("B4:B").getValues();
        const latestEntryDate: Date = new Date(
            columnData[ this.getLastRow() ][ 0 ]
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

    #hideLastWeek() : void
    {
        /**
         * Under construction!
         */
        outgoing.sheet.hideRows(4, this.getLastRow()-3);
    }
}

/**
 * Dataclass for master sheet
 */
class MasterSheet {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheetByName("MASTER SHEET")!;
    lastRow: number = this.sheet.getLastRow();
    // totalCol: number = ...;
    // onHandCol: number = ...;

    addNewEntry() : void {
        this.sheet.insertRowAfter(this.lastRow);

        /**
         * Add necessary formulas for each column in new row
         */

        /**
         * Add TOTAL columns: cost, allotted
         */

        /**
         * Add ON HAND column
         */
    }

    capOffAllotted() : void {}

    copyPrevRowFormat() : void {}
}

/**
 * Dataclass for INCOMING sheet (i.e., incoming funds)
 */
class IncomingSheet {
    sheet: GAS.Spreadsheet.Sheet = spreadsheet.getSheetByName("INCOMING")!;
    lastRow: number = this.sheet.getLastRow();
    totalRow: number = this.findTotalRow();
    newRowOffset: number = 4;
    
    findTotalRow() : number {
        const dataFirstRow: Object[][] = this.sheet.getDataRange().getValues();
        
        for (var i = this.lastRow-1 ; i >= 0 ; i--)
            if (dataFirstRow[i][0] === "TOTAL")
                return i + 1;

        return -1;
    }

    capOffTotal() : void {}

    updateNewMonth() : void {}
}