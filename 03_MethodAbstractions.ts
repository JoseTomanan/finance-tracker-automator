/**
 * Abstraction that adds new day
 * -- Currently very slow! Impractical to use; might be optimizable
 */
class DayAdder {
    addToday() : void
    {
        outgoing.addNewEntry();
    }
    
    compareRecentDate() : void
    {
        if (outgoing.isNeedsNewDay() === true)
            this.startNewDay(new Date());
    }

    startNewDay(mostRecentDate: Date) : void
    {
        this.addToday();
        
        new WeekHider().compareWeek();
        
        if (mostRecentDate.getMonth() != new Date().getMonth())
            new MonthAdder().addNewMonth();
    }
}

/**
 * Abstraction that compares weeks;
 * If new week, then hide all preceding weeks and add new day.
 */
class WeekHider {
    compareWeek() : void
    {
        outgoing.evaluateWeek();
    }
}

/**
 * Abstraction that handles the extensive process of adding a new month
 */
class MonthAdder {
    newMonthName: string = `${months[new Date().getMonth()]} ${new Date().getFullYear()}`;

    addNewMonth() : void
    {
        this.hideCurrentMonth();

        const incomingLastRow = incoming.sheet.getLastRow();

        const incomingTotalRow = this.findIncomingTotalRow(incomingLastRow);
        const incomingNewRowStart = incomingLastRow+4;

        this.updateIncomingSheet(this.newMonthName, incomingLastRow, incomingTotalRow);
        
        const masterLastRow = master.sheet.getLastRow();

        this.capOffAllotted(masterLastRow, incomingLastRow, incomingTotalRow);
    
        const masterNewRow = masterLastRow + 1;
        
        /**
         * For new month;
         * Create new row in masterSheet
         */
        master.sheet.insertRowAfter(masterLastRow);
        master.sheet.getRange(masterNewRow, 1).setValue(this.newMonthName);
    
        /**
         * Add necessary formulas for each column in new row
         */
        for (var i = 1; i <= masterHeaderLabels.length; i++) {
            master.sheet
                .getRange(masterNewRow, i*2)
                .setFormula(
                    `= SUMIF( \
                        ${ this.newMonthName }!$E4: $E, \
                        ${ masterHeaderLabels[i-1][0] }$1, \
                        ${ this.newMonthName }!$F4: $F \
                    )`
                );

            master.sheet
                .getRange(masterNewRow, i*2 + 1)
                .setFormula(
                    `= SUMIF( \
                        INCOMING! $C${ incomingNewRowStart-1 }: $C, \
                        ${ masterHeaderLabels[i-1][0] }$1, \
                        INCOMING! $B${ incomingNewRowStart-1 }: $B \
                    )`
                );
        }
    
        /**
         * Add TOTAL columns
         */
        master.sheet
            .getRange(masterNewRow, 1)
            .setFormula(
                `= SUM( \
                    E${ masterNewRow }, \
                    G${ masterNewRow }, \
                    I${ masterNewRow }, \
                    K${ masterNewRow }, \
                    M${ masterNewRow }, \
                    O${ masterNewRow }, \
                    Q${ masterNewRow } \
                )`
            );
        
        master.sheet
            .getRange(masterNewRow, 2)
            .setFormula(
                `= SUM( \
                    F${ masterNewRow }, \
                    H${ masterNewRow }, \
                    J${ masterNewRow }, \
                    L${ masterNewRow }, \
                    N${ masterNewRow }, \
                    P${ masterNewRow }, \
                    R${ masterNewRow } \
                )`
            );
    
        /**
         * Add ON HAND column
         */
        master.sheet
            .getRange(masterNewRow, 16)
            .setFormula(
                `= C${ masterNewRow } - B${ masterNewRow }`
                );
    
        /**
         * Activate month sheet, then add entry manually
         */
        outgoing.sheet.activate();
        outgoing.sheet.getRange(4, 2)
            .setValue(
                Utilities.formatDate(new Date(), "GMT+8", "MM/dd/yyyy")
                );
    }

    hideCurrentMonth() : void
    {
        outgoing.sheet.activate();
        spreadsheet.moveActiveSheet(3);
        outgoing.sheet.hideSheet();
    }

    createNewMonth() : string
    {
        const outgoingTemplate: GAS.Spreadsheet.Sheet = spreadsheet.getSheetByName("OUTGOINGTEMPLATE")!;
        
        const newMonthName = `${months[new Date().getMonth()]} ${new Date().getFullYear()}`;

        outgoingTemplate.copyTo(spreadsheet).setName(newMonthName);

        outgoing.setSheet( spreadsheet.getSheetByName(newMonthName)! );
        outgoing.sheet.activate();
        spreadsheet.moveActiveSheet(2);

        return newMonthName;
    }

    findIncomingTotalRow(incomingLastRow: number) : number
    {
        const dataIncomingFirstRow = incoming.sheet.getDataRange().getValues();
        
        var incomingTotalRow = -1;
        
        for (var i = incomingLastRow - 1; i >= 0; i--)
            if (dataIncomingFirstRow[i][0] === "TOTAL") {
                incomingTotalRow = i + 1;
                break;
            }

        return incomingTotalRow;
    }

    updateIncomingSheet(newMonthName: string, incomingLastRow: number, incomingTotalRow: number) : void
    {
        incoming.sheet
            .getRange(incomingTotalRow, 2)
            .setFormula("= SUM(B" + (incomingTotalRow + 1) + ":" + incomingLastRow + ")");
        
        const incomingNewRow = incomingLastRow + 3;
        const incomingNewRowStart = incomingNewRow + 1;

        incoming.sheet.insertRows(incomingLastRow + 1, 5);

        const copyDest = incoming.sheet.getRange(incomingNewRow, 1, 1, 3);
        incoming.sheet
            .getRange(incomingTotalRow, 1, 1, 3)
            .copyTo(copyDest);

        incoming.sheet.getRange(incomingNewRow - 1, 1).setValue(newMonthName);

        incoming.sheet
            .getRange(incomingNewRow, 2)
            .setFormula("= SUM(B" + incomingNewRowStart + ":B)");
        
        // incomingSheet.insertRowAfter(incomingNewRow);
    }

    capOffAllotted(masterLastRow: number, incomingLastRow: number, incomingTotalRow: number) {
        for (var i = 1; i <= masterHeaderLabels.length; i++) {
            master.sheet
                .getRange(masterLastRow, i*2 + 1)
                .setFormula(
                    `= SUMIF( \
                        INCOMING! $C${ incomingTotalRow }: $C${ incomingLastRow }, \
                        ${ masterHeaderLabels[i-1][0] }$1, \
                        INCOMING! $B${ incomingTotalRow }: $B${ incomingLastRow } \
                    )`
                );
        }
    }

    addMonthInMaster() : void
    {}

    addCellsInTotal() : void
    {}
}