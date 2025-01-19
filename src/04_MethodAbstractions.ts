/**
 * Abstraction that adds new day
 */
class DayAdder {
    addToday() : void
    {
        outgoing.addNewEntry();
    }
    
    compareRecentEntry() : void
    {
        if ( ! outgoing.isSameDay() ) {
            new WeekHider().compareWeek();
            
            if (outgoing.getMostRecentDate().getMonth() != new Date().getMonth()) {
                new MonthAdder().addNewMonth();
            }
        }
    }
}

/**
 * Abstraction that compares weeks;
 * If new week, then hide all preceding weeks and add new day.
 */
class WeekHider {
    compareWeek() : void
    {
        if ( ! outgoing.isSameWeek() ) {
            this.#hideLastWeek();

            outgoing.labelNewWeek();
            outgoing.addNewEntry();
            
            new PropertyFetcher().setWeekVal( outgoing.getLastRow()+1 );
        }
    }

    /**
     * Hide previous weeks (based on stored key-value)
     */
    #hideLastWeek() : void
    {
        const startHideable = new PropertyFetcher().getWeekVal();
        const endHideable = outgoing.getLastRow();

        if (startHideable != null) {
            outgoing.hideRows(+startHideable, endHideable);
            return;
        }

        outgoing.hideRows(outgoing.datesRowOffset, endHideable);
    }
}

/**
 * Abstraction that handles the extensive process of adding a new month
 */
class MonthAdder {
    newMonthName: string = `${months[new Date().getMonth()]} ${new Date().getFullYear()}`;

    #resetVars() : void
    {
        new PropertyFetcher().setWeekVal(4);
    }

    #instantiateNewMonth() : GAS.Spreadsheet.Sheet
    {
        spreadsheet.getSheetByName("OUTGOINGTEMPLATE")!
            .copyTo(spreadsheet).setName(this.newMonthName);

        const returnable = spreadsheet.getSheetByName(this.newMonthName)!;

        returnable.activate();
        spreadsheet.moveActiveSheet(2);

        return returnable;
    }

    #activateNewMonth() : void
    {
        outgoing.sheet.activate();
        outgoing.sheet.getRange(4, 2)
            .setValue(Utilities.formatDate(new Date(), "GMT+8", "MM/dd/yyyy"));
    }

    addNewMonth() : void
    {
        this.#resetVars();
        
        outgoing.archiveSheet();
        outgoing.setSheet(
            this.#instantiateNewMonth()
            );

        const incomingTotalRow = incoming.getTotalRow();
        const remainingFunds = incoming.capOffAndReturnTotal(incomingTotalRow);

        incoming.hidePrevMonth(incomingTotalRow);
        incoming.addNewMonth(this.newMonthName);
        incoming.addFundsEntry(
            new ExpenseEntry(Tag.SELF, remainingFunds, "overflow from last month", true)
            );

        master.capPrevAllotted(incoming.getLastRow(), incomingTotalRow);
        master.addNewRow(this.newMonthName);
        master.makeFormulas(this.newMonthName, incoming.getLastRow() + 4);

        this.#activateNewMonth();
    }
}