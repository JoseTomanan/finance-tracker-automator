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
        if (outgoing.isNeedsNewDay() === true) {
            new WeekHider().compareWeek();
            this.addToday();
            
            if (outgoing.getMostRecentDate().getMonth() != new Date().getMonth())
                new MonthAdder().addNewMonth();
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
        userProperties.deleteProperty("CURRENT_WEEK_FIRST_ENTRY");

        outgoing.archiveSheet();

        const newSheet = this.#instantiateNewMonth();
        outgoing.setSheet(newSheet);

        incoming.capOffTotalRow();
        incoming.addNewMonth(this.newMonthName);

        master.capPrevAllotted(incoming.getLastRow(), incoming.getTotalRow());
        master.addNewRow(this.newMonthName);
        master.makeFormulas(this.newMonthName, incoming.getLastRow() + 4);

        this.#activateNewMonth();
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
}