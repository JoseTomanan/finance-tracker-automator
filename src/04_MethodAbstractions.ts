/**
 * Abstraction that adds new day
 */
class DayAdder {
    addToday() : void
    {
        outgoing.formatNextRow();
        outgoing.addNewEntry();
    }
    
    compareRecentEntry() : void
    {
        if (outgoing.isNeedsNewDay() == true) {
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
        if (outgoing.isNeedsNewWeek() == true) {
            outgoing.hideLastWeek(
                scriptProperties.getProperty("CURRENT_WEEK_FIRST_ENTRY")
                );
            outgoing.labelNewWeek();
            
            // Logger.log(`BEFORE: ${ scriptProperties.getProperty("CURRENT_WEEK_FIRST_ENTRY") }`);
            scriptProperties.setProperty("CURRENT_WEEK_FIRST_ENTRY", `${ outgoing.getLastRow()+1 }`);
            // Logger.log(`AFTER: ${ outgoing.getLastRow()+1 }`)
            
            outgoing.formatNextRow();
            outgoing.addNewEntry();
        }
    }
}

/**
 * Abstraction that handles the extensive process of adding a new month
 */
class MonthAdder {
    newMonthName: string = `${months[new Date().getMonth()]} ${new Date().getFullYear()}`;

    #resetVars() : void
    {
        scriptProperties.setProperty("CURRENT_WEEK_FIRST_ENTRY", "4");
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
        
        incoming.capOffTotalRow(incomingTotalRow);
        incoming.hidePrevMonth(incomingTotalRow);
        incoming.addNewMonth(this.newMonthName);
        
        incoming.addFundsEntry();

        master.capPrevAllotted(incoming.getLastRow(), incomingTotalRow);
        master.addNewRow(this.newMonthName);
        master.makeFormulas(this.newMonthName, incoming.getLastRow() + 4);

        this.#activateNewMonth();
    }
}