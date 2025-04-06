/**
* Abstraction that adds new day
*/
class DayAdder {
  addToday() : void {
    outgoing.addNewExpense();
  }
  
  compareRecentEntry() : void {
    if ( outgoing.isNewDay() ) {
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
  compareWeek() : void {
    if ( outgoing.isNewWeek() ) {
      this.#hideLastWeek();
      this.#labelNewWeek();
      
      outgoing.addNewExpense();
      
      new PropertyFetcher().setWeekVal(outgoing.getLastRow());
    }
  }
  
  /**
  * Hide previous weeks (based on stored key-value)
  */
  #hideLastWeek() : void {
    const startHideable = new PropertyFetcher().getWeekVal();
    const endHideable = outgoing.getLastRow();
    
    if (startHideable != null) {
      outgoing.hideRowSpan(+startHideable, endHideable);
    }
    else {
      outgoing.hideRowSpan(outgoing.datesRowOffset, endHideable);
    }
  }
  
  #labelNewWeek() {
    const useableRow = outgoing.getLastRow() + 1;
    
    outgoing.setCellValue(useableRow, Column.D, "<-- NEW WEEK -->");
  }
}

/**
* Abstraction that handles the extensive process of adding a new month
*/
class MonthAdder {
  newMonthName: string = `${months[new Date().getMonth()]} ${new Date().getFullYear()}`;
  
  addNewMonth() : void {
    new PropertyFetcher().setWeekVal(4);
    
    outgoing.archiveSheet();
    outgoing.setSheet(
        this.#instantiateNewMonth()
      );
    
    // const incomingTotalRow = incoming.getTotalRow();
    
    // if (incomingTotalRow == -1) {
    //     console.log("ERROR: getTotalRow returns nothing")
    //     return;
    // }
    
    // const remainingFunds = incoming.capOffAndReturnTotal(incomingTotalRow);
    
    // incoming.hidePrevMonth(incomingTotalRow);
    // incoming.initiateNewMonth(this.newMonthName);
    // incoming.addFundsEntry(
    //     new ExpenseEntry(Tag.SELF, remainingFunds, "overflow from last month", true)
    //     );
    
    // this.#capMasterPrevAllotted(incoming.getLastRow(), incomingTotalRow);
    
    master.addRow();
    master.setCellValue(master.getLastRow()+1, Column.B, this.newMonthName);
    
    master.makeFormulas(this.newMonthName, incoming.getLastRow() + 4);
    
    this.#activateNewMonth();
  }
  
  #instantiateNewMonth() : GAS.Spreadsheet.Sheet {
    spreadsheet.getSheetByName("OUTGOINGTEMPLATE")!
      .copyTo(spreadsheet).setName(this.newMonthName);
    
    const returnable = spreadsheet.getSheetByName(this.newMonthName)!;
    
    returnable.activate();
    spreadsheet.moveActiveSheet(2);
    
    return returnable;
  }
  
  #capMasterPrevAllotted(lastRow: RowNumber, totalRow: RowNumber) : void {
    const hereLastRow = master.getLastRow();
    
    for (var i = 0; i < masterHeaderLabels.length; i++) {
      master.setCellValue(
        hereLastRow, i*2 + 1,
        `= SUMIF(INCOMING! $E${ totalRow }: $E${ lastRow }, ${ masterHeaderLabels[i][0] }$1, INCOMING! $D${ totalRow }: $D${ lastRow })`
      );
    }
  }
  
  #activateNewMonth() : void {
    outgoing.sheet.activate();
    outgoing.sheet.getRange(4, 2)
      .setValue(Utilities.formatDate(new Date(), "GMT+8", "MM/dd/yyyy"));
  }
}