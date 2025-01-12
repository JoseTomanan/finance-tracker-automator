/**
 * Actual code for positioning functions exported to GSheets
 */
master = new MasterSheet();
outgoing = new OutgoingSheet();
incoming = new IncomingSheet();


function onOpen(e: any)
{
    new DayAdder().compareRecentEntry();
    
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Methods')
        .addItem('Add today', 'addToday')
        .addItem('Archive previous weeks', 'compareWeek')
        .addItem('Add new month', 'addNewMonth')
        .addToUi();
}

function addToday()
{
    new DayAdder().addToday();
}

function compareWeek()
{
    new WeekHider().compareWeek();
}

function addNewMonth()
{
    new MonthAdder().addNewMonth();
}