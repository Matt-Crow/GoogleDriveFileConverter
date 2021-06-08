/*
The folder queue is used to keep track of which folders still need to be
processed.

This is stored in a Google Sheet so it can persist across program executions,
bypassing this issue of Google Script executions timing out.
*/

class FolderQueue {

    /*
    sheet - the Google Sheet this persists in
    colNum - the 1-indexed column of the given sheet this is recorded in
    */
    constructor(sheet, colNum){
        this.sheet = sheet;
        this.colNum = colNum;
    }

    /*
    call this after creating this' sheet
    */
    insertHeader(){
        this.sheet.getRange(1, this.colNum).setValue(
            "Put folder URLs or IDs below this cell"
        );
    }

    /*
    make sure to pass the folder's ID, NOT the actual folder object
    */
    enqueue(folderId){
        this.sheet.getRange(
            this.sheet.getLastRow() + 1,
            this.colNum
        ).setValue(folderId);
    }

    /*
    use this to check if all folders have been processed
    */
    isEmpty(){
        return this.peek() == null;
    }

    /*
    call this to get the ID of the next folder to process
    */
    peek(){
        return this.sheet.getRange(
            2, // first row after the header
            this.colNum
        ).getValue();
    }

    /*
    Do not call this method before processing the folder given by this.peek().
    Otherwise, if this method is called and the script times out before
    processing completes, the folder is incorrectly marked complete.
    */
    dequeue(){
        let value = this.peek();
        this.sheet.getRange(
            2,
            this.colNum
        ).deleteCells(SpreadsheetApp.Dimension.COLUMNS);
        // shifts cells up after delete
        return value;
    }
}
