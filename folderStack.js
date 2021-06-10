/*
This folder stack is used to keep track of which folders are immediate children
of the folder currently being processed.

This is stored in a Google Sheet so it can persist across program executions,
bypassing the issue of Google Script executions timing out.

This allows pre-order processing to work safely, so the Outputter can still
print the folder structure properly
*/

function getOrCreateFolderStack(){
    let workbook = SpreadsheetApp.getActiveSpreadsheet();
    let name = "Folder Stack";
    let folderStackSheet = workbook.getSheetByName(name);
    if(folderStackSheet == null){
        folderStackSheet = workbook.insertSheet(name);
    }
    let fs = new FolderStack(folderStackSheet, 1);
    fs.insertHeader();
    return fs;
}

class FolderStack {

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
            "The Script will process these folders soon"
        );
    }

    /*
    make sure to pass the folder's ID, NOT the actual folder object
    */
    push(folderId){
        this.sheet.getRange(
            this.sheet.getLastRow() + 1,
            this.colNum
        ).setValue(folderId);
    }

    /*
    use this to check if all folders have been processed
    */
    isEmpty(){
        return this.sheet.getLastRow() == 1;
    }

    /*
    call this to get the ID of the next folder to process
    */
    peek(){
        return this.sheet.getRange(
            this.sheet.getLastRow(),
            this.colNum
        ).getValue();
    }

    /*
    Do not call this method before processing the folder given by this.peek().
    Otherwise, if this method is called and the script times out before
    processing completes, the folder is incorrectly marked complete.
    */
    pop(){
        let value = this.peek();
        this.sheet.getRange(
            this.sheet.getLastRow(),
            this.colNum
        ).deleteCells(SpreadsheetApp.Dimension.ROWS);
        return value;
    }
}
