/*
The Outputter is used to display the folder structure nicely in a Google sheet.
This makes it easier for the user to keep track of which files the script has
processed
*/

class Outputter {
    constructor(sheet){
        this.sheet = sheet;
        this.row = 1;
        this.col = 1;
    }

    outputFolder(folder){
        let link = SpreadsheetApp.newRichTextValue().setText(folder.getName()).setLinkUrl(folder.getUrl()).build();
        this.sheet.getRange(this.row, this.col).setValue(link);
        this.row++;
        this.col++;
    }

    doneWithFolder(){
        this.col--;
    }

    outputFile(file){
        this.sheet.getRange(this.row, this.col).setValue(file.getName());
        this.row++;
    }

    linkForFile(file){
        let link = SpreadsheetApp.newRichTextValue().setText(file.getName()).setLinkUrl(file.getUrl()).build();
    }

    outputReformattedFile(orig, reform){
        this.sheet.getRange(this.row, this.col, 1, 3).setValues([
            [this.linkForFile(orig), "--->", this.linkForFile(reform)]
        ]);
        this.row++;
    }
}
