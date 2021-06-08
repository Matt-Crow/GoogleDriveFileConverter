/*
The folder structure of a Google Drive is tree-like in nature, with folders
(assuming they contain at least one file) as internal nodes, and files as leaf
nodes. This script must process every descendent leaf node of the root nodes it
is given.

Google Script, however, adds a wrinkle into the standard tree-traversal strategy:
Scripts are limited in how long they can run, usually around 5 minutes. Therefore,
the script must have some way of remembering which folders it has already finished
processing, while also remembering the child folders of those folders it has yet
to process.

There are multiple steps involved in processing a folder structure:
1. get a root folder
2. recur for each subfolder in that folder
3. process all files in that folder
4. mark the root folder as done
Steps 2 & 3 are interchangable for either post-order or pre-order traversal.

The problem with this is that Google Script will likely time out before step 4,
and so it may never be marked complete for large folder structures.

Here is the proposed alternative method:
1. get a root folder from the queue
2. push each subfolder to a processing stack, stored in a Google Sheet
3. process all files in the folder
4. mark the root folder as done, removing it from the queue
5. pop the top of the processing stack, and move it to the start of the queue



!!! HERE !!!
Essentially, I need to figure out how to process the folders pre-order while also
maintaining the indentation output given by the outputter and keeping in mind
Google Script's limitations
*/

const PROCESSING_LIST_NAME = "folders to process";
const EXTENSIONS = ["doc", "docx", "xls", "xlsx", "ppt", "pptx"].map((e)=>e.toLowerCase());

// I would like to keep the folder structure display the old version had
// stack and queue must each have their own sheet, otherwise getLastRow() will not work
// (when one is larger than the other)
function onOpen(){
    SpreadsheetApp
    .getUi()
    .createMenu("File Converter")
    .addItem("create resources", "createResources")
    .addItem("run", "run")
    .addToUi();
}

function createResources(){
    new ProcessingList(PROCESSING_LIST_NAME).createIfAbsent();
}

class ProcessingList {
    constructor(sheetName){
        this.sheetName = sheetName;
    }

    createIfAbsent(){
        let workbook = SpreadsheetApp.getActiveSpreadsheet();
        let procList = workbook.getSheetByName(this.sheetName);
        if(procList === null){
            procList = workbook.insertSheet(this.sheetName);
            procList.getRange(1, 1).setValue("Put folder URLs or IDs below this cell");
        }
    }

    getSheet(){
        return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheetName);
    }

    initialize(){
        this.sheet = this.getSheet(); // cache sheet so it only calls the SpreadsheetApp service once
    }

    hasMoreFolders(){
        return this.sheet.getRange(2, 1).getValue() != null;
    }

    getNextFolder(){
        return DriveApp.getFolderById(extractFolderIdFromUrl(this.sheet.getRange(2, 1).getValue()));;
    }

    doneWithFolder(){
        this.sheet.deleteRow(2);
    }

    enqueueFolder(id){
        this.sheet.getRange(
            this.sheet.getLastRow() + 1,
            1
        ).setValue(id);
    }
}

function extractFolderIdFromUrl(url){
    const regex = /\/drive(\/u\/[^\/]*)?\/folders\/([^\/?]*)\/?/;
    let id = url;
    if(regex.test(url)){
        let i = url.match(regex);
        id = url.match(regex)[2]; // ID is now the second regex group
    }
    return id;
}

function run(){
    let outputSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    let processor = new FileProcessor(
        new Outputter(outputSheet),
        new ProcessingList(PROCESSING_LIST_NAME)
    );
    processor.run();
}

/*
This handles file conversion and outputting
*/
class FileProcessor {
    constructor(outputter, processingList){
        this.outputter = outputter;
        this.processingList = processingList;
    }

    run(){
        this.processingList.initialize();
        while(this.processingList.hasMoreFolders()){
            this.process(this.processingList.getNextFolder());
            this.processingList.doneWithFolder();
        }
    }

    process(folder){
        this.outputter.outputFolder(folder);

        let fileIter = folder.getFiles();
        while(fileIter.hasNext()){
            this.processFile(fileIter.next());
        }

        let dirIter = folder.getFolders();
        while(dirIter.hasNext()){
            this.processingList.enqueueFolder(dirIter.next().getId());
        }

        this.outputter.doneWithFolder();
    }

    processFile(file){
        if(file.isTrashed()){
            return;
        }
        let reformatted = null;

        if(this.shouldProcess(file)){
            // do conversion
            reformatted = toGoogleFile(file);
            this.outputter.outputReformattedFile(file, reformatted);
            console.log(reformatted.getUrl());
        } else {
            //this.outputter.outputFile(file);
        }
    }

    shouldProcess(file){
        let name = file.getName();
        let dot = name.lastIndexOf(".");
        let extension = name.substring(dot + 1).toLowerCase();
        return EXTENSIONS.includes(extension);
    }
}

function iterDir(folder, fileConsumer){
    let fileIter = folder.getFiles();
    while(fileIter.hasNext()){
        fileConsumer(fileIter.next());
    }

    let dirIter = folder.getFolders();
    while(dirIter.hasNext()){
        iterDir(dirIter.next(), fileConsumer);
    }
}

//https://developers.google.com/drive/api/v2/v3versusv2
/*
While this is using the v2 of the Google Drive API instead of v3,
I couldn't find any examples of how to convert on upload with the latest version,
and have wasted enough time trying
*/
function toGoogleFile(microsoftFile){
    let parent = microsoftFile.getParents().next();
    let request = UrlFetchApp.fetch(
        "https://www.googleapis.com/upload/drive/v2/files?uploadType=media&convert=true", {
            method: "POST",
            contentType: microsoftFile.getMimeType(),
            payload: microsoftFile.getBlob().getBytes(),
            headers: {
                "Authorization" : `Bearer ${ScriptApp.getOAuthToken()}`
            },
            muteHttpExceptions: true
        }
    );
    let response = JSON.parse(request.getContentText());
    //console.log(response);
    let converted = DriveApp.getFileById(response.id); // not getId()
    converted.moveTo(parent);

    let microName = microsoftFile.getName();
    converted.setName(microName.substring(0, microName.lastIndexOf(".")));

    return converted;
}
