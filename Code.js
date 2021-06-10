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

Essentially, I need to figure out how to process the folders pre-order while also
maintaining the indentation output given by the outputter and keeping in mind
Google Script's limitations
*/

// stack and queue must each have their own sheet, otherwise getLastRow() will not work
// (when one is larger than the other)
function onOpen(){
    SpreadsheetApp
    .getUi()
    .createMenu("File Converter")
    .addItem("create resources", "createResources")
    .addItem("run", "run")
    .addItem("Test run", "testRun")
    .addToUi();
}

function createResources(){
    getOrCreateFolderQueue();
}

function run(){
    doRun(true);
}

function testRun(){
    doRun(false);
}

function doRun(doConvert){
    let outputSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    let processor = new FileProcessor(
        new Outputter(outputSheet),
        getOrCreateFolderQueue(),
        getOrCreateFolderStack(),
        doConvert
    );
    processor.run();
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
