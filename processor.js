/*
The FileProcessor class is used to process files and folders
*/

const EXTENSIONS = ["doc", "docx", "xls", "xlsx", "ppt", "pptx"].map((e)=>e.toLowerCase());
const OUTPUT_ALL = false; // change this to true if you want it to output all files it finds,
// rather than just the ones it will convert

class FileProcessor {

    constructor(outputter, folderQueue, folderStack, doConversion){
        this.outputter = outputter;
        this.folderQueue = folderQueue;
        this.folderStack = folderStack;
        this.doConversion = doConversion;
    }

    run(){
        this.folderQueue.insertHeader();

        // move stack to queue
        while(!this.folderStack.isEmpty()){
          this.folderQueue.enqueue(this.folderStack.peek());
          this.folderStack.pop();
        }

        while(!this.folderQueue.isEmpty()){
            this.processNext();
        }
    }

    processNext(){
        let folder = this.folderQueue.getNextFolder();
        this.outputter.outputFolder(folder);

        let fileIter = folder.getFiles();
        while(fileIter.hasNext()){
            this.processFile(fileIter.next());
        }

        let dirIter = folder.getFolders();
        let numSubdirs = 0;
        while(dirIter.hasNext()){
            this.folderStack.push(dirIter.next().getId());
            ++numSubdirs;
        }
        this.folderQueue.dequeue(); // mark this folder as done

        // process the folders I just pushed
        while(numSubdirs > 0){
            this.folderQueue.pushToFront(this.folderStack.peek());
            this.folderStack.pop();
            --numSubdirs;
            this.processNext();
        }

        this.outputter.doneWithFolder();
    }

    processFile(file){
        if(file.isTrashed()){
            return;
        }
        let reformatted = null;
        let proc = this.shouldProcess(file);
        if(this.doConversion && proc){
            reformatted = toGoogleFile(file);
            this.outputter.outputReformattedFile(file, reformatted);
        } else if(proc || OUTPUT_ALL){
            this.outputter.outputFile(file);
        }
    }

    shouldProcess(file){
        let name = file.getName();
        let dot = name.lastIndexOf(".");
        let extension = name.substring(dot + 1).toLowerCase();
        return EXTENSIONS.includes(extension);
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
