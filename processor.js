/*
The FileProcessor class is used to process files and folders
*/

const EXTENSIONS = ["doc", "docx", "xls", "xlsx", "ppt", "pptx"].map((e)=>e.toLowerCase());

class FileProcessor {

    constructor(outputter, folderQueue, folderStack, doConversion){
        this.outputter = outputter;
        this.folderQueue = folderQueue;
        this.folderStack = folderStack;
        this.doConversion = doConversion;
    }

    run(){
        // TODO: move stack to queue

        this.folderQueue.insertHeader();
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
        while(dirIter.hasNext()){
            this.folderStack.push(dirIter.next().getId());
        }
        this.folderQueue.dequeue(); // mark this folder as done

        // process the folders I just pushed
        while(!this.folderStack.isEmpty()){
            this.folderQueue.pushToFront(this.folderStack.peek());
            this.folderStack.pop();
            this.processNext();
        }

        this.outputter.doneWithFolder();
    }

    processFile(file){
        if(file.isTrashed()){
            return;
        }
        let reformatted = null;

        if(this.doConversion && this.shouldProcess(file)){
            reformatted = toGoogleFile(file);
            this.outputter.outputReformattedFile(file, reformatted);
        } else {
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
