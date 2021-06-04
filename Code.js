const PROCESSING_LIST_NAME = "folders to process";
const EXTENSIONS = ["doc", "docx", "xls", "xlsx", "ppt", "pptx"].map((e)=>e.toLowerCase());

// I would like to keep the folder structure display the old version had

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

class Outputter {
  constructor(sheet){
    this.sheet = sheet;
    this.row = 1;
    this.col = 1;
  }

  outputFolder(folder){
    this.sheet.getRange(this.row, this.col).setValue(folder.getName());
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

  outputReformattedFile(orig, reform){
    this.sheet.getRange(this.row, this.col, 1, 3).setValues([
      [orig.getName(), "--->", reform.getName()]
    ]);
    this.row++;
  }
}

function testGs(){
  let f = DriveApp.getFileById("12oIPncVXV1dSjnYIB1UDOLZ51fWf-oU1");
  toGoogleSheet(f);
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