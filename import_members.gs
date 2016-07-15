function getNewestFileInFolder() {

  var folder = DriveApp.getFolderById('0BwvK5gYQ6D4nOXB2UHFUV0w5WnM');  
  var arryFileDates = [];
  var objFilesByDate = {};

  var files = folder.getFiles();
  var fileDate = "";

  while (files.hasNext()) {
    var file = files.next();
    Logger.log('xxxx: file date: ' + file.getLastUpdated());
    Logger.log('xxxx: file name: ' + file.getName());
    Logger.log(" ");

    fileDate = file.getLastUpdated();
    objFilesByDate[fileDate] = file.getId(); //Create an object of file names by file ID

    arryFileDates.push(file.getLastUpdated());
  }
  arryFileDates.sort(function(a,b){return b-a});

  Logger.log(arryFileDates);

  var newestDate = arryFileDates[0];
  Logger.log('Newest date is: ' + newestDate);

  var newestFileID = objFilesByDate[newestDate];
  Logger.log('newestFile: ' + newestFileID);
    
  var newestFile = DriveApp.getFileById(newestFileID);
  return newestFile;
};

function filterArray(values, chapterName) {
  return values.filter(function(d) {
    return d[2] == chapterName;
  });
}

function getChapterMembers_(chapterName) {
  var file = getNewestFileInFolder();
  var csvFile = file.getBlob().getDataAsString();
  var csvData = CSVToArray_(csvFile);
  var csvDataFiltered = filterArray(csvData, chapterName);
  return csvDataFiltered;
}

function CSVToArray_( strData ){
  var rows = strData.split("\n");
  Logger.log(rows.length);
  var array = [];
  for(n=0;n<rows.length;++n){
    if(rows[n].split(',').length>1){ 
      array.push(rows[n].split(','));
    }
  }
  Logger.log(array);
  return array;
}