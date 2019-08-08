//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//STORAGE ARRAYS
//////////////////////////////////////////////////////////////////////////////////////////////////////////////


var SEARCH_QUERY = "to:me newer_than:5d";
var subjects = [];
var bodys = [];
var dates = [];
var froms = [];
var names = [];
var emailPath = [];
var hasAttachments = [];










//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//CREATES MENU ITEM WHEN DOCUMENT IS OPENED AND FULLY LOADED
//////////////////////////////////////////////////////////////////////////////////////////////////////////////


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('GMAIL APP DATA')
      .addItem('Fetch Data', 'menuItem1')
      .addToUi();
}













//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//RUNS THE FULL CODE WHEN THE MENU IS PRESSED
//////////////////////////////////////////////////////////////////////////////////////////////////////////////


function menuItem1() {
  SpreadsheetApp.getActiveSheet().clear();
  
  
  addHeader(SpreadsheetApp.getActiveSheet());
  getInfo_(SEARCH_QUERY);
  
  if (this.froms.length != 0) {
    appendData_(SpreadsheetApp.getActiveSheet(), this.froms, this.subjects, this.bodys, this.names, this.dates, this.emailPath, this.hasAttachments);
    } 
}













//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//ADDS HEADER INFO - IF THERE IS NO HEADER
//////////////////////////////////////////////////////////////////////////////////////////////////////////////


function addHeader(Sheet1){
  Sheet1.setRowHeight(1, 100)
  var i = 1;
  var headers = ['From','Subject','Body','Name','Date','URL Reply','Has Attachment'];
  
  if(!Sheet1.getLastRow()){
  while( i < 8){
    if( i === 1){
        Sheet1.getRange(Sheet1.getLastRow()+1, i).setValue(headers[0]);
        Sheet1.getRange(Sheet1.getLastRow(), i).setBackground('gray');
        Sheet1.getRange(Sheet1.getLastRow(), i).setFontColor('white');
        Sheet1.getRange(Sheet1.getLastRow(), i).setFontSize(28);
        i++;
    }else {
        Sheet1.getRange(Sheet1.getLastRow(), i).setValue(headers[i-1]);
        Sheet1.getRange(Sheet1.getLastRow(), i).setBackground('gray');
        Sheet1.getRange(Sheet1.getLastRow(), i).setFontColor('white');
        Sheet1.getRange(Sheet1.getLastRow(), i).setFontSize(28);
        i++;
         }
    };
  };
}










//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//GETS INFO FROM EMAIL AND STORES IN ARRAYS
//////////////////////////////////////////////////////////////////////////////////////////////////////////////


 
function getInfo_(q) {
  var threads = GmailApp.search(q);
    for (var i in threads) {
        var msgs = threads[i].getMessages();
       for (var j in msgs) {
         if(!msgs[j].isStarred()){
           
           var message = msgs[j];

           this.bodys.push([message.getPlainBody().replace(/<.*?>/g, '').replace(/^\s*\n/gm, '').replace(/^\s*/gm, '').replace(/\s*\n/gm, '\n')]);
           this.subjects.push([message.getSubject().replace(/<.*?>/g, '\n').replace(/^\s*\n/gm, '').replace(/^\s*/gm, '').replace(/\s*\n/gm, '\n')]);
           this.dates.push([message.getDate()]);
           this.names.push([message.getFrom().replace(/<.*?>/g, '').replace(/^\s*\n/gm, '').replace(/^\s*/gm, '').replace(/\s*\n/gm, '\n')]);
           this.froms.push([message.getFrom().replace(/^.+</, '').replace(">", '') ]);
           this.emailPath.push(['https://mail.google.com/mail/u/0/#inbox/' + message.getId()]);
           message.star();
           
           if(msgs[j].getAttachments()){
             this.hasAttachments.push([message.getAttachments()])
           }else{
             this.hasAttachments.push([" "])
           };  
         }
       }
    }
}











//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//APPENDS ARRAY INFO TO THE SPREADSHEET
//////////////////////////////////////////////////////////////////////////////////////////////////////////////


function appendData_(Sheet1, array2dFroms, array2dSubjects, array2dBodys, array2dNames, array2dDates, array2dURL, array2dAttachments) {
  Sheet1.getRange(Sheet1.getLastRow() + 1, 1, array2dFroms.length, array2dFroms[0].length).setValues(array2dFroms);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 2, array2dSubjects.length, array2dSubjects[0].length).setValues(array2dSubjects);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 3, array2dBodys.length, array2dBodys[0].length).setValues(array2dBodys);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 4, array2dNames.length, array2dNames[0].length).setValues(array2dNames);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 5, array2dDates.length, array2dDates[0].length).setValues(array2dDates);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 6, array2dURL.length, array2dURL[0].length).setValues(array2dURL);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 7, array2dAttachments.length, array2dAttachments[0].length).setValues(array2dAttachments);
      
    }
  




