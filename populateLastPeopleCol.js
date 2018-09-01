//----------------------------------------------Populate last people column-----------------------
//global variables
var projectSheetStartingDateCol = 7; //dates start @ column H or column #7. change as needed

function populateLastPeopleCol() {
  var masterProjectObjArray = [];
  var lastPeopleColDate = peopleSheet.getRange(1, peopleSheet.getLastColumn()).getValue();
  var projectSheetArr = [projSheet, storyboothProjectsSheet];
  
  //populate master project object array
  projectSheetArr.forEach(function(sheet){
    addCellValsToMasterArr(sheet,masterProjectObjArray,lastPeopleColDate);
  });
  
  fillLastPeopleCol(masterProjectObjArray);
}

function fillLastPeopleCol(masterProjectObjArray) {
  // ui.alert('function entered');
  var lastPeopleColRange = peopleSheet.getSheetValues(2, peopleSheet.getLastColumn(), peopleSheet.getLastRow(), 1);
  //for (var i = 2; i <= lastPeopleColRange.length; i++) {
  for (var i = lastPeopleColRange.length; i > 1; i--) {
    var curPersonNameArr = peopleSheet.getRange(i, 1).getValue().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g,"").split(" ");
    var curPersonCell = peopleSheet.getRange(i, peopleSheet.getLastColumn());
    for (var j = 0; j < masterProjectObjArray.length; j++) {
      var projObj = masterProjectObjArray[j];
      var preppedString = projObj.val.toString().replace(",","")
      curPersonNameArr.forEach(function(curPersonName) {
       if (curPersonName.length == 1) {
          curPersonName = curPersonName.substring(1);
        }
        if (curPersonName && preppedString.indexOf(curPersonName) > -1) {
          var currentName = curPersonName;
          curPersonCell.setValue(projObj.projName);
          if (projObj.projName.indexOf("Storybooth") >= 0 || projObj.projName.indexOf("storybooth") >= 0) { //todo: optimize this
            curPersonCell.setBackground("#d9d2e9");
          } else {
            curPersonCell.setBackground(projObj.color);
          }
        }
      });
    }
  }
}

function addCellValsToMasterArr(sheet,masterProjectObjArray,lastPeopleColDate) {
  projectSheetDateRange = sheet.getSheetValues(1, projectSheetStartingDateCol, 1, sheet.getLastColumn());
  projectSheetDateArr = projectSheetDateRange[0].filter(function(date) {
    return isDate(date);
  });
  for (var numDateCols = projectSheetDateArr.length - 1; numDateCols > 0; numDateCols--) { //start @end because you will always return date sooner
    var date = projectSheetDateArr[numDateCols];
    if (typeof(date) !== 'undefined' && (date.toString() == lastPeopleColDate.toString())) { //unsure why string conversion is necessary
      relevantProjColumnRange = sheet.getSheetValues(1, numDateCols + projectSheetStartingDateCol, sheet.getMaxRows(), 1);  //TODO: make LastRow, not MaxRows
      //if someone adds 5000 rows this'll blow up
      for (var rowNum = 2; rowNum < relevantProjColumnRange.length; rowNum++) {
        var cellValue = relevantProjColumnRange[rowNum -1][0];
        var cellObj = {
          val: '',
          projName: '',
          color: '',
        };
        if (cellValue != "") { 
          cellObj.val = cellValue;
          cellObj.projName = sheet.getRange(rowNum,2).getValue();
          cellObj.color = sheet.getRange(rowNum,2).getBackground();
          masterProjectObjArray.push(cellObj);
        }      
      }    
      break;
    }
  }
}

function isDate (value) {
  return value instanceof Date;
}