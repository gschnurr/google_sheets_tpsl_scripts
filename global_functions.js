//Global Functions
function flatten_arr(targetArr) {
  var flatArr = [];
  var row, column;

  for (row = 0; row < targetArr.length; row++) {
    for (column = 0; column < targetArr[row].length; column++) {
      flatArr.push(targetArr[row][column]);
    }
  }
  return flatArr
}

function find_col(tleColFlatArr, colToFind) {
  var colPos = 'dne';
  for (var i = 0; i < tleColFlatArr.length; i++) {
    if (tleColFlatArr[i] == colToFind) {
      var colPos = i + 1; //add one because arrays start at 0 not 1
    }
    else {
      continue;
    }
  }
  return colPos;
}

function find_row(rowFlatArr, rowToFind) {
  var rowPos = 'dne';
  for (var i = 0; i < rowFlatArr.length; i++) {
    if (rowFlatArr[i] == rowToFind) {
      var rowPos = i + 2; //add two because arrays start at 0 not 1 and title row can be ignored
    }
    else {
      continue;
    }
  }
  return rowPos;
}

function find_cell_value(sheet, colPos, valueRow) {
  var cellValue = sheet.getRange(valueRow, colPos, 1, 1).getValue();
  return cellValue;
}


function filter_rows(valueArr, filItArr, workSheet) {
  logs_tst('Row filtering has begun. The valueArr = ' + valueArr +
  '. The items to filter on are ' + filItArr +
  '. and the worksheet to complete this action in is ' + workSheet + '.');
  var arrAdj = 2;
  //arrAdj starts at 2 because the first row in the array is not the title column so we need to adjust for that
  for (var delr = 0; delr < valueArr.length; delr++) {
    if (filItArr.indexOf(valueArr[delr]) == -1) {
      var rowPos = delr + arrAdj;
      workSheet.deleteRow(rowPos);
      --arrAdj;
      logs_tst('Row # ' + rowPos + ' was filtered out. The value of that row is ' +
      valueArr[delr] + '. The loop adjuster equals ' + arrAdj + ' post-decrement.');
    }
    else {
      continue;
    }
  }
}

function filter_cols(valueArr, filItArr, workSheet) {
  logs_tst('Column filtering has begun. The valueArr = ' + valueArr +
  '. The items to filter on are ' + filItArr +
  '. and the worksheet to complete this action in is ' + workSheet + '.');
  var arrAdj = 1;
  for (var delc = 0; delc < valueArr.length; delc++) {
    if (filItArr.indexOf(valueArr[delc]) == -1) {
      var colPos = delc + arrAdj;
      workSheet.deleteColumn(colPos);
      --arrAdj;
      logs_tst('Row # ' + colPos + ' was filtered out. The value of that row is ' +
      valueArr[delc] + '. The loop adjuster equals ' + arrAdj + ' post-decrement.');
    }
    else {
      continue;
    }
  }
}

//creating two different logger wrappers one for testing so it will have all of the nuanced things that I need and one for live to report information to me.
//global variable set above that will be var loggingOnOff = ON or OFF;
//log content will be the exact content of logging this wrapper is just for each logging function to allow turning it on or off
function logs_tst(logContent) {
  if (tstloggingOnOff == 'ON') {
    Logger.log(logContent);
  }
  else {
  }
}

function logs_prd(logContent) {
  if (prdloggingOnOff == 'ON') {
    Logger.log(logContent);
  }
  else {
  }
}
