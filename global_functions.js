
//first global function
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

//test showing how it should work
function test_flatten_arr() {
  var findGdprColumnArr = tpsl.getRange(2, 1, 1, tpslLc).getValues();
  var tcaOned = flatten_arr(findGdprColumnArr);
  SpreadsheetApp.getUi().alert(findGdprColumnArr[0]);
  SpreadsheetApp.getUi().alert(tcaOned[0]);
}
