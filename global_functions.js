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
  var colPos;
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
