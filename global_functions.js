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
