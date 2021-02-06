function _col(named_range) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // return "Col" + ss.getRangeByName(named_range).getColumn();
  var regex = /([A-Za-z])(?=.+\1)/;
  var range = ss.getRangeByName(named_range).getA1Notation();

  return range.match(regex)
}

function _numOnly(val) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // return "Col" + ss.getRangeByName(named_range).getColumn();
  var regex = /[0-9]+/;

  return val.match(regex)
}
