function takeDailySnapshot() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rngCopy = ss.getRangeByName("StockVolume_Daily!rng_DailyCopy");
  var rngPaste = ss.getRangeByName("StockVolume_Daily!rng_DailyPaste");
  var intNextCol = ss.getRangeByName("StockVolume_Daily!cell_NextCol").getValue();
  
  rngPaste.offset(0, intNextCol+1).setValues(rngCopy.getValues());
  
  
}