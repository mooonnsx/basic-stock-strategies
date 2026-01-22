function mainSellHigh() {
  const apikey = "YOUR_API_KEY";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  var startdate = "2025-07-21"; // Date to begin, goes to current date
  var stock = "tsla"; // NYSE Ticker Symbol
  var interval = "5min"; // Supported intervals: 1min, 5min, 15min, 30min, 45min, 1h, 2h, 4h, 8h, 1day, 1week, 1month
  var percent = 1.5; // Percent difference from timerange that changes if model buys or sells
  var timerange = 10; // amount of intervals to look back 
  addStockValues(apikey, stock, startdate, interval, sheet);
  var sourcerange = sheet.getRange(`B2:B${2+timerange}`); // adds data to sheet that models the first x amount of days 
  var startcell = sheet.getRange("C2");
  sourcerange.copyTo(startcell);
  sellHighStrategy(sheet, percent, timerange);
}

function mainMomentumAcceleration() {
  const apikey = "YOUR_API_KEY";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  var startdate = "2025-07-21"; // Date to begin, goes to current date
  var stock = "wmt"; // NYSE Ticker Symbol
  var interval = "1min"; // Supported intervals: 1min, 5min, 15min, 30min, 45min, 1h, 2h, 4h, 8h, 1day, 1week, 1month
  var shortMA = 2; // length of short moving average based on interval
  var longMA = 20; // length of long moving average based on interval
  addStockValues(apikey, stock, startdate, interval, sheet);
  var sourcerange = sheet.getRange(`B2:B${2+longMA}`); // adds data to sheet that models the first x amount of days 
  var startcell = sheet.getRange("C2");
  sourcerange.copyTo(startcell);
  momentumAccelerationStrategy(sheet, shortMA, longMA);
}

function addStockValues(apikey, stocksymbol, startdate, interval, thissheet) { // brings stock values from twelvedata.com's api and adds it to first 2 rows of the stock. NOTE: capped at 5000 rows of data with free plan
  const sheet = thissheet;
  let url = `https://api.twelvedata.com/time_series?symbol=${stocksymbol}&interval=${interval}&start_date=${startdate}&apikey=${apikey}`; 
  const getJson = UrlFetchApp.fetch(url);
  const parse = JSON.parse(getJson.getContentText());
  var length = parse.values.length;
  sheet.getRange(1,1).setValue("time");
  sheet.getRange(1,2).setValue(stocksymbol);
  for (var i = 0; i<=length-1; i++) {
    let value = parse.values[i].close;
    let datetime = parse.values[i].datetime;
    sheet.getRange(length-i+1,2).setValue(value);
    sheet.getRange(length-i+1,1).setValue(datetime);
    console.log(value);
  }
}

function momentumAccelerationStrategy(mysheet, shortMAlength, longMAlength) { // looks at the short term and long term slope. if the short term slope is positive and greater than long term, it invests. otherwise, it pulls out.
  var sheet = mysheet;
  var range = sheet.getDataRange().getValues();
  this.shortMAlength = shortMAlength;
  this.longMAlength = longMAlength;
  var shortMA;
  var longMA;
  let invested = true;
  range[0][2] = "slopecomp";
  for (var i = longMAlength; i < range.length; i++) {
    if (invested) {
      range[i][2] = range[i][1] / range[i-1][1] * range[i-1][2];
    } else {
      range[i][2] = range[i-1][2];
    }
    if (shortMA > longMA && shortMA > 0) {
      invested = true;
    } else {
      invested = false;
    }
    shortMA = (range[i][1] - range[i-shortMAlength][1]) / shortMAlength; // reset values
    longMA = (range[i][1] - range[i-longMAlength][1]) / longMAlength;
  }
  sheet.getRange(1, 1, range.length, range[0].length).setValues(range); // add to sheet
}

function sellHighStrategy(mysheet, percent, timerange) { // sells after specified increase after specified time period, buys after same decrease in same time period
  var sheet = mysheet;
  var range = sheet.getDataRange().getValues();
  var value = range[1][1];
  range[0][2] = "sellhigh";
  var invested = true;
  for (var i = timerange; i < range.length; i++) {
    if (invested == true) {
      range[i][2] = range[i][1] / range[i-1][1] * range[i-1][2];
    }
    if (invested == false) {
      range[i][2] = range[i-1][2];
    }
    if (range[i][1] >= (1+percent/100)*value && invested == true) {
      invested = false;
      lasttrade = 0;
    } else if (range[i][1] <= (1-percent/100)*value && invested == false) {
      invested = true;
      lasttrade = 0;
    }
    value = range[i-timerange][1];
  }
  console.log(range);
  sheet.getRange(1, 1, range.length, range[0].length).setValues(range); // add to sheet
}


