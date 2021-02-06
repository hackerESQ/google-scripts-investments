/** 
* maintains dividend sheet
*/

function dividends(){
  var debug = false;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var TransactionsSheet = ss.getSheetByName("Transactions");
  var DividendSheet = ss.getSheetByName("Dividends Paid");
  
  var rangeIn = TransactionsSheet.getRange("B2:M1999").getValues(); //Range of symbols to check, pulled from transactions sheet

  // creates list of stonks currently owned
  var ownedStocks = [];
  for (var symbol in rangeIn) {
    var qtyOwned = rangeIn[symbol][11]
    if (qtyOwned > 0)  {
      ownedStocks.push(rangeIn[symbol][0]);
    }
  }
  ownedStocks = arrayUnique(ownedStocks);

  if (debug) Logger.log(ownedStocks)

  var count = 0;
  var symbl;
  
  for (var i in ownedStocks) { // loop through symbols and pull info for each symbol
    
    symbl = String(ownedStocks[i]).split(":");
    symbl = symbl[1];
    //Logger.log(String(ownedStocks[i]).split(":"));
    
    var textFinder = DividendSheet.createTextFinder(ownedStocks[i]);
    var firstOccurrence = textFinder.findNext();
    //Logger.log(firstOccurrence.getA1Notation());
    var newSymbl = firstOccurrence ? false : true;
    
    var outputSize = newSymbl ? "full" : "compact";
   
    // https://www.alphavantage.co/query?function=TIME_SERIES_DAILY_ADJUSTED&symbol=FFRHX&outputsize=full&apikey=**api key**
    var keys = [
      ** api key**
      ];
    const key = keys[Math.floor(Math.random() * keys.length)];
    //var key = PropertiesService.getScriptProperties().getProperty('av_key');
    url="https://www.alphavantage.co/query?function=TIME_SERIES_DAILY_ADJUSTED&symbol="+symbl+"&outputsize="+outputSize+"&apikey="+key;
    
    if (debug) Logger.log(url)
    
    var res = UrlFetchApp.fetch(url);
    var json = JSON.parse(res);
    
    if (debug) Logger.log('symbl: '+JSON.stringify(json))

    var todayDate = new Date(date);
    var lastUpdate = new Date(DividendSheet.getParent().getRangeByName("dividends_last_update").getValue());
    
    dayArray = json["Time Series (Daily)"];

      
    if (newSymbl) { // is this a new addition to the portfolio?  if so, let's import the series..
      
      for (var date in dayArray) { 
        
        var value = dayArray[date];
        
        if ( value['7. dividend amount'] > 0) { // populate into sheet
          DividendSheet.insertRowBefore(2);
          
          DividendSheet.getRange("A2").setValue(date);
          DividendSheet.getRange("B2").setValue(ownedStocks[i]);
          DividendSheet.getRange("C2").setValue(value['7. dividend amount']);
          DividendSheet.getRange("D2").setFormula('=IF(SUMIFS(Transactions!F:F,Transactions!D:D,"Buy",Transactions!B:B,B2,Transactions!A:A,"<="&A2)-SUMIFS(Transactions!F:F,Transactions!D:D,"Sell",Transactions!B:B,B2,Transactions!A:A,"<="&A2)=0,,SUMIFS(Transactions!F:F,Transactions!D:D,"Buy",Transactions!B:B,B2,Transactions!A:A,"<="&A2)-SUMIFS(Transactions!F:F,Transactions!D:D,"Sell",Transactions!B:B,B2,Transactions!A:A,"<="&A2))');
          DividendSheet.getRange("E2").setFormula('=if(NOT(D2=""), D2*C2 ,)');
          
          count++;
        }
      }
      
    } else { // just need to update.
      
      for (var date in dayArray) {
        
        var value = dayArray[date];
        
        if ( value['7. dividend amount'] > 0 && todayDate > lastUpdate) { //check date, only update new dividends since last update
          
            DividendSheet.insertRowBefore(2);
            
            DividendSheet.getRange("A2").setValue(date);
            DividendSheet.getRange("B2").setValue(ownedStocks[i]);
            DividendSheet.getRange("C2").setValue(value['7. dividend amount']);
            DividendSheet.getRange("D2").setFormula('=IF(SUMIFS(Transactions!F:F,Transactions!D:D,"Buy",Transactions!B:B,B2,Transactions!A:A,"<="&A2)-SUMIFS(Transactions!F:F,Transactions!D:D,"Sell",Transactions!B:B,B2,Transactions!A:A,"<="&A2)=0,,SUMIFS(Transactions!F:F,Transactions!D:D,"Buy",Transactions!B:B,B2,Transactions!A:A,"<="&A2)-SUMIFS(Transactions!F:F,Transactions!D:D,"Sell",Transactions!B:B,B2,Transactions!A:A,"<="&A2))');
            DividendSheet.getRange("E2").setFormula('=if(NOT(D2=""), D2*C2 ,)');
            
            count++;
        }
      }
    }
    
    Utilities.sleep(9500); //force script to sleep so we don't exceed quotas
  }
  
  DividendSheet.getParent().getRangeByName("dividends_last_update").setValue(Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd")); 
  
  if (debug) Logger.log("Done.")
  
  //Report completion
  GmailApp.sendEmail('**email address**', '[Investments] Dividends Paid tab updated '+Utilities.formatDate(new Date(DividendSheet.getParent().getRangeByName("dividends_last_update").getValue()), "UTC", "MM/dd/yyyy"), 
                     'Report date: ' + DividendSheet.getParent().getRangeByName("dividends_last_update").getValue()+
                     '\n\nTotal dividends paid as of today: $'+formatMoney(DividendSheet.getParent().getRangeByName("total_dividends_paid_all_time").getValue())+
                     '\n\nAdded '+ count +' new dividend payments to the tab',
  {
    name: 'Investments Script',
    noReply: true
  });
  
}


/** 
* helper function returns unique array
*/
function arrayUnique(arr) {
  
  var tmp = [];
  
  // filter out duplicates
  return arr.filter(function(item, index){
    
    // convert row arrays to strings for comparison
    var stringItem = item.toString(); 
    
    // push string items into temporary arrays
    tmp.push(stringItem);
    
    // only return the first occurrence of the strings
    return tmp.indexOf(stringItem) >= index;
    
  });
}

// /** 
// * helper function removes blanks
// */
// function arrayRemoveBlanks(arr) {
  
//   var tmp = [];
  
//   return arr.filter(function(a) {
    
//     return a.filter(Boolean).length > 0;
    
//   });
// }

function formatMoney(amount, decimalCount = 2, decimal = ".", thousands = ",") {
  try {
    decimalCount = Math.abs(decimalCount);
    decimalCount = isNaN(decimalCount) ? 2 : decimalCount;

    const negativeSign = amount < 0 ? "-" : "";

    let i = parseInt(amount = Math.abs(Number(amount) || 0).toFixed(decimalCount)).toString();
    let j = (i.length > 3) ? i.length % 3 : 0;

    return negativeSign + (j ? i.substr(0, j) + thousands : '') + i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + thousands) + (decimalCount ? decimal + Math.abs(amount - i).toFixed(decimalCount).slice(2) : "");
  } catch (e) {
    console.log(e)
  }
};
