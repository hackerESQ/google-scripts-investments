function daily() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var DailyChangeSheet = ss.getSheetByName("Daily Change");
  var DashboardSheet = ss.getSheetByName("Dashboard");
  
  //Update Daily Change Sheet?
  var day = new Date();
  if (day.getDay()==0 || day.getDay()==6) { //Skip week-end
    return;
  }
  
  // add new formula sheet (creates new 2nd row)
  DailyChangeSheet.insertRowBefore(2);
  DailyChangeSheet.getRange("A2").setFormula('=today()'); // calc new date
  DailyChangeSheet.getRange("B2").setFormula('=Dashboard!total_current_value'); // daily value
  DailyChangeSheet.getRange("C2").setFormula('=Dashboard!total_dollars_gain_loss'); // daily change
  DailyChangeSheet.getRange("D2").setFormula('=A2'); // day of week
  DailyChangeSheet.getRange("E2").setFormula('=total_cost_basis'); 
  DailyChangeSheet.getRange("F2").setFormula('=day_dollars_gain_loss'); 
  DailyChangeSheet.getRange("G2").setFormula('=total_dividends_paid_all_time'); 
  DailyChangeSheet.getRange("H2").setFormula('=sumif(transactions.type,"Sell",transactions.gain_loss)'); 
  DailyChangeSheet.getRange("I2").setFormula('=E2-E3'); 
  
  // freezes previous first row of data (copies 3rd row without formulas)
  var freeze = DailyChangeSheet.getRange("3:3"); 
  freeze.copyTo(freeze,{contentsOnly:true});  
  
  // Update Chart Ranges
  var DateRange = DailyChangeSheet.getRange("A1:A30")
  var DailyGainRange = DailyChangeSheet.getRange("C1:C30")
  
  var chart = DashboardSheet.getCharts()[0];
  
  chart = chart.modify()
  .clearRanges()
  .addRange(DateRange)
  .addRange(DailyGainRange)
  .setOption('title', 'Daily Gains/Losses (Last 30 days)')
  .setOption('animation.duration', 500)
  .build();
  DashboardSheet.updateChart(chart);
  
  //Report completion
  GmailApp.sendEmail('corey@coreyvarma.com', '[Investments] Daily Change tab updated '+Utilities.formatDate(new Date(DailyChangeSheet.getRange("A2").getValue()), "UTC", "MM/dd/yyyy"), 
                     'Report date: ' + DailyChangeSheet.getRange("A2").getValue()+
                     '\n\nThe daily change today was: $'+formatMoney(DailyChangeSheet.getRange("F2").getValue())+
                     '\n\nTotal cost basis today is: $'+formatMoney(DailyChangeSheet.getRange("E2").getValue())+
                     '\n\nTotal market value today is: $'+formatMoney(DailyChangeSheet.getRange("B2").getValue()), 
  {
    name: 'Investments Script',
    noReply: true
  });
}

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
