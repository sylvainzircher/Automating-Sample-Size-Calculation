function onOpen() {
  ui = SpreadsheetApp.getUi();
  
  ui.createMenu("New Menu")
  .addItem("Sample Size", "sampleSizes") 
  .addToUi();  
}

function sampleSizes() {
  var ss = SpreadsheetApp.getActive();
  
  /** Please replace below the "Automated" by the name you picked for the Sheet **/
  var s = ss.getSheetByName("Automated");
  /*******************************************************************************/
  
  /** Some Specifications **/
  // Make sure you abide to the specific setup of the sheets for the Macro to work
  // Alpha should be in cell B1
  // Power should be in cell B2
  // The baseline conversion rate in cell B3
  // The historical traffic in cell B4
  // The different combinations of the number of variants and Minimum Detectable Effects should start from A8:B8
    /*******************************************************************************/
  
  var lastRow = s.getLastRow() - 6;
  var lastCol = s.getLastColumn();
  var variants = s.getRange(7, 1, lastRow, 1);
  var mdes = s.getRange(7, 2, lastRow, 1);  
  var alpha = s.getRange("B1").getValue();
  var oneMinusbeta = s.getRange("B2").getValue();
  var baseline = s.getRange("B3").getValue();  
  var traffic = s.getRange("B4").getValue();
  
  s.getRange("C1").setFormula("=NORMINV(1-B1/2,0,1)");
  var t_alpha2 = s.getRange("C1").getValue();
  s.getRange("C1").setValue("");
  
  s.getRange("C2").setFormula("=NORMINV(B2,0,1)");  
  var t_beta = s.getRange("C2").getValue();
  s.getRange("C2").setValue("");  
  
  var p1 = baseline;
  
  var samples = [];
  var days = [];

//  samples = (t_alpha2 + t_beta)^2 * (p1 * (1 - p1) + p2 * (1 - p2)) / delta^2
  for (var i = 2; i <= lastRow; i++) {
    var mde = mdes.getCell(i,1).getValue();
    var variant = variants.getCell(i,1).getValue();
    var p2 = p1 * (1 + mde);
    var sample = (t_alpha2 + t_beta) * (t_alpha2 + t_beta) * (p1 * (1 - p1) + p2 * (1 - p2)) / ((p2 - p1) * (p2 - p1));
    samples.push([sample]);
    var day = sample * (variant + 1) / traffic;
    days.push([day]);
  }
  s.getRange("C7").setValue("Size for one Variant");
  s.getRange(8, 3, lastRow - 1, 1).setValues(samples);
  
  s.getRange("D7").setValue("Days");
  s.getRange(8, 4, lastRow - 1, 1).setValues(days);  
}
