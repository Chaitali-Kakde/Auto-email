# Auto-email
This is an app script code to auto send tables in email body daily on scheduled time which copies content from google sheet and paste it into email body. It will save time and emails can be shared automatically without human intervention.  
1. Create Google sheet and add table which you want to share in email daily.
2. copy below code into app script and run.
3. Make changes in referance and sheets name as required.

function SendMail(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Shrinkage Report");
  var EmailID = sheet.getRange("O3").getValue();
  var Name = sheet.getRange("O4").getValue();

  var dataRange = sheet.getRange("B14:F20");
  var data = dataRange.getValues();
  var colors = dataRange.getBackgrounds();
  var formats = dataRange.getNumberFormats();
  //var colors3 = dataRange3.getBackgrounds();
  //var formats3 = dataRange3.getNumberFormats();


  var table1 = "<table style='border: 1px solid black; border-collapse: collapse; text-align: center;' cellpadding='5'>";
  //var table3 = "<table style='border: 1px solid black; border-collapse: collapse; text-align: center;' cellpadding='5'>";
  
  for (var i = 0; i < data.length; i++) {
    var cells = data[i];
    table1 += "<tr>";
    
    for (var u = 0; u < cells.length; u++) {
      var cellColor = colors[i][u];
      var format = formats[i][u];
      var finalValue;

      if (format && format.includes("$")) {
        finalValue = "$" + cells[u];
      } else if (format && format.includes("%")) {
        finalValue = (parseFloat(cells[u]) * 100).toFixed(2) + " %";
      } else {
        finalValue = cells[u];
      }

      table1 += "<td style='background-color:" + cellColor + "; border: 1px solid black;'>" + finalValue + "</td>";
      //table3 += "<td style='background-color:" + cellColor + "; border: 1px solid black;'>" + finalValue + "</td>";
    }
    
    table1 += "</tr>";
    //table3 += "</tr>"
  }
  
  table1 += "</table>";
  

  var dataRange2 = sheet.getRange("B33:Q42");
  var data2 = dataRange2.getValues();
  var table2 = "<table style='border: 1px solid black; border-collapse: collapse; text-align: center;' cellpadding='5'>";
  var colors2 = dataRange2.getBackgrounds();
  var formats2 = dataRange2.getNumberFormats();

  for (var j = 0; j < data2.length; j++) {
    var cells2 = data2[j];
    table2 += "<tr>";

    for (var k = 0; k < cells2.length; k++) {
      var cellColor2 = colors2[j][k];
      var format = formats2[j][k];
      var finalValue2;

      if (format && format.includes("$")) {
        finalValue = "$" + cells2[k];
      } else if (format && format.includes("%")) {
        finalValue2 = (parseFloat(cells2[k]) * 100).toFixed(2) + " %";
      } else {
        finalValue2 = cells2[k];
      }

      table2 += "<td style='background-color:" + cellColor2 + "; border: 1px solid black;'>"  + finalValue2 + "</td>";
    }
    
    table2 += "</tr>";
  }
  
  table2 += "</table>";

  // Table 3

  var dataRange3 = sheet.getRange("I21:M30");
  var data3 = dataRange3.getValues();
  var table3 = "<table style='border: 1px solid black; border-collapse: collapse; text-align: center;' cellpadding='5'>";
  var colors3 = dataRange3.getBackgrounds();
  var formats3 = dataRange3.getNumberFormats();

  for (var m = 0; m < data3.length; m++) {
    var cells3 = data3[m];
    table3 += "<tr>";

    for (var n = 0; n < cells3.length; n++) {
      var cellColor3 = colors3[m][n];
      var format = formats3[m][n];
      var finalValue3;

      if (format && format.includes("$")) {
        finalValue = "$" + cells3[n];
      } else if (format && format.includes("%")) {
        finalValue3 = (parseFloat(cells3[n]) * 100).toFixed(2) + " %";
      } else {
        finalValue3 = cells3[n];
      }

      table3 += "<td style='background-color:" + cellColor3 + "; border: 1px solid black;'>"  + finalValue3 + "</td>";
    }
    
    table3 += "</tr>";
  }
  
  table3 += "</table>";

  // end of table 3

  var htmlBody = "Hello All," + "<br><br>Please find today’s attendance details (EA/PRA shrinkage summary and EA journal-wise shrinkage summary) as mentioned below:<br><br>" + "1. Attendance<br><br>" + table2 + "<br><br>2.	EA/PRA-wise shrinkage summary<br><br>" + table1 + "<br><br> 3. EA/PRA journal-wise shrinkage summary<br><br>" + table3 + "<br><br>Regards,"+ "<br><br>Chaitali";

  MailApp.sendEmail({
    to: EmailID,
    cc : "chaitalikakde210@gmail.com",
    subject: Name,
    htmlBody: htmlBody,
  });
}

