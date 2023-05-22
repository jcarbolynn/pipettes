var SpreadSheetID = "1LwApcvc1Jizm6LMIGWEqhTPhK0xesCpn70-Z-YSAtNU"
var SheetName = "Pipettes"


function PipetteSchedule() {
  var ss = SpreadsheetApp.openById(SpreadSheetID);
  var pipettes = ss.getSheetByName(SheetName);

  var pip_obj = getData(pipettes);

  // convert dates
  // actually still appears to run even with that exception?
  // might be an issue later: Exception: The parameters (String,String,String) don't match the method signature for Utilities.formatDate.
  
// /*
  const now = new Date();
  const MILLS_PER_DAY = 1000 * 60 * 60 * 24;
  var plus_week = new Date(now.getTime() + 6*MILLS_PER_DAY)
  
  // l should be number of rows
  for(var i = 0, l= pip_obj.length; i<l ; i++){
    if (pip_obj[i]['Internal Exp'] < plus_week || pip_obj[i]['External Exp'] < plus_week){
      internal = pip_obj[i]['Internal Exp'];
      external = pip_obj[i]['External Exp'];
      pipette = pip_obj[i]['Name'];
      person = pip_obj[i]['User'];

      MailApp.sendEmail({to: pip_obj[i].Email, subject: "pipette verification/calibration " + pipette, htmlBody: "Please verify " + pipette + " this week (or update the google sheet if you have verified it recently: https://docs.google.com/spreadsheets/d/1LwApcvc1Jizm6LMIGWEqhTPhK0xesCpn70-Z-YSAtNU/edit#gid=110014716). Internal verification expires: " + Utilities.formatDate(internal, 'America/New_York', 'MMMM dd, yyyy') + " and the external calibration expires:  " + Utilities.formatDate(external, 'America/New_York', 'MMMM dd, yyyy'), noReply:true})
    }
  }
// */

}

function getData(pipette_data){
  var dataArray = [];
  var rows = pipette_data.getRange(4,1,pipette_data.getLastRow()-1, pipette_data.getLastColumn()).getValues();

  for(var i = 0, l= rows.length; i<l ; i++){
    var dataRow = rows[i];
    var record = {};
    record['Email'] = dataRow[0];
    record['User'] = dataRow[1];
    record['Name'] = dataRow[2];
    record['Type'] = dataRow[3];
    record['Internal Ver'] = dataRow[4];
    record['Internal Exp'] = dataRow[5];
    record['External Exp'] = dataRow[6];
    record['External Cal.'] = dataRow[7];
    dataArray.push(record);
  }
  return dataArray;
}


