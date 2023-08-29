var SpreadSheetID = "1dWewM4NoKxmaAcJlOW9MkOU9wHEeMIxy7q7yfQEUxyg"
var PipetteSheet = "Pipettes"
var IncubatorSheet = "Incubators"


function PipetteSchedule() {
  var ss = SpreadsheetApp.openById(SpreadSheetID);
  var pipettes = ss.getSheetByName(PipetteSheet);
  var incubators = ss.getSheetByName(IncubatorSheet);

  var pip_sheet_data = getPipData(pipettes);
  var inc_sheet_data = getIncData(incubators)

  // convert dates
  // actually still appears to run even with that exception?
  // might be an issue later: Exception: The parameters (String,String,String) don't match the method signature for Utilities.formatDate.
  // https://developers.google.com/google-ads/scripts/docs/features/dates
  // https://www.reddit.com/r/GoogleAppsScript/comments/knfzs4/google_sheet_to_google_doc_template_date_format/
  
// /*
  const now = new Date();
  const MILLS_PER_DAY = 1000 * 60 * 60 * 24;
  var plus_week = new Date(now.getTime() + 6*MILLS_PER_DAY)

  verify = [];
  
  // l should be number of rows
  for(var i = 0, l= pip_sheet_data.length; i<l ; i++){    
    if (pip_sheet_data[i]['Internal Exp'] < plus_week || pip_sheet_data[i]['External Exp'] < plus_week){
      internal = pip_sheet_data[i]['Internal Exp'];
      external = pip_sheet_data[i]['External Exp'];
      pipette = pip_sheet_data[i]['Pipette ID'];
      type = pip_sheet_data[i]['Type'];
      person = pip_sheet_data[i]['User'];

      verify.push(pip_sheet_data[i]);

      if (pip_sheet_data[i].Email != "" & internal != "" & external != ""){
        MailApp.sendEmail({to: pip_sheet_data[i].Email,
                           subject: "pipette verification/calibration " + pipette,
                           htmlBody: "Please verify " + pipette + " " + type + " this week (or update the google sheet if you have verified it recently: https://docs.google.com/spreadsheets/d/1dWewM4NoKxmaAcJlOW9MkOU9wHEeMIxy7q7yfQEUxyg/edit#gid=0). Internal verification expires on " + Utilities.formatDate(internal, 'America/New_York', 'MMMM dd, yyyy') + " and the external calibration expires on " + Utilities.formatDate(external, 'America/New_York', 'MMMM dd, yyyy'),
                           noReply:true})
      }
    }
  }

  to_clean_list = [];
  
  var dataRange = incubators.getRange(2,1,incubators.getLastRow()-1, incubators.getLastColumn()).getValues();
  for(var i = 0, l= dataRange.length; i<l ; i++){
    if (inc_sheet_data[i]['next clean'] < plus_week){
      // put email stuff in here

      previously = inc_sheet_data[i]['last cleaned'];
      incubator = inc_sheet_data[i]['incubator'];
      person = inc_sheet_data[i]['initials'];
      email = inc_sheet_data[i]['email'];

      to_clean_list.push(inc_sheet_data[i]);

      if (email != ""){
        MailApp.sendEmail({to: email, subject: "clean incubator " + incubator, htmlBody: "Please clean incubator " + incubator + " this week (or update the google sheet if you have cleaned it this month: https://docs.google.com/spreadsheets/d/1Yk6gTs7ETao4AJUHReb5GiVjgEeNJR2DSid0Emn-Yz4/edit#gid=0). It was last cleaned: " + Utilities.formatDate(previously, 'America/New_York', 'MMMM dd, yyyy'), noReply:true})
      }      
    }
  }

  // //making sure I am accessing the entire table
  // first = pip_obj[0];
  // last = pip_obj[Object.keys(pip_obj).length-1];
  // console.log(first,last);

  if (Object.keys(verify).length != 0){
      MailApp.sendEmail({to: EMAIL,
                         subject: "Equipment Maintenance",
                         htmlBody: printPip(verify) + printInc(to_clean_list),
                         noReply:true});
  }

  else{
    MailApp.sendEmail({to: EMAIL,
                       subject: "Hooray no pipettes to verify",
                       htmlBody: "",
                       noReply:true});
  }

}

function getPipData(pipette_data){
  var dataArray = [];
  var rows = pipette_data.getRange(3,1,pipette_data.getLastRow()-1, pipette_data.getLastColumn()).getValues();

  for(var i = 0, l= rows.length; i<l ; i++){
    if (rows[i][0] !== ''){
      var dataRow = rows[i];
      var record = {};
      record['Email'] = dataRow[0];
      record['User'] = dataRow[1];
      record['Pipette ID'] = dataRow[2];
      record['Type'] = dataRow[3];
      record['Internal Ver'] = dataRow[4];
      record['Internal Exp'] = dataRow[5];
      record['External Cal'] = dataRow[6];
      record['External Exp'] = dataRow[7];
      dataArray.push(record);
    }
  }
  return dataArray;
}

function getIncData(incubator_schedule){
  var dataArray = [];
// collecting data from 2nd Row , 1st column to last row and last    // column sheet.getLastRow()-1
  var rows = incubator_schedule.getRange(2,1,incubator_schedule.getLastRow()-1, incubator_schedule.getLastColumn()).getValues();

  for(var i = 0, l= rows.length; i<l ; i++){
    var dataRow = rows[i];
    var record = {};
    record['incubator'] = dataRow[0];
    record['initials'] = dataRow[1];
    record['email'] = dataRow[2];
    record['last cleaned'] = dataRow[3];
    record['next clean'] = dataRow[4];
    dataArray.push(record);
  }
  return dataArray;
}

function printPip(pipettes){
  string = "<html><body><br><table border=1><tr><th>Person</th><th>Pipette ID</th><th>Pipette Type</th><th>Verification Expiration</th><th>Calibration Expiration</th></tr></br>";
  for (var i=0; i<pipettes.length; i++){
    string = string + "<tr>";

    temp = `<td> ${pipettes[i]['User']} </td><td> ${pipettes[i]['Pipette ID']}  </td><td> ${pipettes[i]['Type']} </td><td> ${Utilities.formatDate(pipettes[i]['Internal Exp'], 'America/New_York', 'MMMM dd, yyyy')}</td><td> ${Utilities.formatDate(pipettes[i]['External Exp'], 'America/New_York', 'MMMM dd, yyyy')}</td>`;

    string = string.concat(temp);
    string = string + "</tr>";
  }
  string = string + "</table></body></html>";
  return string;
}

function printInc(incs){
  string = "<html><body><br><table border=1><tr><th>Person</th><th>Incubator</th><th>Next Clean Date</th></tr></br>";
  for (var i=0; i<incs.length; i++){
    string = string + "<tr>";

    temp = `<td> ${incs[i]['initials']} </td><td> ${incs[i]['incubator']}  </td><td> ${Utilities.formatDate(incs[i]['next clean'], 'America/New_York', 'MMMM dd, yyyy')}</td>`;
    string = string.concat(temp);
    
    string = string + "</tr>";
  }
  string = string + "</table></body></html>";
  return string;
}
