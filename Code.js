function onOpen() {
  createMenu();
  scrapeNBAData();
}

// Creates a menu button called "Update Table"
function createMenu() {
  SpreadsheetApp.getUi().createMenu("Update Table")
    .addItem("Update Table", "UpdateTable")
    .addToUi();
}

function UpdateTable() {
  var STARTROW = 658; // Variable starting row

  // Collects Spreadsheet Data
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var playerNames = sheet.getRange(STARTROW, 3, lastRow - STARTROW + 1).getValues().flat(); // filters through player name column

  // Loops through each player's RealGM page
  for (var ticker = 0; ticker < playerNames.length; ticker++) {
    var playerUrl = sheet.getRange(ticker + STARTROW, 28).getValue().toString();

    // Fetches the HTML content of the player's RealGM page
    var response = UrlFetchApp.fetch(playerUrl);
    var htmlContent = response.getContentText();

    // Stores the First and Last name of the player
    var name = playerNames[ticker];

    // Stores the Position and Jersey Number of the player (A little buggy)
    var startIdx = htmlContent.indexOf("<h2");
    var endIdx = htmlContent.indexOf("</h2>", startIdx);
    var contentStart = startIdx + 3;
    var bio = htmlContent.substring(contentStart, endIdx);

    if (bio.length > 90){
      var jerseyNumber = bio.split(">")[4].replace("</span", "").replace("#", "");
      var position = bio.split(">")[2].replace("</span", "");
    }
    else {
      var jerseyNumber = "";
      var position = "";
    }

    // Stores the name of the player's team
    var team = extractData(htmlContent, /<strong>Current Team:<\/strong>\s*<a.*?>(.*?)<\/a>/);

    // some preferred exceptions
    if (team == "Pittsburgh"){
      team = "Pitt"
    }
    if (team == "Southern Methodist"){
      team = "SMU"
    }
    if (team == "Miami (FL)"){
      team = "Miami"
    }
    if (team == "Texas-San Antonio"){
      team = "UTSA"
    }
    if (team == "Southeastern Louisiana"){
      team = "SE Louisiana"
    }
    if (team == "Brigham Young"){
      team = "BYU"
    }
    if (team == "California"){
      team = "Cal"
    }
    if (team == "South Florida"){
      team = "USF"
    }
    if (team == "Loyola (IL)"){
      team = "Loyola Chicago"
    }
    if (team == "Saint Joseph's"){
      team = "St. Joseph's"
    }
    if (team == "Brisbane"){
      team = "Brisbane Bullets"
    }
    if (team == "Perth"){
      team = "Perth Wildcats"
    }
    if (team == "South East Melbourne"){
      team = "SE Melbourne Phoenix"
    }
    if (team == "Tasmania"){
      team = "Tasmania Jackjumpers"
    }
    if (team == "Sydney"){
      team = "Sydney Kings"
    }

    // Stores the birthday of the player
    var birthdate = extractData(htmlContent, /<strong>Born:<\/strong>\s*<a.*?>(.*?)<\/a>/);
    var birthdateObj = new Date(birthdate);
    var currentDate = new Date(); // today's date

    // converting age to ##.# years format 
    var ageInMilliseconds = currentDate - birthdateObj;
    var ageInYears = ageInMilliseconds / (365.25 * 24 * 60 * 60 * 1000);
    var age = ageInYears.toFixed(1);

    // manually insert birthday if not found on RealGM page
    if (isNaN(age)) {
      if (sheet.getRange(ticker + STARTROW, 29).getValue() != "") {
        birthdate = sheet.getRange(ticker + STARTROW, 29).getValue();
        birthdateObj = new Date(birthdate);
        currentDate = new Date();
        ageInMilliseconds = currentDate - birthdateObj;
        ageInYears = ageInMilliseconds / (365.25 * 24 * 60 * 60 * 1000);
        age = ageInYears.toFixed(1);
      }
      else{
        age = "";
      }
    }

    // Stores the height of the player
    var heightMatch = extractData(htmlContent, /<strong>Height:<\/strong>\s*([\d'-]+)/);
    var height = formatHeight(heightMatch);

    // Stores the weight of the player
    var weight = extractData(htmlContent, /<strong>Weight:<\/strong>\s*([\d]+)/);

    // Extracts the 2023-24 stats data
   var per36Stats = [];
   var statsPattern = extractData(htmlContent, /<tr[^>]*class\s*=\s*["']per_game["'][^>]*>\s*<td[^>]*>2023-24[^<]*<\/td>([\s\S]*?)<\/tr>/);
   var floatValue = null;
   var class2 = "";

  // going through the player stats table
    for (var j = 0; j < statsPattern.split('</td>\n<td>').length; j++) {
      var currentString = statsPattern.split('</td>\n<td>')[j];
      var parsedFloat = parseFloat(currentString);

      if (statsPattern.split('</td>\n<td>')[j].includes(".") && j > 0) {
        floatValue = parsedFloat;
        var class2 = statsPattern.split('</td>\n<td>')[j-3].slice(-2).toUpperCase();
        per36Stats[0] = statsPattern.split('</td>\n<td>')[j-2]; // gp
        per36Stats[1] = statsPattern.split('</td>\n<td>')[j]; // minutes
        per36Stats[2] = ((statsPattern.split('</td>\n<td>')[j+1]/per36Stats[1])*36).toFixed(1); // pts per 36
        per36Stats[3] = (((statsPattern.split('</td>\n<td>')[j+3]-statsPattern.split('</td>\n<td>')[j+6])/per36Stats[1])*36).toFixed(1); // 2PA per 36
        per36Stats[4] = (((statsPattern.split('</td>\n<td>')[j+2]-statsPattern.split('</td>\n<td>')[j+5])/(statsPattern.split('</td>\n<td>')[j+3]-statsPattern.split('</td>\n<td>')[j+6]))*100).toFixed(1); // 2P%
        per36Stats[5] = ((statsPattern.split('</td>\n<td>')[j+6]/per36Stats[1])*36).toFixed(1); // 3PA per 36
        per36Stats[6] = (statsPattern.split('</td>\n<td>')[j+7]*100).toFixed(1); // 3P%
        per36Stats[7] = ((statsPattern.split('</td>\n<td>')[j+9]/per36Stats[1])*36).toFixed(1); // FTA per 36
        per36Stats[8] = (statsPattern.split('</td>\n<td>')[j+10]*100).toFixed(1); // FT%
        per36Stats[9] = ((statsPattern.split('</td>\n<td>')[j+14]/per36Stats[1])*36).toFixed(1); // ast per 36
        per36Stats[10] = ((statsPattern.split('</td>\n<td>')[j+17]/per36Stats[1])*36).toFixed(1); // tov per 36
        per36Stats[11] = ((statsPattern.split('</td>\n<td>')[j+11]/per36Stats[1])*36).toFixed(1); // oreb per 36
        per36Stats[12] = ((statsPattern.split('</td>\n<td>')[j+12]/per36Stats[1])*36).toFixed(1); // dreb per 36
        per36Stats[13] = ((statsPattern.split('</td>\n<td>')[j+15]/per36Stats[1])*36).toFixed(1); // stl per 36
        per36Stats[14] = ((statsPattern.split('</td>\n<td>')[j+16]/per36Stats[1])*36).toFixed(1); // blk per 36
        per36Stats[15] = ((statsPattern.split('</td>\n<td>')[j+18].split('</td>')[0]/per36Stats[1])*36).toFixed(1); // fls per 36

        break;
      }
    }

    // Write data back into the sheet
    if (sheet.getRange(ticker + STARTROW, 4).getValue() == "") {
      sheet.getRange(ticker + STARTROW, 4).setValue(position);
    }
    if (sheet.getRange(ticker + STARTROW, 5).getValue() == "") {
      sheet.getRange(ticker + STARTROW, 5).setValue(team);
    }
    if (sheet.getRange(ticker + STARTROW, 6).getValue() == "") {
      sheet.getRange(ticker + STARTROW, 6).setValue(class2);
    }
    sheet.getRange(ticker + STARTROW, 7).setValue(age);
    sheet.getRange(ticker + STARTROW, 8).setValue(height);
    sheet.getRange(ticker + STARTROW, 10).setValue(weight);

    for (var i = 0; i < per36Stats.length; i++) {
      if (isNaN(per36Stats[i])){
        sheet.getRange(ticker + STARTROW, 12 + i).setValue(0);
      }
      else{
        sheet.getRange(ticker + STARTROW, 12 + i).setValue(per36Stats[i]);
      }
    }
      sheet.getRange(ticker + STARTROW, 30).setValue(jerseyNumber);

    Logger.log((STARTROW + ticker) + ": " + name);
  }
}

function extractData(content, regexPattern) {
  var regex = new RegExp(regexPattern);
  var match = regex.exec(content);
  return match ? match[1] : "";
}

// converting height data to #'##" format
function formatHeight(heightMatch) {
  if (heightMatch) {
    var parts = heightMatch.split('-');
    if (parts.length === 2) {
      var feet = parts[0];
      var inches = parts[1];
      return feet + "'" + inches + "\"";
    }
  }
  return "";
}
