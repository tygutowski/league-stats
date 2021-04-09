function UpdateGraph() {
  ResetGraph(); // Clears the entire chart, except the column titles
  var matches = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Matches"); // Gets the "Matches" sheet
  var config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration"); // Gets the "Configuration" sheet

  // Collects the data from "Configuration"
  var userName = config.getRange("B1").getValue();
  var apiKey = config.getRange("B2").getValue();
  var numberOfMatches = config.getRange("B3").getValue();

  // Default numberOfMatches = 10
  if (numberOfMatches == "") {
    numberOfMatches = 10;
  }
  var accountDetails = "https://na1.api.riotgames.com/lol/summoner/v4/summoners/by-name/" + userName + "?api_key=" + apiKey;
  try {
    var response = UrlFetchApp.fetch(accountDetails, {'muteHttpExceptions': true});
    var json = response.getContentText();
    var accountData = JSON.parse(json);

  } 
  catch (e) {
    Logger.log(e); // This will catch fetching error, like if the MH link is  wrong
  }
  var accountId = accountData["accountId"];
  var matchList = "https://na1.api.riotgames.com/lol/match/v4/matchlists/by-account/" + accountId + "?api_key=" + apiKey; 

  try {
    var response = UrlFetchApp.fetch(matchList, {'muteHttpExceptions': true});
    var json = response.getContentText();
    var matchListData = JSON.parse(json);
  } 
  catch (e) {
    Logger.log(e); // This will catch fetching error
  }

  var matchIndex = 0;
  for(var match = matchListData["startIndex"]; match < numberOfMatches; match++) {
    var gameId = matchListData["matches"][match]["gameId"]
    var champ = matchListData["matches"][match]["champion"];
    var champName = getChampionName(getChampions(), champ);
    var gameURL = "https://na1.api.riotgames.com/lol/match/v4/matches/" + gameId + "?api_key=" + apiKey;
    try {
      var response = UrlFetchApp.fetch(gameURL, {'muteHttpExceptions': true});
      var json = response.getContentText();
      var gameData = JSON.parse(json);
    }
    catch (e) {
      Logger.log(e); // This will catch fetching error
    }
    if(gameData["queueId"] == 420){
      matchIndex += 1;
    var winningTeam = "";
    if(gameData["teams"]["0"]["win"] == "Win") {
      winningTeam = "blue";
    }
    else if(gameData["teams"]["1"]["win"] == "Win") {
      winningTeam = "red";
    }
    var userParticipant = 0;
    for (var participant = 0; participant <= 9; participant++) {
      if(gameData["participantIdentities"][participant]["player"]["summonerName"] == userName) {
        userParticipant = participant;
      }
    }
    var userTeam = "";
    if (gameData["participants"][userParticipant]["teamId"] == 100){
      userTeam = "blue";
    }
    else {
      userTeam = "red";
    }

    var outcomePrint;
    if (userTeam == winningTeam) {
      outcomePrint = "Victory";
    }
    else {
      outcomePrint = "Defeat";
    }
    var blueKills = 0;
    for (var i = 0; i < 5; i++) {
      blueKills += gameData["participants"][i]["stats"]["kills"];
    }
    var redKills = 0;
    for (var i = 5; i < 10; i++) {
      redKills += gameData["participants"][i]["stats"]["kills"];
    }
    var blueGold = 0;
    for (var i = 0; i < 5; i++) {
      blueGold += gameData["participants"][i]["stats"]["goldEarned"];
    }
    var redGold = 0;
    for (var i = 5; i < 10; i++) {
      redGold += gameData["participants"][i]["stats"]["goldEarned"];
    }
    var blueVision = 0;
    for (var i = 0; i < 5; i++) {
      blueVision += gameData["participants"][i]["stats"]["visionScore"];
    }
    var redVision = 0;
    for (var i = 5; i < 10; i++) {
      redVision += gameData["participants"][i]["stats"]["visionScore"];
    }
      var teamKills = 0;
      var teamGold = 0;
      var teamVision = 0;
      var enemyKills = 0;
      var enemyGold = 0;
      var enemyVision = 0;
    if(userTeam == "red"){
      teamKills = redKills;
      teamGold =  redGold;
      teamVision = redVision;
      enemyKills = blueKills;
      enemyGold = blueGold;
      enemyVision = blueVision;
      ally1 = getChampionName(getChampions(), gameData["participants"][5]["championId"]);
      ally2 = getChampionName(getChampions(), gameData["participants"][6]["championId"]);
      ally3 = getChampionName(getChampions(), gameData["participants"][7]["championId"]);
      ally4 = getChampionName(getChampions(), gameData["participants"][8]["championId"]);
      ally5 = getChampionName(getChampions(), gameData["participants"][9]["championId"]);
      enemy1 = getChampionName(getChampions(), gameData["participants"][0]["championId"]);
      enemy2 = getChampionName(getChampions(), gameData["participants"][1]["championId"]);
      enemy3 = getChampionName(getChampions(), gameData["participants"][2]["championId"]);
      enemy4 = getChampionName(getChampions(), gameData["participants"][3]["championId"]);
      enemy5 = getChampionName(getChampions(), gameData["participants"][4]["championId"]);
    }
    else if(userTeam == "blue"){
      teamKills = blueKills;
      teamGold =  blueGold;
      teamVision = blueVision;
      enemyKills = redKills;
      enemyGold = redGold;
      enemyVision = redVision;
      ally1 = getChampionName(getChampions(), gameData["participants"][0]["championId"]);
      ally2 = getChampionName(getChampions(), gameData["participants"][1]["championId"]);
      ally3 = getChampionName(getChampions(), gameData["participants"][2]["championId"]);
      ally4 = getChampionName(getChampions(), gameData["participants"][3]["championId"]);
      ally5 = getChampionName(getChampions(), gameData["participants"][4]["championId"]);
      enemy1 = getChampionName(getChampions(), gameData["participants"][5]["championId"]);
      enemy2 = getChampionName(getChampions(), gameData["participants"][6]["championId"]);
      enemy3 = getChampionName(getChampions(), gameData["participants"][7]["championId"]);
      enemy4 = getChampionName(getChampions(), gameData["participants"][8]["championId"]);
      enemy5 = getChampionName(getChampions(), gameData["participants"][9]["championId"]);
    }
    var userKills = gameData["participants"][userParticipant]["stats"]["kills"];
    var userDeaths = gameData["participants"][userParticipant]["stats"]["deaths"];
    var userAssists = gameData["participants"][userParticipant]["stats"]["assists"];
    //var userCreepScore = gameData["participants"][userParticipant]["stats"]
    var userGold = gameData["participants"][userParticipant]["stats"]["goldEarned"];
    var userVision = gameData["participants"][userParticipant]["stats"]["visionScore"];
    var userWardsPlaced = gameData["participants"][userParticipant]["stats"]["wardsPlaced"];
    var userWardsKilled =gameData["participants"][userParticipant]["stats"]["wardsKilled"];
    var userControlWards = gameData["participants"][userParticipant]["stats"]["visionWardsBoughtInGame"];
    var userLevel = gameData["participants"][userParticipant]["stats"]["champLevel"];
    var matchDuration = gameData["gameDuration"];
    //var userSumD = gameData["participants"][userParticipant]["spell1Id"].getSumSpell();
    //var userSumF = gameData["participants"][userParticipant]["spell2Id"].getSumSpell();
    var userDamageDealt = gameData["participants"][userParticipant]["stats"]["totalDamageDealtToChampions"];
    var userDamageTaken = gameData["participants"][userParticipant]["stats"]["totalDamageTaken"];
    var userDamageHealed = gameData["participants"][userParticipant]["stats"]["totalHeal"];
    var userTowerDamage = gameData["participants"][userParticipant]["stats"]["damageDealtToTurrets"];
    var userFirstBlood = gameData["participants"][userParticipant]["stats"]["firstBloodKill"];
    var userDoubleKills = gameData["participants"][userParticipant]["stats"]["doubleKills"];
    var userTripleKills = gameData["participants"][userParticipant]["stats"]["tripleKills"];
    var userQuadraKills = gameData["participants"][userParticipant]["stats"]["quadraKills"];
    var userPentaKills = gameData["participants"][userParticipant]["stats"]["pentaKills"];
    var userItem1 = getItemName(getItems(), gameData["participants"][userParticipant]["stats"]["item0"]);
    var userItem2 = getItemName(getItems(), gameData["participants"][userParticipant]["stats"]["item1"]);
    var userItem3 = getItemName(getItems(), gameData["participants"][userParticipant]["stats"]["item2"]);
    var userItem4 = getItemName(getItems(), gameData["participants"][userParticipant]["stats"]["item3"]);
    var userItem5 = getItemName(getItems(), gameData["participants"][userParticipant]["stats"]["item4"]);
    var userItem6 = getItemName(getItems(), gameData["participants"][userParticipant]["stats"]["item5"]);

    // Set each specific cell to its appropriate value
    matches.getRange(matchIndex+2,1).setValue(gameId);
    matches.getRange(matchIndex+2,2).setValue(outcomePrint);
    matches.getRange(matchIndex+2,3).setValue(champName);
    matches.getRange(matchIndex+2,4).setValue(userKills);
    matches.getRange(matchIndex+2,5).setValue(userDeaths);
    matches.getRange(matchIndex+2,6).setValue(userAssists);
    matches.getRange(matchIndex+2,7).setValue((userKills+userAssists)/userDeaths);
    matches.getRange(matchIndex+2,8).setValue(userTeam);
    //matches.getRange(matchIndex+2,9).setValue(userCreepScore); (Yet to be implemented)
    matches.getRange(matchIndex+2,10).setValue(userGold);
    matches.getRange(matchIndex+2,11).setValue(teamGold);
    matches.getRange(matchIndex+2,12).setValue(enemyGold);
    matches.getRange(matchIndex+2,13).setValue(userGold/teamGold*100.0);
    matches.getRange(matchIndex+2,14).setValue(userWardsPlaced);
    matches.getRange(matchIndex+2,15).setValue(userWardsKilled);
    matches.getRange(matchIndex+2,16).setValue(userControlWards);
    matches.getRange(matchIndex+2,17).setValue(userVision);
    matches.getRange(matchIndex+2,18).setValue(teamVision);
    matches.getRange(matchIndex+2,19).setValue(enemyVision);
    matches.getRange(matchIndex+2,20).setValue(userVision/teamVision*100.0);
    matches.getRange(matchIndex+2,21).setValue(teamKills);
    matches.getRange(matchIndex+2,22).setValue(enemyKills);
    matches.getRange(matchIndex+2,23).setValue((userKills+userAssists)/teamKills*100.0);
    matches.getRange(matchIndex+2,24).setValue(userLevel);
    matches.getRange(matchIndex+2,25).setValue(matchDuration);
    //matches.getRange(matchIndex+2,26).setValue(userSumD);
    //matches.getRange(matchIndex+2,27).setValue(userSumF);
    matches.getRange(matchIndex+2,28).setValue(userItem1);
    matches.getRange(matchIndex+2,29).setValue(userItem2);
    matches.getRange(matchIndex+2,30).setValue(userItem3);
    matches.getRange(matchIndex+2,31).setValue(userItem4);
    matches.getRange(matchIndex+2,32).setValue(userItem5);
    matches.getRange(matchIndex+2,33).setValue(userItem6);
    matches.getRange(matchIndex+2,34).setValue(userDamageDealt);
    matches.getRange(matchIndex+2,35).setValue(userDamageTaken);
    matches.getRange(matchIndex+2,36).setValue(userDamageHealed);
    matches.getRange(matchIndex+2,37).setValue(userTowerDamage);
    matches.getRange(matchIndex+2,38).setValue(userFirstBlood);
    matches.getRange(matchIndex+2,39).setValue(userDoubleKills);
    matches.getRange(matchIndex+2,40).setValue(userTripleKills);
    matches.getRange(matchIndex+2,41).setValue(userQuadraKills);
    matches.getRange(matchIndex+2,42).setValue(userPentaKills);
    matches.getRange(matchIndex+2,43).setValue(ally1);
    matches.getRange(matchIndex+2,44).setValue(ally2);
    matches.getRange(matchIndex+2,45).setValue(ally3);
    matches.getRange(matchIndex+2,46).setValue(ally4);
    matches.getRange(matchIndex+2,47).setValue(ally5);
    matches.getRange(matchIndex+2,48).setValue(enemy1);
    matches.getRange(matchIndex+2,49).setValue(enemy2);
    matches.getRange(matchIndex+2,50).setValue(enemy3);
    matches.getRange(matchIndex+2,51).setValue(enemy4);
    matches.getRange(matchIndex+2,52).setValue(enemy5);
    }
  }
}
// Clear the entire chart,except the column names
function ResetGraph() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Matches").getRange("A2:ZZ1000").clearContent();
}
// Get a list of the champions
function getChampions() {
  var url = 'http://ddragon.leagueoflegends.com/cdn/11.7.1/data/en_US/champion.json';
  try {
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    var json = response.getContentText();
    return JSON.parse(json);
  } 
  catch (e) { // This will catch error if the champion doesn't exist
    Logger.log(e);
  }
}
// Get the specific name of the champion, given the ID
function getChampionName(champions, id) {
  var championName = "?";
  var championKeys = Object.keys(champions['data']);
  for (var i = 0; i < championKeys.length; i++) {
    if (champions['data'][championKeys[i]]['key'] == id)
      championName = champions['data'][championKeys[i]]['name'];
  }
  return championName;
}

// Get a list of the items
function getItems() {
  var url = 'http://ddragon.leagueoflegends.com/cdn/11.7.1/data/en_US/item.json';
  try {
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    var json = response.getContentText();
    return JSON.parse(json);
  } 
  catch (e) { // This will catch error if the item doesn't exist
    Logger.log(e);
  }
}
// Get the specific name of the item, given the ID
function getItemName(items, id) {
  var itemName = "?";
  var itemKeys = Object.keys(items['data']);
  for (var i = 0; i < itemKeys.length; i++) {
    if(items['data'][itemKeys[i]]['image']['full'] == (id + ".png")){
    
      itemName = items['data'][itemKeys[i]]['name'];
    }
  }
  return itemName;
}
