# leagueStats
A Google Scripts script which collects data from the Riot API and uploads the data to the sheet.

Under the "Configuration" sheet:
  userName = Summoner's name
  apiKey = The API key that developer.riotgames.com offers (Must be updated every 24 hours)
  gameList = The list of games that will be collected, between 0 and 100, inclusive.
  queueType = The queue that will be collected (Note: Only Solo/Duo is collected currently, changing it does nothing)

Sometimes the API will limit data collection rates, so it may only offer 30 or 40 rows, even if 100 were requested.
