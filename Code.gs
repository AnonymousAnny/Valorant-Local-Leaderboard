function fetchData(url) {
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText();
  return JSON.parse(json);
}

function getPlayerRank(url) {
  var data = fetchData(url);
  return data.data.current_data.currenttierpatched;
}

function getPlayerRR(url) {
  var data = fetchData(url);
  return data.data.current_data.ranking_in_tier;
}

function getPlayerLogo(url) {
  var data = fetchData(url);
  return data.data.current_data.images.large;
}

function getElo(url) {
  var data = fetchData(url);
  return data.data.current_data.elo;
}

function getGamesNeededForRating(url) {
  var data = fetchData(url);
  return data.data.current_data.games_needed_for_rating;
}

function getPlayerCard(url) {
  var data = fetchData(url);
  return data.data.card.small;
}

function getAccountLevel(url) {
  var data = fetchData(url);
  return data.data.account_level;
}

function getHighestRank(url) {
  var data = fetchData(url);
  return data.data.highest_rank.patched_tier;
}

function getHighestRankSeason(url) {
  var data = fetchData(url);
  return data.data.highest_rank.season;
}
