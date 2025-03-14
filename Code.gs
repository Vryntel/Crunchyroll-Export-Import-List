function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Crunchyroll')
    .addSubMenu(ui.createMenu('Watch List')
      .addItem('Export Watch List', 'exportWatchlist')
      .addItem('Import Watch List', 'importWatchlist'))
    .addSubMenu(ui.createMenu('History')
      .addItem('Export History', 'exportHistory')
      .addItem('Import History', 'importHistory'))
    .addSubMenu(ui.createMenu('Crunchylists')
      .addItem('Export Crunchylists', 'exportCrunchyLists')
      .addItem('Import Crunchylists', 'importCrunchyLists'))
    .addItem('Refresh Anime List', 'getAnimeList')
    .addToUi();
}


function getToken() {
  var token = "";
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("Enter the Authorization token");
  token = result.getResponseText();
  return token;
}


function getAccountId(options) {
  const urlProfileInfo = "https://www.crunchyroll.com/accounts/v1/me";
  return JSON.parse(UrlFetchApp.fetch(urlProfileInfo, options).getContentText()).account_id;
}


function exportWatchlist() {

  // Get the Authentication token from user input
  const token = getToken();
  if (token == "") {
    return;
  }

  // Get the account_id that is necessary to export/import the Watchlist
  const options = {
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };
  const account_id = getAccountId(options);

  // Get Watchlist data (you can change n=500 parameter in the urlWatchlist to return n anime)
  const urlWatchlist = "https://www.crunchyroll.com/content/v2/discover/" + account_id + "/watchlist?order=desc&n=500";
  var watchlistJSON = JSON.parse(UrlFetchApp.fetch(urlWatchlist, options).getContentText()).data;

  var watchlist = []
  for (let i = 0; i < watchlistJSON.length; i++) {
    if (watchlistJSON[i].panel.type == "episode") {
      watchlist.push([watchlistJSON[i].panel.episode_metadata.series_id, watchlistJSON[i].panel.episode_metadata.series_title])
    }
    else if (watchlistJSON[i].panel.type == "movie") {
      watchlist.push([watchlistJSON[i].panel.movie_metadata.movie_listing_id, watchlistJSON[i].panel.movie_metadata.movie_listing_title])
    }
  }

  // Clear previous data
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Watch List");
  ss.getRange("A2:B").clear();

  // Write new data
  if (watchlist.length > 1) {
    ss.getRange("A2:B" + (watchlist.length + 1)).setValues(watchlist);
  }
}



function importWatchlist() {

  // Get the Authentication token from user input
  const token = getToken();
  if (token == "") {
    return;
  }

  // Get the account_id that is necessary to export/import the Watchlist
  const options = {
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };
  const account_id = getAccountId(options);

  // Get Anime codes from the sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Watch List");
  const watchlist = ss.getRange("A2:A" + ss.getLastRow()).getValues().flat();


  // Import the watchlist in the account
  const urlWatchList = "https://www.crunchyroll.com/content/v2/" + account_id + "/watchlist";

  var optionsWatchlist = {
    muteHttpExceptions: true,
    "method": "post",
    contentType: 'application/json',
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };


  for (let i = 0; i < watchlist.length; i++) {
    optionsWatchlist.payload = JSON.stringify({
      "content_id": watchlist[i]
    });

    // fetchAll() doesn't work, it only adds one anime from the list and ignore the others
    UrlFetchApp.fetch(urlWatchList, optionsWatchlist);
  }
}




function exportHistory() {

  // Get the Authentication token from user input
  const token = getToken();
  if (token == "") {
    return;
  }

  // Get the account_id that is necessary to export/import the Watchlist
  const options = {
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };
  const account_id = getAccountId(options);

  // Get History data (you can change n=700 parameter in the urlWatchlist to return n anime)
  const urlHistory = "https://www.crunchyroll.com/content/v2/" + account_id + "/watch-history?page_size=1000";
  const historyJSON = JSON.parse(UrlFetchApp.fetch(urlHistory, options).getContentText()).data;

  var history = []
  for (let i = 0; i < historyJSON.length; i++) {
    if (historyJSON[i].panel.type == "episode") {
      history.push([historyJSON[i].panel.id, historyJSON[i].panel.episode_metadata.series_title, historyJSON[i].panel.episode_metadata.season_number, historyJSON[i].panel.episode_metadata.episode]);
    }
    else if (historyJSON[i].panel.type == "movie") {
      history.push([historyJSON[i].panel.id, historyJSON[i].panel.movie_metadata.movie_listing_title, "", ""]);
    }
  }

  // Clear previous data in the sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("History");
  ss.getRange("A2:D").clear();

  if (history.length > 1) {
    ss.getRange("A2:D" + (history.length + 1)).setValues(history);
  }

}


function importHistory() {

  // Get the Authentication token from user input
  const token = getToken();
  if (token == "") {
    return;
  }

  // Get the account_id that is necessary to export/import the Watchlist
  const options = {
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };
  const account_id = getAccountId(options);

  // Get Anime codes from the sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("History");
  const history = ss.getRange("A2:A" + ss.getLastRow()).getValues().flat();

  // Import History
  const urlMarkAsWatched = "https://www.crunchyroll.com/content/v2/discover/" + account_id + "/mark_as_watched/";

  var optionsMarkAsWatched = {
    muteHttpExceptions: true,
    "method": "post",
    contentType: 'application/json',
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };

  for (let i = 0; i < history.length; i++) {
    // fetchAll() doesn't work, it only adds one anime from the list and ignore the others
    UrlFetchApp.fetch(urlMarkAsWatched + history[i], optionsMarkAsWatched);
  }
}



function exportCrunchyLists() {

  // Get the Authentication token from user input
  const token = getToken();
  if (token == "") {
    return;
  }

  // Get the account_id that is necessary to export/import the Watchlist
  const options = {
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };
  const account_id = getAccountId(options);

  // Get all Crunchylists
  const urlCrunchylists = "https://www.crunchyroll.com/content/v2/" + account_id + "/custom-lists";
  const crunchylistsJSON = JSON.parse(UrlFetchApp.fetch(urlCrunchylists, options).getContentText()).data;

  // Get anime inside each crunchylist
  var crunchylists = []
  var crunchylistData;
  for (let i = 0; i < crunchylistsJSON.length; i++) {
    crunchylistData = JSON.parse(UrlFetchApp.fetch(urlCrunchylists + "/" + crunchylistsJSON[i].list_id, options).getContentText()).data;
    for (let x = 0; x < crunchylistData.length; x++) {
      crunchylists.push([crunchylistsJSON[i].title, crunchylistData[x].id, crunchylistData[x].panel.title])
    }
  }

  // Clear previous data
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Crunchylist");
  ss.getRange("A2:C").clear();

  // Write new data
  if (crunchylists.length > 0) {
    ss.getRange("A2:C" + (crunchylists.length + 1)).setValues(crunchylists);
  }
}


function importCrunchyLists() {

  // Get the Authentication token from user input
  const token = getToken();
  if (token == "") {
    return;
  }

  // Get the account_id that is necessary to export/import the Watchlist
  const options = {
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };
  const account_id = getAccountId(options);




  // Get Crunchylist from the sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Crunchylist");
  var listData = ss.getRange("A2:B" + ss.getLastRow()).getValues();

  // Aggregate Crunchylist titles with animes
  /* Example:
  {
    CrunchyList1: [anime1, anime2, anime3]
    CrunchyList2: [anime5, anime6]
  }
  */

  listData = listData.reduce((obj, [key, value]) => {
    if (obj[key]) {
      obj[key].push(value);
    } else {
      obj[key] = [value];
    }
    return obj;
  }, {});


  // Create Crunchylists and add anime inside
  const urlCrunchylist = "https://www.crunchyroll.com/content/v2/" + account_id + "/custom-lists";

  var optionsCreateList = {
    muteHttpExceptions: true,
    "method": "post",
    contentType: 'application/json',
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };

  var optionsAddAnimeToList = {
    muteHttpExceptions: true,
    "method": "post",
    contentType: 'application/json',
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };

  for (var crunchylistTtitle in listData) {
    optionsCreateList.payload = JSON.stringify({
      "title": crunchylistTtitle
    });
    let list_id = JSON.parse(UrlFetchApp.fetch(urlCrunchylist, optionsCreateList).getContentText()).data[0].list_id;
    for (var anime of listData[crunchylistTtitle]) {
      optionsAddAnimeToList.payload = JSON.stringify({
        "content_id": anime
      });
      UrlFetchApp.fetch(urlCrunchylist + "/" + list_id, optionsAddAnimeToList).getContentText();
    }
  }
}



function getAnimeList() {

  // Get the Authentication token
  const token = getToken();
  if (token == "") {
    return;
  }

  // You can change the n=1500 part (currently there are 1261 anime so it's enough)
  const url = "https://www.crunchyroll.com/content/v2/discover/browse?start=0&n=1500&sort_by=alphabetical";

  const options = {
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/*",
      "Accept-Encoding": "gzip, deflate, br, zstd",
      "Authorization": token
    }
  };

  const languages = {
    "ja-JP": "Japanese",
    "en-US": "English",
    "en-IN": "English (India)",
    "id-ID": "Bahasa Indonesia",
    "ms-MY": "Bahasa Melayu",
    "ca-ES": "Català",
    "de-DE": "Deutsch",
    "es-419": "Español (América Latina)",
    "es-ES": "Español (España)",
    "fr-FR": "Français",
    "it-IT": "Italiano",
    "pl-PL": "Polski",
    "pt-BR": "Português (Brasil)",
    "pt-PT": "Português (Portugal)",
    "vi-VN": "Tiếng Việt",
    "tr-TR": "Türkçe",
    "ru-RU": "Русский",
    "ar-SA": "العربية",
    "hi-IN": "हिंदी",
    "ta-IN": "தமிழ்",
    "te-IN": "తెలుగు",
    "zh-CN": "中文 (普通话)",
    "zh-HK": "中文 (粵語)",
    "zh-TW": "中文 (國語)",
    "ko-KR": "한국어",
    "th-TH": "ไทย"
  }

  const response = JSON.parse(UrlFetchApp.fetch(url, options).getContentText()).data;
  var rows = [];

  for (let i = 0; i < response.length; i++) {
    var title = response[i].title;
    var type = response[i].type;
    var link;
    var animeCode = response[i]["id"];
    if (type == "series") {
      link = "https://www.crunchyroll.com/series/" + animeCode;
    }
    else if (type == "movie_listing") {
      link = "https://www.crunchyroll.com/watch/" + animeCode;
    }

    // Check if language is present in object and map it with its value
    var audio = [];
    try {
      audio = response[i].series_metadata.audio_locales;
      audio = audio.filter(key => languages.hasOwnProperty(key))
        .map(key => languages[key]);
      audio.sort();
    }
    catch (e) {
      audio = [];
    }

    var sub = [];
    try {
      if (type == "series") {
        sub = response[i].series_metadata.subtitle_locales;
      }
      else if (type == "movie_listing") {
        sub = response[i].movie_listing_metadata.subtitle_locales;
      }
      sub = sub.filter(key => languages.hasOwnProperty(key))
        .map(key => languages[key]);
      sub.sort();
    }
    catch (e) {
      sub = [];
    }

    rows.push([title, link, animeCode, audio.join(","), sub.join(",")]);
  }

  // Clear previous data
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Anime");
  ss.getRange("A2:E").clear();

  ss.getRange("A2:E" + (rows.length + 1)).setValues(rows);
}
