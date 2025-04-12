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

// Used in cell A2 inside "Columns to Use" sheet 
function showAllAvailableColumns() {
  var columns;
  columns = Object.entries(WATCHLIST_COLUMNS).map(subarray => [...subarray, "Watchlist"]);
  columns = columns.concat(Object.entries(HISTORY_COLUMNS).map(subarray => [...subarray, "History"]));
  columns = columns.concat(Object.entries(CRUNCHYLIST_COLUMNS).map(subarray => [...subarray, "Crunchylists"]));
  return columns;
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

  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Watch List");
  const sheetColumns = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues().flat();

  // Check column names
  for (const key of sheetColumns) {
    if (!WATCHLIST_COLUMNS.hasOwnProperty(key)) {
      SpreadsheetApp.getUi().alert(key + " is not a valid column name");
      return;
    }
  }

  //Get the Authentication token from user input
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


  // For each Anime (First loop) I need to extract each column specified in the sheet (second loop) which value can be in a sub-object of the returned JSON (Third loop, example: Anime Code -> panel.episode_metadata.series_id)
  var watchlist = []
  for (let i = 0; i < watchlistJSON.length; i++) {
    var row = [];
    for (let x = 0; x < sheetColumns.length; x++) {
      var json_path = WATCHLIST_COLUMNS[sheetColumns[x]].split(".");
      var currentobj = watchlistJSON[i];

      for (let y = 0; y < json_path.length; y++) {
        if (currentobj.hasOwnProperty(json_path[y])) {
          // Key found inside JSON
          currentobj = currentobj[json_path[y]];
        }
        else if (json_path[y] in movie_keys && currentobj.hasOwnProperty(movie_keys[json_path[y]])) {
          // It's a movie
          currentobj = currentobj[movie_keys[json_path[y]]];
        }
        else {
          // Something else like an episode_media 
          currentobj = "";
          break;
        }
      }
      row.push(currentobj)
    }
    watchlist.push(row);
  }


  // Clear previous data from the sheet
  ss.getRange(2, 1, ss.getLastRow(), ss.getLastColumn()).clear();

  // Write new data
  if (watchlist.length > 0) {
    ss.getRange(2, 1, watchlist.length, sheetColumns.length).setValues(watchlist);
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

  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("History");
  const sheetColumns = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues().flat();

  // Check column names
  for (const key of sheetColumns) {
    if (!HISTORY_COLUMNS.hasOwnProperty(key)) {
      SpreadsheetApp.getUi().alert(key + " is not a valid column name");
      return;
    }
  }

  // Get the Authentication token from user input
  const token = getToken();
  if (token == "") {
    return;
  }

  // Get the account_id that is necessary to export/import the Watchlist
  const options = {
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      "Accept": "*/* ",
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
    var row = [];

    for (let x = 0; x < sheetColumns.length; x++) {
      var json_path = HISTORY_COLUMNS[sheetColumns[x]].split(".");
      var currentobj = historyJSON[i];
      console.log()
      for (let y = 0; y < json_path.length; y++) {
        if (currentobj.hasOwnProperty(json_path[y])) {
          currentobj = currentobj[json_path[y]];
        }
        else if (json_path[y] in movie_keys) {
          if (currentobj.hasOwnProperty(movie_keys[json_path[y]])) {
            // It's a movie
            currentobj = currentobj[movie_keys[json_path[y]]];
          }
          else {
            // Something else like episode_media
            currentobj = "";
            break;
          }
        }
        else {
          currentobj = "";
          break;
        }
      }

      row.push(currentobj)
    }

    history.push(row);
  }


  // Clear previous data from the sheet
  ss.getRange(2, 1, ss.getLastRow(), ss.getLastColumn()).clear();

  // Write new data
  if (history.length > 0) {
    ss.getRange(2, 1, history.length, sheetColumns.length).setValues(history);
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

  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Crunchylist");
  const sheetColumns = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues().flat();


  // Check column names
  for (const key of sheetColumns) {
    if (!CRUNCHYLIST_COLUMNS.hasOwnProperty(key)) {
      SpreadsheetApp.getUi().alert(key + " is not a valid column name");
      return;
    }
  }


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
    for (let z = 0; z < crunchylistData.length; z++) {
      //crunchylists.push([crunchylistsJSON[i].title, crunchylistData[x].id, crunchylistData[x].panel.title])
      var row = [];
      for (let x = 0; x < sheetColumns.length; x++) {
        var json_path = CRUNCHYLIST_COLUMNS[sheetColumns[x]].split(".");
        var currentobj = crunchylistData[z];

        // If it's the crunchylist title I extract it from the first response of the crunchylists
        if (json_path.length > 0 && json_path[0] == "title") {
          currentobj = crunchylistsJSON[i].title;
        }
        else {
          for (let y = 0; y < json_path.length; y++) {
            if (currentobj.hasOwnProperty(json_path[y])) {
              currentobj = currentobj[json_path[y]];
            }
            // if it's a movie i replace the key with the equivalent of series
            else if (json_path[y] in movie_keys && currentobj.hasOwnProperty(movie_keys[json_path[y]])) {
              // It's a movie
              currentobj = currentobj[movie_keys[json_path[y]]];
            }
            else {
              // Something else like an episode_media 
              currentobj = "";
              break;
            }
          }
        }

        row.push(currentobj)
      }
      crunchylists.push(row);
    }
  }

  // Clear previous data from the sheet
  ss.getRange(2, 1, ss.getLastRow(), ss.getLastColumn()).clear();

  // Write new data
  if (crunchylists.length > 0) {
    ss.getRange(2, 1, crunchylists.length, sheetColumns.length).setValues(crunchylists);
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
