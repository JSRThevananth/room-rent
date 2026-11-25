/**
 * Roommate Rent & Expense Splitter
 *
 * Sheet layout (created by setupSpreadsheet()):
 *   - "Roommates"   : one name per row in column A. Edit this sheet directly
 *                     to add/remove roommates — the web app always reads it live.
 *   - "Settings"    : key/value rows, e.g. "Monthly Rent (CAD)" -> 530.
 *   - "<yyyy-MM>"   : one sheet per calendar month, auto-created the first time
 *                     someone logs an expense that month. Columns: Timestamp,
 *                     Roommate, Item, Amount.
 *   - "RentHistory" : locked, append-only record of every rent calculation
 *                     that has ever been shown, so a later edit to an old
 *                     month's expenses can never silently change a rent figure
 *                     that was already shown/paid.
 */

var SHEET_ROOMMATES = 'Roommates';
var SHEET_SETTINGS = 'Settings';
var SHEET_RENT_HISTORY = 'RentHistory';
var SETTING_RENT_KEY = 'Monthly Rent (CAD)';
var SETTING_DISPLAY_DAY_KEY = 'Rent Display Day';
var SETTING_PIN_KEY = 'PIN';
var EXPENSE_HEADERS = ['Timestamp', 'Roommate', 'Item', 'Amount'];
var RENT_HISTORY_HEADERS = ['Timestamp', 'RentMonth', 'SpendMonth', 'Roommate', 'Spend', 'AverageSpend', 'Rent'];

/**
 * Every other function reaches the spreadsheet through this, never through
 * SpreadsheetApp.getActiveSpreadsheet() directly. "Active spreadsheet" only
 * exists for a script bound to a Sheet (opened via that Sheet's Extensions
 * menu); a standalone script (created from script.google.com or Drive's
 * "New > Apps Script") has no such thing and getActiveSpreadsheet() returns
 * null even when run from the editor. Storing the spreadsheet's own ID in
 * Script Properties works the same way for both kinds of project, and for
 * the deployed web app too.
 */
function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('No spreadsheet is set up yet. Run setupSpreadsheet() from the Apps Script editor once first.');
  }
  return SpreadsheetApp.openById(id);
}

/**
 * Run this once from the Apps Script editor (select it in the function
 * dropdown, click Run) to create the sheets this app needs. Safe to run
 * again later — it never overwrites a sheet that already exists, and it
 * always reuses the same spreadsheet once one has been set up.
 */
function setupSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    // Bound script (opened from inside a Sheet): reuse that Sheet.
    // Standalone script: create a fresh Sheet so there's something to use.
    ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('Roommate Rent Tracker');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }

  var roommatesSheet = ss.getSheetByName(SHEET_ROOMMATES);
  if (!roommatesSheet) {
    roommatesSheet = ss.insertSheet(SHEET_ROOMMATES);
    roommatesSheet.getRange(1, 1).setValue('Name').setFontWeight('bold');
    roommatesSheet.getRange(2, 1, 4, 1).setValues([['Roy'], ['Bavan'], ['Kana'], ['Prashan']]);
    roommatesSheet.setFrozenRows(1);
    roommatesSheet.autoResizeColumn(1);
  }

  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_SETTINGS);
    settingsSheet.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]).setFontWeight('bold');
    settingsSheet.getRange(2, 1, 3, 2).setValues([
      [SETTING_RENT_KEY, 530],
      [SETTING_DISPLAY_DAY_KEY, 20],
      [SETTING_PIN_KEY, '4655']
    ]);
    settingsSheet.setFrozenRows(1);
    settingsSheet.autoResizeColumns(1, 2);
  }

  var rentHistorySheet = ss.getSheetByName(SHEET_RENT_HISTORY);
  if (!rentHistorySheet) {
    rentHistorySheet = ss.insertSheet(SHEET_RENT_HISTORY);
    rentHistorySheet.getRange(1, 1, 1, RENT_HISTORY_HEADERS.length).setValues([RENT_HISTORY_HEADERS]).setFontWeight('bold');
    rentHistorySheet.setFrozenRows(1);
  }

  // Make sure there's a sheet ready for the current month's expenses.
  getOrCreateMonthSheet_(getTodayInfo_().currentMonthKey);

  SpreadsheetApp.flush();
}

/**
 * Serves the web app. Intentionally renders NO data — not even the
 * roommate list — until the PIN is verified, so nothing sensitive sits in
 * the page's HTML source before unlock. Everything is fetched by the
 * client via google.script.run once getHomeData(pin) succeeds.
 */
function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Roommate Rent Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lets Index.html pull in Stylesheet.html / JavaScript.html as includes. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Exposed to the client. Checks the PIN, then returns everything the home
 * page needs: the roommate list and, once unlocked, the current rent
 * breakdown.
 */
function getHomeData(pin) {
  var settings = requirePin_(pin);
  cleanupOldMonthSheets_();
  var today = getTodayInfo_();
  var showRent = today.day >= settings.displayDay;
  return {
    roommates: getRoommates_(),
    currentMonthLabel: formatMonthLabel_(today.currentMonthKey),
    showRent: showRent,
    rentInfo: showRent ? getOrComputeLockedRent_(today.currentMonthKey, today.prevMonthKey, settings) : null
  };
}

/**
 * Exposed to the client. Checks the PIN, then returns everything one
 * roommate needs for a "Details" view: their itemized entries for the
 * current (still-open) month, and — once the rent for this month has been
 * unlocked (see getOrComputeLockedRent_) — the itemized entries from the
 * spend month plus the full calculation breakdown for every roommate, so
 * nothing about the number is a black box.
 */
function getRoommateDetails(pin, roommateName) {
  var settings = requirePin_(pin);
  var roommates = getRoommates_();
  if (roommates.indexOf(roommateName) === -1) {
    throw new Error('"' + roommateName + '" is not in the Roommates sheet.');
  }

  var today = getTodayInfo_();

  var result = {
    currentMonthLabel: formatMonthLabel_(today.currentMonthKey),
    currentItems: getMonthItemsForRoommate_(today.currentMonthKey, roommateName),
    showRent: today.day >= settings.displayDay,
    rent: null
  };
  result.currentTotal = roundMoney_(result.currentItems.reduce(function (s, it) { return s + it.amount; }, 0));

  if (result.showRent) {
    var rentInfo = getOrComputeLockedRent_(today.currentMonthKey, today.prevMonthKey, settings);
    var mine = rentInfo.rows.filter(function (r) { return r.name === roommateName; })[0];
    result.rent = {
      spendMonthLabel: rentInfo.spendMonthLabel,
      items: getMonthItemsForRoommate_(today.prevMonthKey, roommateName),
      myTotal: mine ? mine.spend : 0,
      average: rentInfo.average,
      baseRent: settings.rent,
      myRent: mine ? mine.rent : 0,
      warning: rentInfo.warning,
      allRows: rentInfo.rows
    };
  }

  return result;
}

/**
 * Exposed to the client. Checks the PIN, then returns every month's locked
 * rent breakdown, newest first — the permanent record that survives even
 * after that month's itemized expense rows have been cleaned up.
 */
function getRentHistoryList(pin) {
  requirePin_(pin);
  return getRentHistory_();
}

/**
 * Exposed to the client. Checks the PIN, then validates and appends rows
 * to the current month's sheet. items: [{item: string, amount: number}, ...].
 * Returns the roommate's updated running total for the current month.
 */
function addExpenseEntries(pin, roommateName, items) {
  requirePin_(pin);
  var roommates = getRoommates_();
  if (roommates.indexOf(roommateName) === -1) {
    throw new Error('"' + roommateName + '" is not in the Roommates sheet. Add them there first.');
  }
  if (!items || items.length === 0) {
    throw new Error('Add at least one item with an amount before submitting.');
  }

  var cleanRows = [];
  var now = new Date();
  for (var i = 0; i < items.length; i++) {
    var itemName = (items[i].item || '').toString().trim();
    var amount = Number(items[i].amount);
    if (!itemName) {
      throw new Error('Item ' + (i + 1) + ' is missing a name.');
    }
    if (!isFinite(amount) || amount <= 0) {
      throw new Error('"' + itemName + '" needs a valid amount greater than 0.');
    }
    cleanRows.push([now, roommateName, itemName, roundMoney_(amount)]);
  }

  var monthKey = getTodayInfo_().currentMonthKey;
  var sheet = getOrCreateMonthSheet_(monthKey);
  sheet.getRange(sheet.getLastRow() + 1, 1, cleanRows.length, EXPENSE_HEADERS.length).setValues(cleanRows);

  var totals = getMonthTotals_(monthKey, roommates);
  return {
    monthLabel: formatMonthLabel_(monthKey),
    addedTotal: roundMoney_(cleanRows.reduce(function (sum, r) { return sum + r[3]; }, 0)),
    runningTotal: totals.totalsByRoommate[roommateName] || 0
  };
}

// ---------- internal helpers ----------

function getTodayInfo_() {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var day = Number(Utilities.formatDate(now, tz, 'd'));
  var currentMonthKey = Utilities.formatDate(now, tz, 'yyyy-MM');
  var prevMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var prevMonthKey = Utilities.formatDate(prevMonthDate, tz, 'yyyy-MM');
  return { day: day, currentMonthKey: currentMonthKey, prevMonthKey: prevMonthKey };
}

function formatMonthLabel_(monthKey) {
  // monthKey is "yyyy-MM"; render as "September 2026".
  var parts = monthKey.split('-');
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM yyyy');
}

function getSettings_() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    throw new Error('Settings sheet is missing. Run setupSpreadsheet() from the Apps Script editor first.');
  }
  var rows = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 2).getValues();
  var map = {};
  rows.forEach(function (r) { map[r[0]] = r[1]; });

  var rent = Number(map[SETTING_RENT_KEY]);
  var displayDay = Number(map[SETTING_DISPLAY_DAY_KEY]);
  var pin = String(map[SETTING_PIN_KEY] || '').trim();
  if (!isFinite(rent) || rent <= 0) {
    throw new Error('"' + SETTING_RENT_KEY + '" in the Settings sheet is missing or not a valid number.');
  }
  if (!isFinite(displayDay) || displayDay < 1 || displayDay > 28) {
    throw new Error('"' + SETTING_DISPLAY_DAY_KEY + '" in the Settings sheet is missing or not a valid day (1-28).');
  }
  if (!pin) {
    throw new Error('"' + SETTING_PIN_KEY + '" in the Settings sheet is missing.');
  }
  return { rent: rent, displayDay: displayDay, pin: pin };
}

/** Throws if `pin` doesn't match the Settings sheet's PIN. Returns settings for reuse. */
function requirePin_(pin) {
  var settings = getSettings_();
  if (String(pin || '').trim() !== settings.pin) {
    throw new Error('Incorrect PIN.');
  }
  return settings;
}

function getRoommates_() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_ROOMMATES);
  if (!sheet) {
    throw new Error('Roommates sheet is missing. Run setupSpreadsheet() from the Apps Script editor first.');
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return values
    .map(function (r) { return (r[0] || '').toString().trim(); })
    .filter(function (name) { return name.length > 0; });
}

/**
 * Every item one roommate logged in a given month, oldest first, with the
 * timestamp already formatted in the script's timezone (so it reads the
 * same for everyone regardless of whose browser/timezone is viewing it).
 */
function getMonthItemsForRoommate_(monthKey, roommateName) {
  var sheet = getSpreadsheet_().getSheetByName(monthKey);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var tz = Session.getScriptTimeZone();
  var rows = sheet.getRange(2, 1, lastRow - 1, EXPENSE_HEADERS.length).getValues();
  var items = [];
  rows.forEach(function (r) {
    var name = (r[1] || '').toString().trim();
    if (name === roommateName) {
      items.push({
        timestamp: Utilities.formatDate(new Date(r[0]), tz, 'MMM d, h:mm a'),
        item: r[2],
        amount: roundMoney_(Number(r[3]))
      });
    }
  });
  return items;
}

/**
 * Keeps only the current and previous month's itemized expense sheets.
 * Older ones are deleted outright — this is safe because the app only ever
 * needs the current month (to log into) and the previous month (as the
 * spend behind the current month's rent); once a month is no longer either
 * of those, the app will never read its sheet again. What must survive
 * permanently (each month's final rent per roommate) already lives in
 * RentHistory, which this never touches.
 */
function cleanupOldMonthSheets_() {
  var today = getTodayInfo_();
  var keep = {};
  keep[today.currentMonthKey] = true;
  keep[today.prevMonthKey] = true;

  var ss = getSpreadsheet_();
  var monthSheetPattern = /^\d{4}-\d{2}$/;
  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    if (monthSheetPattern.test(name) && !keep[name]) {
      ss.deleteSheet(sheet);
    }
  });
}

/** Every locked month's rent breakdown, newest first. */
function getRentHistory_() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_RENT_HISTORY);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, RENT_HISTORY_HEADERS.length).getValues();
  var byMonth = {};
  var order = [];
  data.forEach(function (r) {
    var rentMonthKey = r[1];
    if (!byMonth[rentMonthKey]) {
      byMonth[rentMonthKey] = { spendMonthKey: r[2], rows: [] };
      order.push(rentMonthKey);
    }
    byMonth[rentMonthKey].rows.push({ name: r[3], spend: Number(r[4]), rent: Number(r[6]) });
  });

  order.sort().reverse();
  return order.map(function (key) {
    var m = byMonth[key];
    return {
      rentMonthLabel: formatMonthLabel_(key),
      spendMonthLabel: formatMonthLabel_(m.spendMonthKey),
      rows: m.rows
    };
  });
}

function getOrCreateMonthSheet_(monthKey) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(monthKey);
  if (!sheet) {
    sheet = ss.insertSheet(monthKey);
    sheet.getRange(1, 1, 1, EXPENSE_HEADERS.length).setValues([EXPENSE_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Totals for a given month, keyed by every current roommate name (0 if they
 * logged nothing). Distinguishes "sheet doesn't exist yet" (nobody has used
 * the app for that month at all) from "sheet exists but someone has $0".
 */
function getMonthTotals_(monthKey, roommates) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(monthKey);
  var totals = {};
  roommates.forEach(function (name) { totals[name] = 0; });

  var sheetExists = !!sheet;
  var hasAnyEntries = false;
  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var rows = sheet.getRange(2, 1, lastRow - 1, EXPENSE_HEADERS.length).getValues();
      rows.forEach(function (r) {
        var name = (r[1] || '').toString().trim();
        var amount = Number(r[3]);
        if (name && isFinite(amount)) {
          hasAnyEntries = true;
          totals[name] = roundMoney_((totals[name] || 0) + amount);
        }
      });
    }
  }
  return { totalsByRoommate: totals, sheetExists: sheetExists, hasAnyEntries: hasAnyEntries };
}

/**
 * Returns the rent breakdown for rentMonthKey (the month rent is due for),
 * computed from spendMonthKey's logged expenses. The FIRST time this is
 * computed for a given rentMonthKey it is written to RentHistory and locked;
 * every later call that month reads back the locked numbers instead of
 * recomputing, so an edit to old expense rows can't silently change a rent
 * figure that was already shown.
 */
function getOrComputeLockedRent_(rentMonthKey, spendMonthKey, settings) {
  var locked = readLockedRent_(rentMonthKey);
  if (locked) return locked;

  var roommates = getRoommates_();
  if (roommates.length === 0) {
    return { rentMonthLabel: formatMonthLabel_(rentMonthKey), spendMonthLabel: formatMonthLabel_(spendMonthKey), rows: [], average: 0, warning: 'No roommates are configured. Add names to the Roommates sheet.' };
  }

  var totalsInfo = getMonthTotals_(spendMonthKey, roommates);
  var totalsByRoommate = totalsInfo.totalsByRoommate;
  var sum = roommates.reduce(function (s, name) { return s + totalsByRoommate[name]; }, 0);
  var average = sum / roommates.length;

  var rows = roommates.map(function (name) {
    var spend = totalsByRoommate[name];
    var rent = roundMoney_(settings.rent + (average - spend));
    return { name: name, spend: roundMoney_(spend), rent: rent };
  });

  var warning = null;
  if (!totalsInfo.sheetExists || !totalsInfo.hasAnyEntries) {
    warning = 'No expenses were logged for ' + formatMonthLabel_(spendMonthKey) + ', so rent defaults to an even split of the base rent.';
  }

  writeLockedRent_(rentMonthKey, spendMonthKey, rows, average);

  return { rentMonthLabel: formatMonthLabel_(rentMonthKey), spendMonthLabel: formatMonthLabel_(spendMonthKey), rows: rows, average: roundMoney_(average), warning: warning };
}

function readLockedRent_(rentMonthKey) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_RENT_HISTORY);
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var data = sheet.getRange(2, 1, lastRow - 1, RENT_HISTORY_HEADERS.length).getValues();
  var rows = [];
  var spendMonthKey = null;
  var average = 0;
  data.forEach(function (r) {
    if (r[1] === rentMonthKey) {
      spendMonthKey = r[2];
      average = Number(r[5]);
      rows.push({ name: r[3], spend: Number(r[4]), rent: Number(r[6]) });
    }
  });
  if (rows.length === 0) return null;
  return { rentMonthLabel: formatMonthLabel_(rentMonthKey), spendMonthLabel: formatMonthLabel_(spendMonthKey), rows: rows, average: average, warning: null };
}

function writeLockedRent_(rentMonthKey, spendMonthKey, rows, average) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_RENT_HISTORY);
  if (!sheet) return;
  var now = new Date();
  var values = rows.map(function (r) {
    return [now, rentMonthKey, spendMonthKey, r.name, r.spend, roundMoney_(average), r.rent];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, RENT_HISTORY_HEADERS.length).setValues(values);
}

function roundMoney_(n) {
  return Math.round(n * 100) / 100;
}
