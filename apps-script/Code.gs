/****************************
 * CONFIG – EDIT FOR YOUR SETUP
 ****************************/

var RAW_DATA_SHEET = "RAW_DATA";  // Main data sheet name

// Valid transaction types
// EXP = Expense, CC = Card spend, BILL = Card bill payment, CR = Credit/Income, SAV = Savings/Investments
var ALLOWED_TYPES = ["EXP", "CC", "BILL", "CR", "SAV"];

// Valid card names – keep in sync with your actual cards in use
var VALID_CARDS   = ["HDFC", "ICICI", "AMAZON", "AXIS", "YES", "SBI"];

// Email address to receive monthly summary
var SUMMARY_EMAIL_RECIPIENT = "shubhamjog7@gmail.com";  // TODO: set to your email

function generateAiInsightsPDF(htmlContent) {
  try {
    // Wrap content in minimal HTML for PDF
    var fullHtml =
      "<html><head>" +
      "<style>" +
      "body { font-family: Arial, sans-serif; padding: 20px; }" +
      ".ai-section { margin-bottom: 14px; }" +
      ".ai-section-title { font-weight: bold; font-size: 16px; margin-bottom: 6px; }" +
      "</style>" +
      "</head><body>" +
      htmlContent +
      "</body></html>";

    var blob = Utilities.newBlob(fullHtml, "text/html", "ai_insights.html");

    var pdfBlob = blob.getAs("application/pdf");

    return Utilities.base64Encode(pdfBlob.getBytes());

  } catch (err) {
    throw new Error("PDF generation failed: " + err.message);
  }
}
  
/****************************
 * BUDGET CONFIG (HYBRID: OVERALL + CATEGORY)
 ****************************/

function getBudgetConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("BUDGET");
  if (!sheet) {
    // No budget sheet yet
    return {
      overallBudget: null,
      categoryBudgets: {} // { CATEGORY_NAME_UPPER: budgetNumber }
    };
  }

  var values = sheet.getDataRange().getValues();
  var overallBudget = null;
  var categoryBudgets = {};

  for (var i = 0; i < values.length; i++) {
    var key = (values[i][0] || "").toString().trim();
    var val = values[i][1];

    if (!key) continue;

    if (key.toUpperCase() === "OVERALLBUDGET") {
      overallBudget = Number(val) || 0;
    } else {
      var catName = key.toUpperCase();
      var budgetVal = Number(val);
      if (!isNaN(budgetVal) && budgetVal > 0) {
        categoryBudgets[catName] = budgetVal;
      }
    }
  }

  return {
    overallBudget: overallBudget,
    categoryBudgets: categoryBudgets
  };
}

function getAiTrendComparisonCompact(monthKey) {
  try {
    if (!monthKey) {
      const d = new Date();
      monthKey = d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2);
    }

    // Compute previous month
    const [y, m] = monthKey.split("-").map(n => parseInt(n, 10));
    const prev = new Date(y, m - 2, 1);
    const prevMonthKey =
      prev.getFullYear() + "-" + ("0" + (prev.getMonth() + 1)).slice(-2);

    // Load data
    const curr = getMonthlyData(monthKey, "ALL").summary || {};
    const prevData = getMonthlyData(prevMonthKey, "ALL").summary || {};

    const payload = {
      currentMonth: monthKey,
      previousMonth: prevMonthKey,
      current: curr,
      previous: prevData
    };

    // ---------------------------------------------
    // Compact AI prompt
    // ---------------------------------------------
    const prompt =
      "You are an Indian personal finance analyst. Compare TWO months and produce a compact, dashboard‑style HTML summary.\n\n" +
      "RULES:\n" +
      "• Always use ₹ (never $).\n" +
      "• Use ↑ (green) and ↓ (red) arrows.\n" +
      "• Keep insights short.\n" +
      "• Format output EXACTLY like this:\n\n" +

      "<div class='ai-report'>\n" +
      "  <div class='ai-section ai-summary'>\n" +
      "    <div class='ai-section-title'>📆 Month vs Month</div>\n" +
      "    <div><b>Previous:</b> PREV_MONTH</div>\n" +
      "    <div><b>Current:</b> CURR_MONTH</div>\n" +
      "  </div>\n\n" +

      "  <div class='ai-section ai-spend'>\n" +
      "    <div class='ai-section-title'>📊 Spending Trends</div>\n" +
      "    <ul></ul>\n" +
      "  </div>\n\n" +

      "  <div class='ai-section ai-savings'>\n" +
      "    <div class='ai-section-title'>💰 Savings & Cashflow</div>\n" +
      "    <ul></ul>\n" +
      "  </div>\n\n" +

      "  <div class='ai-section ai-reco'>\n" +
      "    <div class='ai-section-title'>🎯 Quick Recommendations</div>\n" +
      "    <ul></ul>\n" +
      "  </div>\n" +
      "</div>\n\n" +

      "Fill each <ul> with 2–3 bullets. Use ₹ amounts and arrows. Keep it compact.\n\n" +
      "Data:\n" +
      JSON.stringify(payload, null, 2);

    // Gemini API call
    var apiKey  = "aAkienfjn8649NBIBIk4uybh";     // same as your insights function
    var modelId = "gemini-3.1-flash-lite-preview"; // same model ID

    var url =
      "https://generativelanguage.googleapis.com/v1beta/models/" +
      modelId +
      ":generateContent?key=" + apiKey;

    var body = {
      contents: [{ parts: [{ text: prompt }] }]
    };

    var res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify(body)
    });

    var json = JSON.parse(res.getContentText());

    var aiText =
      json?.candidates?.[0]?.content?.parts?.[0]?.text || null;

    if (!aiText) throw new Error("No content returned. Raw: " + res.getContentText());

    return aiText.replace("PREV_MONTH", prevMonthKey)
                 .replace("CURR_MONTH", monthKey);

  } catch (err) {
    return "<b>Error:</b> " + err.message;
  }
}
/****************************
 * UTIL: ERROR LOGGING
 ****************************/

function logError(source, err, extra) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("LOGS") || ss.insertSheet("LOGS");
    sheet.appendRow([
      new Date(),
      source,
      err && err.message ? err.message : err,
      JSON.stringify(extra || {})
    ]);
  } catch (e) {
    // Avoid recursive failure if logging itself breaks
  }
}

/****************************
 * CATEGORY CONFIG (Category -> SubCategory)
 * Sheet: CATEGORY_CONFIG
 * Columns: A=Category, B=SubCategory, C=Active(TRUE/FALSE)
 ****************************/

function getCategoryConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CATEGORY_CONFIG");
  if (!sheet) {
    // If config sheet missing, return empty map
    return { categories: [], map: {} };
  }

  var values = sheet.getDataRange().getValues();
  var map = {};       // { CAT: [SUB1, SUB2...] }
  var catsSet = {};

  for (var i = 1; i < values.length; i++) {
    var cat = (values[i][0] || "").toString().trim();
    var sub = (values[i][1] || "").toString().trim();
    var active = (values[i][2] === "" || values[i][2] === null) ? true : String(values[i][2]).toUpperCase() === "TRUE";

    if (!cat || !sub || !active) continue;

    var CAT = cat.toUpperCase();
    catsSet[CAT] = true;

    if (!map[CAT]) map[CAT] = [];
    // avoid duplicates
    if (map[CAT].indexOf(sub) === -1) map[CAT].push(sub);
  }

  var categories = Object.keys(catsSet).sort();
  // sort subcats for each category
  categories.forEach(function(C) {
    map[C].sort();
  });

  return { categories: categories, map: map };
}

/****************************
 * SUB-CATEGORY MAINTENANCE (RENAME / MERGE)
 * Assumes RAW_DATA schema:
 * 0 Timestamp | 1 Type | 2 Amount | 3 Category | 4 SubCategory | 5 Notes | 6 Card | 7 Tag
 ****************************/

function renameSubCategory(category, fromSub, toSub, deactivateOld) {
  category = (category || "").toString().trim().toUpperCase();
  fromSub = (fromSub || "").toString().trim();
  toSub   = (toSub || "").toString().trim();

  if (!category || !fromSub || !toSub) throw new Error("Category, fromSub, toSub required.");
  if (fromSub === toSub) return "No change needed (same sub-category).";

  // 1) Update RAW_DATA
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RAW_DATA_SHEET);
  if (!sheet) throw new Error("RAW_DATA sheet not found.");

  var values = sheet.getDataRange().getValues();
  var updated = 0;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var cat = String(row[3] || "").toUpperCase();
    var sub = String(row[4] || "");
    if (cat === category && sub === fromSub) {
      sheet.getRange(i + 1, 5).setValue(toSub); // col 5 = SubCategory
      updated++;
    }
  }

  // 2) Ensure new mapping exists in CATEGORY_CONFIG
  addCategorySubcategory(category, toSub);

  // 3) Deactivate old mapping (optional)
  if (deactivateOld) {
    deactivateCategorySubRow_(category, fromSub);
  }

  return "Renamed '" + fromSub + "' → '" + toSub + "' in " + updated + " rows.";
}

function mergeSubCategories(category, fromSubs, toSub, deactivateOld) {
  category = (category || "").toString().trim().toUpperCase();
  toSub = (toSub || "").toString().trim();
  if (!category || !toSub) throw new Error("Category and target sub-category required.");
  if (!fromSubs || !fromSubs.length) throw new Error("Select at least one sub-category to merge.");

  // 1) Update RAW_DATA: map all fromSubs -> toSub
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RAW_DATA_SHEET);
  if (!sheet) throw new Error("RAW_DATA sheet not found.");

  var values = sheet.getDataRange().getValues();
  var updated = 0;

  var fromSet = {};
  fromSubs.forEach(function(s){ fromSet[String(s)] = true; });

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var cat = String(row[3] || "").toUpperCase();
    var sub = String(row[4] || "");
    if (cat === category && fromSet[sub]) {
      sheet.getRange(i + 1, 5).setValue(toSub); // col 5 = SubCategory
      updated++;
    }
  }

  // 2) Ensure target mapping exists
  addCategorySubcategory(category, toSub);

  // 3) Deactivate old mappings (optional)
  if (deactivateOld) {
    fromSubs.forEach(function(s){
      deactivateCategorySubRow_(category, String(s));
    });
  }

  return "Merged " + fromSubs.length + " sub-categories → '" + toSub + "' in " + updated + " rows.";
}

// Helper: set Active=FALSE in CATEGORY_CONFIG for a given category/sub
function deactivateCategorySubRow_(categoryUpper, subExact) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CATEGORY_CONFIG");
  if (!sheet) return;

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var c = (values[i][0] || "").toString().trim().toUpperCase();
    var s = (values[i][1] || "").toString().trim();
    if (c === categoryUpper && s === subExact) {
      sheet.getRange(i + 1, 3).setValue("FALSE"); // Active column
    }
  }
}

/**
 * Adds a new Category/SubCategory row to CATEGORY_CONFIG
 * If category doesn't exist, it is created.
 */
function addCategorySubcategory(cat, sub) {
  cat = (cat || "").toString().trim();
  sub = (sub || "").toString().trim();
  if (!cat || !sub) throw new Error("Category and SubCategory are required.");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CATEGORY_CONFIG") || ss.insertSheet("CATEGORY_CONFIG");

  // Ensure headers exist
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Category", "SubCategory", "Active"]);
  } else if (sheet.getLastRow() === 1 && sheet.getLastColumn() < 3) {
    sheet.getRange(1,1,1,3).setValues([["Category","SubCategory","Active"]]);
  }

  // Prevent duplicates (case-insensitive category + exact sub)
  var values = sheet.getDataRange().getValues();
  var CAT = cat.toUpperCase();

  for (var i = 1; i < values.length; i++) {
    var c = (values[i][0] || "").toString().trim().toUpperCase();
    var s = (values[i][1] || "").toString().trim();
    var a = (values[i][2] === "" || values[i][2] === null) ? "TRUE" : String(values[i][2]).toUpperCase();
    if (c === CAT && s === sub && a === "TRUE") {
      return "OK (already exists)";
    }
  }

  sheet.appendRow([cat, sub, "TRUE"]);
  return "OK";
}

/**
 * Returns a combined list of suggestion tokens for Notes:
 * - all subcategories
 * - all categories
 */
function getNotesSuggestionTokens() {
  var cfg = getCategoryConfig();
  var out = [];
  cfg.categories.forEach(function(C) {
    out.push(C);
    (cfg.map[C] || []).forEach(function(sub) {
      out.push(sub);
    });
  });
  // unique
  var seen = {};
  var uniq = [];
  out.forEach(function(x) {
    var k = x.toLowerCase();
    if (!seen[k]) { seen[k] = true; uniq.push(x); }
  });
  return uniq;
}

/**
 * Optional helper: check if RAW_DATA has a SubCategory column.
 * Expected columns (new): Timestamp | Type | Amount | Category | SubCategory | Notes | Card | Tag
 * Old columns: Timestamp | Type | Amount | Category | Notes | Card | Tag
 */
function rawDataHasSubCategoryColumn_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RAW_DATA_SHEET);
  if (!sheet) return false;

  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(x){ return String(x||"").trim().toLowerCase(); });
  return header.indexOf("subcategory") !== -1;
}

/****************************
 * FORM SUBMIT HANDLER
 * (Google Form → RAW_DATA)
 ****************************/

function onFormSubmit(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RAW_DATA_SHEET);
    if (!sheet) {
      throw new Error("Sheet '" + RAW_DATA_SHEET + "' not found");
    }

    Logger.log(JSON.stringify(e.namedValues));

    var message   = e.namedValues["Message"]     ? e.namedValues["Message"][0]     : "";
    var entryDate = e.namedValues["Entry Date"] ? e.namedValues["Entry Date"][0] : "";

    if (!message) {
      Logger.log("No 'Message' field in form submission. Skipping.");
      return;
    }

    var parts = message.trim().split(" ");
    if (parts.length < 3) {
      throw new Error("Invalid message format. Use: TYPE AMOUNT CATEGORY [Notes]");
    }

    var type   = parts[0].toUpperCase();
    var amount = Number(parts[1]);
    if (isNaN(amount) || amount <= 0) {
      throw new Error("Amount must be a positive number.");
    }

    if (ALLOWED_TYPES.indexOf(type) === -1) {
      throw new Error("Invalid type. Use EXP / CC / BILL / CR / SAV");
    }

    var category = parts[2];
    var notes    = parts.slice(3).join(" ");
    var card     = "";

    // CC / BILL → requires valid card
    if (type === "CC" || type === "BILL") {
      if (VALID_CARDS.indexOf(category.toUpperCase()) === -1) {
        throw new Error("Invalid card name. Valid: " + VALID_CARDS.join(", "));
      }
      card     = category.toUpperCase();
      category = "CreditCard";
    }

    // CR (credits) – category like SALARY, REFUND, FRIEND, etc.
    if (type === "CR") {
      category = category.toUpperCase();
    }

    // SAV (savings) – category like MF, GOLD, NPS, STOCKS, RD, EMERGENCY_FUND, etc.
    if (type === "SAV") {
      category = category.toUpperCase();
    }

    var timestamp;
    if (entryDate) {
      timestamp = new Date(entryDate);
      var now = new Date();
      if (timestamp > now) {
        throw new Error("Entry Date cannot be in the future.");
      }
    } else {
      timestamp = new Date();
    }

    // Tag: CR is Credit, everything else (EXP/CC/BILL/SAV) is money OUT
    var tag = (type === "CR") ? "Credit" : "Debit";

    sheet.appendRow([timestamp, type, amount, category, notes, card, tag]);
  } catch (err) {
    logError("onFormSubmit", err, e ? e.namedValues : null);
    throw err;
  }
}


/****************************
 * WEB APP ENDPOINTS
 ****************************/

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Expense Manager – Monthly View")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function addTransaction(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RAW_DATA_SHEET);
    if (!sheet) {
      throw new Error("Sheet '" + RAW_DATA_SHEET + "' not found");
    }

    var type      = (data.type || "").toUpperCase();
    var amount    = Number(data.amount);
    var category  = data.category || "";
    var notes     = data.notes || "";
    var card      = data.card || "";
    var entryDate = data.entryDate || "";
    var subCategory = data.subCategory || "";
    

    if (!type || isNaN(amount) || amount <= 0) {
      throw new Error("Invalid type or amount.");
    }

    if (ALLOWED_TYPES.indexOf(type) === -1) {
      throw new Error("Invalid type. Use EXP / CC / BILL / CR / SAV.");
    }

    // Optional: prevent absurd amounts
    if (amount > 10000000) {
      throw new Error("Amount seems too large. Please double check.");
    }

    // CC / BILL logic
    if (type === "CC" || type === "BILL") {
      if (!card || VALID_CARDS.indexOf(card.toUpperCase()) === -1) {
        throw new Error("Invalid card. Valid: " + VALID_CARDS.join(", "));
      }
      category = "CreditCard";
      card     = card.toUpperCase();
    } else {
      card = "";
    }

    // CR → normalize category
    if (type === "CR") {
      category = category.toUpperCase();
    }

    // SAV → normalize category (MF, GOLD, NPS, etc.)
    if (type === "SAV") {
      category = category.toUpperCase();
    }

    var timestamp;
    if (entryDate) {
      timestamp = new Date(entryDate);
      var now = new Date();
      if (timestamp > now) {
        throw new Error("Entry Date cannot be in the future.");
      }
    } else {
      timestamp = new Date();
    }

    // Tag: only CR is Credit; everything else (EXP/CC/BILL/SAV) is Debit
    var tag = (type === "CR") ? "Credit" : "Debit";

    category = (category || "").toString().trim();
    subCategory = (subCategory || "").toString().trim();

var hasSubCol = rawDataHasSubCategoryColumn_();
var finalCategory = category;

if (!hasSubCol && subCategory) {
  finalCategory = category.toUpperCase() + " > " + subCategory;
}

var tag = (type === "CR") ? "Credit" : "Debit";

if (hasSubCol) {
  sheet.appendRow([timestamp, type, amount, category.toUpperCase(), subCategory, notes, card, tag]);
} else {
  sheet.appendRow([timestamp, type, amount, finalCategory, notes, card, tag]);
}
return "OK";

    return "OK";
  } catch (err) {
    logError("addTransaction", err, data);
    throw err;
  }
  
}


/****************************
 * MONTHLY SUMMARY + DAILY (CARD FILTER AWARE, WITH SAVINGS)
 ****************************/

function getMonthlyData(monthKey, cardFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RAW_DATA_SHEET);
    if (!sheet) {
      throw new Error("Sheet '" + RAW_DATA_SHEET + "' not found");
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        summary: {
          totalNormalExpenses: 0,
          totalCCSpend: 0,
          totalCCBillsPaid: 0,
          actualSpending: 0,
          ccOutstanding: 0,
          totalCredits: 0,
          totalSavings: 0,
          netCashFlow: 0,
          overallBudget: null,
          overallBudgetRemaining: null,
          overallBudgetUsedPct: null
        },
        daily: [],
        cardSummary: null,
        categoryBudgets: [],
        cardTotals: []
      };
    }

    var now = new Date();
    if (!monthKey) {
      var m = ("0" + (now.getMonth() + 1)).slice(-2);
      monthKey = now.getFullYear() + "-" + m;  // "YYYY-MM"
    }

    cardFilter = (cardFilter || "").toUpperCase();
    var hasCardFilter = cardFilter && cardFilter !== "ALL";

    var monthExp     = 0;
    var monthCCSpend = 0;
    var monthBill    = 0;
    var monthCr      = 0;
    var monthSav     = 0;

    var globalCCSpend = 0;
    var globalCCBill  = 0;

    var cardMonthSpend  = 0;
    var cardMonthBill   = 0;
    var cardGlobalSpend = 0;
    var cardGlobalBill  = 0;

    var dailyMap      = {};   // date -> {debit, credit}
    var categorySpend = {};   // EXP categories this month: CAT -> amount
    var subCategorySpend = {}; // key: CAT||SUB -> amount
    var cardMonthMap  = {};   // for charts: CARD -> {ccSpend, ccBill} in this month

    var tz = Session.getScriptTimeZone();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var type = String(row[1] || "").toUpperCase();
      var ts  = row[0];
      if (!(ts instanceof Date)) continue;

      var type  = String(row[1] || "").toUpperCase();
      var amt  = Number(row[2] || 0);
      var cat   = String(row[3] || "").toUpperCase();  // Category column
      var subCat = String(row[4] || "").toUpperCase();  // SubCategory column (NEW schema)
      var card  = String(row[5] || "").toUpperCase();  // Card column
      var tag   = String(row[6] || "");                // Tag: Credit / Debit

      if (!amt || amt <= 0) continue;

      var ym = Utilities.formatDate(ts, tz, "yyyy-MM");

      // ----- GLOBAL CC OUTSTANDING -----
      if (type === "CC") {
        globalCCSpend += amt;
      } else if (type === "BILL") {
        globalCCBill += amt;
      }

      // ----- GLOBAL PER-CARD for cardFilter cardSummary -----
      if (hasCardFilter && card === cardFilter) {
        if (type === "CC") {
          cardGlobalSpend += amt;
        } else if (type === "BILL") {
          cardGlobalBill += amt;
        }
      }

      // ----- ONLY STATS FOR SELECTED MONTH -----
      if (ym !== monthKey) continue;

      // Per-type monthly aggregates
      // ---- 1. Type-based totals ----
if (type === "EXP") {
  monthExp += amt;
  if (cat) {
    categorySpend[cat] = (categorySpend[cat] || 0) + amt;
  }

} else if (type === "CC") {
  monthCCSpend += amt;

} else if (type === "BILL") {
  monthBill += amt;

} else if (type === "CR") {
  monthCr += amt;

} else if (type === "SAV") {
  monthSav += amt;
}

// ---- 2. Sub-category aggregation (separate concern) ----
if (type === "EXP" || type === "SAV") {
  if (cat && subCat) {
    var key = cat + "||" + subCat;
    subCategorySpend[key] = (subCategorySpend[key] || 0) + amt;
  }
}


      // Card-specific monthly for cardFilter (detail view in summary)
      if (hasCardFilter && card === cardFilter) {
        if (type === "CC") {
          cardMonthSpend += amt;
        } else if (type === "BILL") {
          cardMonthBill += amt;
        }
      }

      // CardTotals for chart (ALL CARDS)
      if (type === "CC" || type === "BILL") {
        if (!cardMonthMap[card]) {
          cardMonthMap[card] = { ccSpend: 0, ccBill: 0 };
        }
        if (type === "CC") {
          cardMonthMap[card].ccSpend += amt;
        } else if (type === "BILL") {
          cardMonthMap[card].ccBill += amt;
        }
      }

      // Daily map (for trend / daily summary)
      var dateKey = Utilities.formatDate(ts, tz, "yyyy-MM-dd");
      

      if (!hasCardFilter) {
        // Overall daily view (debit includes EXP/CC/BILL/SAV)
        // ✅ DAILY SUMMARY LOGIC (FIXED)
      if (!dailyMap[dateKey]) {
        dailyMap[dateKey] = { date: dateKey, credit: 0, debit: 0 };
          }

        if (type === "CR") {
        // Income
          dailyMap[dateKey].credit += amt;
        } else {
        // EXP, CC, BILL, SAV
        dailyMap[dateKey].debit += amt;
}

      } else {
        // Daily view for cardFilter – only CC/BILL for that card
        if ((type === "CC" || type === "BILL") && card === cardFilter) {
          if (!dailyMap[dateKey]) {
            dailyMap[dateKey] = { debit: 0, credit: 0 };
          }
          dailyMap[dateKey].debit += amt;
        }
      }
    }

    // ----- SUMMARY LEVEL AGGREGATES -----
    var actualSpending = monthExp + monthBill;   // only EXP + BILL
    var ccOutstanding  = globalCCSpend - globalCCBill;
    var netCashFlow    = monthCr - actualSpending - monthSav; // CR - (EXP+BILL) - SAV

    var cardSummary = null;
    if (hasCardFilter) {
      var cardOutstanding = cardGlobalSpend - cardGlobalBill;
      cardSummary = {
        card: cardFilter,
        monthSpend: cardMonthSpend,
        monthBill: cardMonthBill,
        outstanding: cardOutstanding
      };
    }

    // ----- HYBRID BUDGETS -----
    var budgetConfig = getBudgetConfig();
    var overallBudget = budgetConfig.overallBudget;
    var overallBudgetRemaining = null;
    var overallBudgetUsedPct   = null;

    if (overallBudget && overallBudget > 0) {
      overallBudgetRemaining = overallBudget - actualSpending;
      overallBudgetUsedPct   = (actualSpending / overallBudget) * 100;
    }

    var catBudgets = budgetConfig.categoryBudgets;
    var categoryBudgetArr = [];
    for (var catName in catBudgets) {
      if (!catBudgets.hasOwnProperty(catName)) continue;
      var budget = catBudgets[catName];
      var spent  = categorySpend[catName] || 0;
      var diff   = budget - spent;
      var usedPct = budget > 0 ? (spent / budget * 100) : null;

      categoryBudgetArr.push({
        category: catName,
        budget: budget,
        spent: spent,
        diff: diff,
        usedPct: usedPct
      });
    }

    // ----- DAILY ARRAY -----
    var dates = Object.keys(dailyMap).sort();
    var dailyArr = dates.map(function(d) {
      var obj   = dailyMap[d];
      var debit = obj.debit || 0;
      var credit= obj.credit || 0;
      return {
        date: d,
        debit: debit,
        credit: credit,
        net: credit - debit
      };
    });

    // ----- CARD TOTALS FOR BAR CHART -----
    var cardTotalsArr = [];
    for (var c in cardMonthMap) {
      if (!cardMonthMap.hasOwnProperty(c)) continue;
      var v = cardMonthMap[c];
      // Only show cards that had any activity this month
      if (v.ccSpend > 0 || v.ccBill > 0) {
        cardTotalsArr.push({
          card: c,
          ccSpend: v.ccSpend,
          ccBill: v.ccBill
        });
      }
    }

    
var subCategoryTotalsArr = [];
Object.keys(subCategorySpend).forEach(function(k) {
  var parts = k.split("||");
  subCategoryTotalsArr.push({
    category: parts[0],
    subCategory: parts[1],
    amount: subCategorySpend[k]
  });
});


    return {
      summary: {
        totalNormalExpenses: monthExp,
        totalCCSpend:       monthCCSpend,
        totalCCBillsPaid:   monthBill,
        actualSpending:     actualSpending,
        ccOutstanding:      ccOutstanding,
        totalCredits:       monthCr,
        totalSavings:       monthSav,
        netCashFlow:        netCashFlow,
        overallBudget:          overallBudget,
        overallBudgetRemaining: overallBudgetRemaining,
        overallBudgetUsedPct:   overallBudgetUsedPct
      },
      daily: dailyArr,
      cardSummary: cardSummary,
      categoryBudgets: categoryBudgetArr,
      subCategoryTotals: subCategoryTotalsArr,
      cardTotals: cardTotalsArr
    };
  } catch (err) {
    logError("getMonthlyData", err, { monthKey: monthKey, cardFilter: cardFilter });
    throw err;
  }
}

/****************************
 * AI MONTHLY INSIGHTS – BACKEND
 ****************************/

 /*
function forceExternalRequestPermission() {
  var response = UrlFetchApp.fetch("https://api.openai.com/v1/");
  Logger.log(response.getResponseCode());
}*/


/****************************
 * AI MONTHLY INSIGHTS – BACKEND (GEMINI)
 ****************************/


/****************************
 * AI MONTHLY INSIGHTS – BACKEND (GEMINI 3.1 Flash Lite Preview)
 ****************************/
function getAiFinanceInsights(monthKey) {
  try {
    // 1) Default to current month if not provided
    if (!monthKey) {
      var now = new Date();
      monthKey = now.getFullYear() + "-" + ("0" + (now.getMonth() + 1)).slice(-2);
    }

    // 2) Get your existing monthly data
    var data = getMonthlyData(monthKey, "ALL");  // summary + daily + categoryBudgets + cardTotals
    var summary    = data.summary || {};
    var catBudgets = data.categoryBudgets || [];
    var daily      = data.daily || [];
    var cardTotals = data.cardTotals || [];

    // 3) Build compact JSON payload for AI
    var payloadForAi = {
      month: monthKey,
      summary: summary,
      topCategories: catBudgets
        .slice()
        .sort(function (a, b) { return b.spent - a.spent; })
        .slice(0, 5),
      dailyStats: daily,
      cardStats: cardTotals,
      subCategoryStats: data.subCategoryTotals || []
    };

    // 4) Build the prompt Gemini should respond to
   var userPrompt =
  "You are an Indian personal finance analyst. Interpret the user’s JSON and return a structured HTML report.\n\n" +
  "IMPORTANT – Sub‑Category Analysis:\n"+
  "• If subCategoryStats is provided, analyze spending at sub‑category level.\n"+
  "• Identify top contributing sub‑categories inside each category.\n"+
  "• Call out overspends clearly (₹ and % if possible).\n"+
  "• Mention which sub‑category is driving category-level increase.\n"+
  "• Use ₹ only (Indian Rupees).\n"+
  "• Do not invent numbers – only use provided JSON.\n"+
  "STRICT RULES:\n" +
  "• ALWAYS use Indian Rupees (₹), NEVER dollars.\n" +
  "• If you mention amounts, prefix with <span class='ai-rupee'>₹</span>.\n" +
  "• Keep bullet points short and practical.\n" +
  "• Use clear insights — no generic advice.\n" +
  "• Use trend arrows:\n" +
  "    ↑ = positive trend (use <span class='ai-trend-up'>↑</span>)\n" +
  "    ↓ = negative trend (use <span class='ai-trend-down'>↓</span>)\n\n" +
  "OUTPUT HTML EXACTLY IN THIS STRUCTURE:\n" +
  "<div class='ai-report'>\n" +
  "  <div class='ai-section ai-summary'>\n" +
  "    <div class='ai-section-title'>📌 Monthly Summary</div>\n" +
  "    <ul></ul>\n" +
  "  </div>\n" +
  "  <div class='ai-section ai-spend'>\n" +
  "    <div class='ai-section-title'>📊 Spending Breakdown</div>\n" +
  "    <ul></ul>\n" +
  "  </div>\n" +
  "  <div class='ai-section ai-risk'>\n" +
  "    <div class='ai-section-title'>⚠️ Risks & Overspending Alerts</div>\n" +
  "    <ul></ul>\n" +
  "  </div>\n" +
  "  <div class='ai-section ai-savings'>\n" +
  "    <div class='ai-section-title'>💰 Savings & Credits Overview</div>\n" +
  "    <ul></ul>\n" +
  "  </div>\n" +
  "  <div class='ai-section ai-daily'>\n" +
  "    <div class='ai-section-title'>📅 Daily Trend Highlights</div>\n" +
  "    <ul></ul>\n" +
  "  </div>\n" +
  "  <div class='ai-section ai-reco'>\n" +
  "    <div class='ai-section-title'>🎯 Recommendations for Next Month</div>\n" +
  "    <ul></ul>\n" +
  "  </div>\n" +
  "</div>\n\n" +
  "Insert meaningful insights into each <ul> based on this JSON:\n" +
  "If subCategoryStats is empty, fall back to category-level insights only.\n"+
  "Use phrases like “mainly due to”, “driven by”, or “largest contributor”.\n"+
  "When discussing categories, always mention the top 1–2 sub‑categories responsible.\n"+
  JSON.stringify(payloadForAi, null, 2);

    // 5) Gemini API key from AI Studio
    var apiKey = "AIzaSyBTeSO0wyX7D2FyC53TUDABODRDsr8sXE0";  // <-- REPLACE with your real Gemini key

    if (!apiKey || apiKey === "YOUR_GEMINI_AI_STUDIO_KEY_HERE") {
      return "<b>AI not configured:</b> Please set your Gemini API key in Code.gs.";
    }

    // 6) Exact model ID for "Gemini 3.1 Flash Lite Preview"
    //    Confirm this ID in AI Studio; adjust if name is slightly different.
    var modelId = "gemini-3.1-flash-lite-preview";

    // 7) Correct AI Studio REST endpoint: v1beta + models/<MODEL_ID>:generateContent?key=
    var apiUrl =
      "https://generativelanguage.googleapis.com/v1beta/models/" +
      modelId +
      ":generateContent?key=" +
      apiKey;

    // 8) Request body using userPrompt
    var requestBody = {
      contents: [
        {
          parts: [
            { text: userPrompt }
          ]
        }
      ]
    };

    var options = {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify(requestBody)
    };

    // 9) Call Gemini
    var response = UrlFetchApp.fetch(apiUrl, options);
    var resText  = response.getContentText();
    var json     = JSON.parse(resText);

    // 10) Extract AI text from Gemini response
    var aiText = "";

    if (json &&
        json.candidates &&
        json.candidates.length > 0 &&
        json.candidates[0].content &&
        json.candidates[0].content.parts &&
        json.candidates[0].content.parts.length > 0 &&
        json.candidates[0].content.parts[0].text) {

      aiText = json.candidates[0].content.parts[0].text;
    } else {
      throw new Error("Gemini returned no content. Raw: " + resText);
    }

    // 11) Return formatted HTML to be shown in the AI card
    return "<b>AI Insights for " + monthKey + ":</b><br><br>" +
           aiText.replace(/\n/g, "<br>");

  } catch (err) {
    logError("getAiFinanceInsights", err, { monthKey: monthKey });
    return "<b>AI Insights Error:</b> " + err.message;
  }
}


/****************************
 * SEARCH & FILTER – BACKEND
 ****************************/

function searchTransactions(filters) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RAW_DATA_SHEET);
    if (!sheet) {
      throw new Error("Sheet '" + RAW_DATA_SHEET + "' not found");
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return [];
    }

    // Normalize filters
    var text       = (filters.text || "").toString().toLowerCase().trim();
    var typeFilter = (filters.type || "ALL").toUpperCase();
    var cardFilter = (filters.card || "ALL").toUpperCase();
    var categoryFilter    = (filters.category || "ALL").toUpperCase();
    var subCategoryFilter = (filters.subCategory || "ALL").toUpperCase();
    var dateFrom   = filters.dateFrom ? new Date(filters.dateFrom) : null;
    var dateTo     = filters.dateTo   ? new Date(filters.dateTo)   : null;
    var minAmount  = filters.minAmount ? Number(filters.minAmount) : null;
    var maxAmount  = filters.maxAmount ? Number(filters.maxAmount) : null;

    var tz = Session.getScriptTimeZone();
    var results = [];

    // RAW_DATA columns:
    // 0: Timestamp | 1: Type | 2: Amount | 3: Category | 4: Notes | 5: Card | 6: Tag
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      var ts      = row[0];
      var type    = String(row[1] || "").toUpperCase();
      var amount  = Number(row[2]);
      var category    = String(row[3] || "").toUpperCase();
      var subCategory = String(row[4] || "").toUpperCase();
      var notes   = String(row[5] || "");
      var card    = String(row[6] || "").toUpperCase();
      var tag     = String(row[7] || "");

      if (!(ts instanceof Date)) {
        continue; // skip invalid rows
      }

      // ----- Apply filters -----

      // Type filter
      if (typeFilter !== "ALL" && type !== typeFilter) {
        continue;
      }

      // Card filter
      if (cardFilter !== "ALL" && card !== cardFilter) {
        continue;
      //new block
      if (categoryFilter !== "ALL" && category !== categoryFilter) {
  continue;
}

if (subCategoryFilter !== "ALL" && subCategory !== subCategoryFilter) {
  continue;
}
      }

      // Date range
      if (dateFrom && ts < dateFrom) {
        continue;
      }
      if (dateTo) {
        // make dateTo inclusive for the whole day
        var endOfDay = new Date(dateTo);
        endOfDay.setHours(23, 59, 59, 999);
        if (ts > endOfDay) {
          continue;
        }
      }

      // Amount range
      if (minAmount != null && !isNaN(minAmount) && amount < minAmount) {
        continue;
      }
      if (maxAmount != null && !isNaN(maxAmount) && amount > maxAmount) {
        continue;
      }

      // Text search (in type, category, notes, card, tag)
      if (text) {
        var combined = (type + " " + category + " " + notes + " " + card + " " + tag).toLowerCase();
        if (combined.indexOf(text) === -1) {
          continue;
        }
      }

      // If passed all filters, add to results
      results.push({
        date: Utilities.formatDate(ts, tz, "yyyy-MM-dd HH:mm"),
        type: type,
        amount: amount,
        category: category,
        notes: notes,
        card: card,
        tag: tag
      });
    }

    return results;
  } catch (err) {
    logError("searchTransactions", err, filters);
    throw err;
  }
}


/****************************
 * MONTHLY EMAIL SUMMARY
 ****************************/

function sendMonthlyEmailSummary() {
  try {
    if (!SUMMARY_EMAIL_RECIPIENT) {
      throw new Error("Please set SUMMARY_EMAIL_RECIPIENT to a valid email address.");
    }

    var now = new Date();
    var lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    var m = ("0" + (lastMonth.getMonth() + 1)).slice(-2);
    var ym = lastMonth.getFullYear() + "-" + m;  // e.g. "2026-02"

    var data = getMonthlyData(ym, "ALL");
    var s = data.summary;

    var subject = "Monthly Expense Summary – " + ym;
    var body =
      "Monthly Expense Summary for " + ym + "\n\n" +
      "Total Normal Expenses (EXP):       " + s.totalNormalExpenses + "\n" +
      "Total CC Spend (swipes):           " + s.totalCCSpend       + "\n" +
      "Total CC Bills Paid:               " + s.totalCCBillsPaid   + "\n" +
      "Actual Spending (EXP+BILL):        " + s.actualSpending     + "\n" +
      "Total Savings (SAV):               " + s.totalSavings       + "\n" +
      "Total Credits (CR):                " + s.totalCredits       + "\n" +
      "Net Cash Flow (after savings):     " + s.netCashFlow        + "\n" +
      "CC Outstanding (overall):          " + s.ccOutstanding      + "\n" +
      "Overall Budget (if set):           " + (s.overallBudget || "N/A") + "\n" +
      "Overall Budget Remaining (if set): " + (s.overallBudgetRemaining != null ? s.overallBudgetRemaining : "N/A") + "\n";

    MailApp.sendEmail(SUMMARY_EMAIL_RECIPIENT, subject, body);
  } catch (err) {
    logError("sendMonthlyEmailSummary", err, null);
    throw err;
  }
}