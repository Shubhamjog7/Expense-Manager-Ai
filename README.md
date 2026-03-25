# 📊 Expense Manager – AI‑Powered Finance Dashboard

A personal finance dashboard built using **Google Apps Script** that tracks expenses, savings, income, and generates **AI‑powered financial insights**.

This project focuses on **clean financial modeling**, **data stability**, and **explainable AI insights** rather than flashy visuals.

---

## ✨ Features

### ✅ Transaction Management
- Supports:
  - Expenses (EXP)
  - Credit Card Spends (CC)
  - Bill Payments (BILL)
  - Savings (SAV)
  - Income (CR)
- Accurate Debit / Credit / Net calculation
- Monthly and daily summaries

### ✅ Category → Sub‑Category System
- Hierarchical classification
- Central taxonomy control via `CATEGORY_CONFIG`
- Safe merge / rename of sub‑categories without breaking historical data

### ✅ Daily Summary
- Day‑wise Debit, Credit, and Net
- Correct accounting logic:
  - Only `CR` contributes to Credit
  - All other types contribute to Debit

### ✅ AI Finance Coach
- AI‑generated monthly insights
- Explains *why* spending changed
- Uses sub‑category intelligence:
  > “Food overspend was mainly driven by Swiggy/Zomato (₹4,200).”
- Copy & PDF export support

### ✅ Search & Filter
- Filter by date, amount, category, sub‑category, card, type
- Quick action shortcuts (Last 7 days, Only EXP, etc.)

---

## 🧠 Architecture Overview
Google Sheet (Data)
│
├── RAW_DATA
├── CATEGORY_CONFIG
├── SUBCATEGORY_BUDGET (optional)
│
└── Google Apps Script
├── Code.gs      (backend logic)
└── Index.html   (frontend UI)

---

## 📄 Data Model

### RAW_DATA (Source of Truth)

| Index | Column | Description |
|------:|--------|------------|
| 0 | Timestamp | Date of transaction |
| 1 | Type | EXP / CC / BILL / CR / SAV |
| 2 | Amount | Transaction amount |
| 3 | Category | Top‑level category |
| 4 | SubCategory | Drill‑down |
| 5 | Notes | Free text |
| 6 | Card | Card name |
| 7 | Tag | Optional (not used for logic) |

### Accounting Rules (Locked)

| Type | Debit | Credit |
|------|------|--------|
| EXP | ✅ | ❌ |
| CC | ✅ | ❌ |
| BILL | ✅ | ❌ |
| SAV | ✅ | ❌ |
| CR | ❌ | ✅ |

> Debit / Credit is derived strictly from **Type**, not Tag.

---

## ⚙️ Backend Logic

Key backend function:
- `getMonthlyData(monthKey, cardFilter)`

Daily Summary logic:
```js
if (type === "CR") {
  credit += amount;
} else {
  debit += amount;
}


🤖 AI Integration
The AI receives:

Monthly summary
Category totals
Daily debit/credit
Sub‑category totals

Rules enforced in prompt:

Only top sub‑category mentioned
No invented numbers
Uses ₹ consistently
Falls back gracefully if data missing


🚀 Deployment

Built as a Google Apps Script Web App
Editor code ≠ deployed version
Always deploy via:

Deploy → New deployment


Rollbacks handled via Project History


📈 Project Status
✅ Stable v1.0
✅ Clean data model
✅ Correct financial logic
✅ AI insights enabled
Future enhancements:

Budget breach explanations (AI)
Alerts & notifications
Goal planning
Visual drill‑downs (isolated cards)


🧑‍💻 Learning Outcomes
This project demonstrates:

Financial data modeling
Schema evolution handling
UI stabilization techniques
AI prompt discipline
Professional rollback practices