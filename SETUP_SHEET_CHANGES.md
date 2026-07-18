# Google Sheet Changes Required for New Features

This document explains what changes you need to make in your Google Spreadsheet and Google Apps Script so that the new features start working with cloud sync.

---

## Feature: "How was your day?" (Star Rating + Summary)

Currently this feature saves data only in your browser (localStorage). To sync it with your Google Sheet so that data persists across devices and isn't lost when you clear browser data, follow these steps:

---

### Step 1: Create a New Sheet Tab in Google Sheets

1. Open your Google Spreadsheet.
2. At the bottom, click the **"+"** button to add a new sheet tab.
3. Rename it to exactly: **`DayReflections`**

---

### Step 2: Add Headers in Row 1

In the **DayReflections** sheet, type these headers in Row 1:

| Column A | Column B | Column C | Column D | Column E | Column F |
|----------|----------|----------|----------|----------|----------|
| `reflection_id` | `date` | `rating` | `summary` | `reflection` | `notice` |

- **reflection_id** — Unique ID for each entry
- **date** — The date of the reflection (e.g., `2026-07-14`)
- **rating** — Star rating from 0.5 to 5 (supports half-stars)
- **summary** — Short text summary of the day
- **reflection** — Deeper reflection on the day (optional)
- **notice** — "Did you notice?" observation (optional)

---

### Step 3: Update Google Apps Script

Open your Google Sheet → **Extensions** → **Apps Script** and make the following additions to your existing code:

#### 3a. Add the sheet name constant at the top (near the other constants):

```javascript
var REFLECTIONS_SHEET = 'DayReflections';
```

#### 3b. Add a new case in the `doGet()` function (inside the `try` block):

```javascript
} else if (action === 'getDayReflection') {
  return getDayReflection(doc, e);
}
```

#### 3c. Add a new case in the `doPost()` function (inside the `try` block):

```javascript
} else if (action === 'saveDayReflection') {
  return saveDayReflection(doc, data);
}
```

#### 3d. Add these two new functions at the bottom of the script (before the `jsonResponse` helper):

```javascript
// ==================== DAY REFLECTION OPERATIONS ====================

function saveDayReflection(doc, data) {
  var sheet = doc.getSheetByName(REFLECTIONS_SHEET);
  if (!sheet) {
    sheet = doc.insertSheet(REFLECTIONS_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([['reflection_id', 'date', 'rating', 'summary']]);
  }

  var dateStr = data.date; // e.g., "2026-07-14"

  // Check if today's reflection already exists — update it
  var allData = sheet.getDataRange().getValues();
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][1] === dateStr) {
      // Update existing row
      sheet.getRange(i + 1, 3).setValue(data.rating);
      sheet.getRange(i + 1, 4).setValue(data.summary || '');
      return jsonResponse({ result: 'success', updated: true });
    }
  }

  // Otherwise add new row
  var row = [
    data.reflection_id || Utilities.getUuid(),
    dateStr,
    data.rating,
    data.summary || ''
  ];

  sheet.appendRow(row);
  return jsonResponse({ result: 'success', updated: false });
}

function getDayReflection(doc, e) {
  var sheet = doc.getSheetByName(REFLECTIONS_SHEET);
  if (!sheet) return jsonResponse({ result: 'success', reflection: null });

  var dateStr = e.parameter.date; // e.g., "2026-07-14"
  var allData = sheet.getDataRange().getValues();

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][1] === dateStr) {
      return jsonResponse({
        result: 'success',
        reflection: {
          reflection_id: allData[i][0],
          date: allData[i][1],
          rating: parseInt(allData[i][2]),
          summary: allData[i][3]
        }
      });
    }
  }

  return jsonResponse({ result: 'success', reflection: null });
}
```

---

### Step 4: Re-deploy the Apps Script

After making the above code changes:

1. Click **Deploy** → **Manage deployments**
2. Click the **pencil icon** (edit) on your existing deployment
3. Under **Version**, select **"New version"**
4. Click **Deploy**

> ⚠️ You must create a new version every time you change the code. The old URL will automatically use the new version.

---

### Step 5 (Optional): Update Frontend to Sync with Sheet

The current frontend saves reflections to localStorage only. If you want it to also sync with the Google Sheet, replace the save logic in `index.html`:

In the `btn-save-day-rating` click handler, **after** the localStorage save, add:

```javascript
// Also save to Google Sheet
const dateFormatted = new Date().toISOString().split('T')[0]; // "2026-07-14"
apiPost({
  action: 'saveDayReflection',
  reflection_id: generateUUID(),
  date: dateFormatted,
  rating: dayRating,
  summary: summary
});
```

---

## Summary Checklist

| # | Task | Where |
|---|------|-------|
| 1 | Create **DayReflections** sheet tab | Google Sheets |
| 2 | Add headers: `reflection_id`, `date`, `rating`, `summary` | Row 1 of DayReflections |
| 3 | Add `REFLECTIONS_SHEET` constant | Apps Script (top) |
| 4 | Add `getDayReflection` case in `doGet()` | Apps Script |
| 5 | Add `saveDayReflection` case in `doPost()` | Apps Script |
| 6 | Add `saveDayReflection()` and `getDayReflection()` functions | Apps Script (bottom) |
| 7 | Re-deploy as new version | Apps Script Deploy menu |
| 8 | (Optional) Add `apiPost` call in frontend | index.html |

---

## Notes

- The **mobile navigation** and **vertical scroll** features are purely frontend (HTML/CSS/JS) — they require **no changes** to the Google Sheet or Apps Script.
- The "How was your day?" feature already works offline via localStorage. The Sheet sync is an enhancement for cross-device persistence.
