# Revision Tracker - Bugs Solved & Future Notes

## Bugs Fixed During Development

### Bug 1: CORS Preflight Error (405 Method Not Allowed)

**Problem:**  
When sending POST requests with `Content-Type: application/json`, the browser sends an `OPTIONS` preflight request first. Google Apps Script does NOT handle `OPTIONS` requests, resulting in `405 Method Not Allowed`.

**Root Cause:**  
Any cross-origin request with `Content-Type: application/json` is considered a "non-simple" request by the browser, which triggers a CORS preflight check.

**Fix:**  
Changed `apiPost()` to send data using `FormData` instead of raw JSON. FormData uses `multipart/form-data` content type which is a "simple request" — no preflight needed.

```javascript
// Before (broken)
fetch(SCRIPT_URL, {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(data),
});

// After (working)
const formData = new FormData();
formData.append('payload', JSON.stringify(data));
fetch(SCRIPT_URL, {
  method: 'POST',
  body: formData,
});
```

---

### Bug 2: `Cannot read properties of undefined (reading 'type')`

**Problem:**  
After fixing Bug 1, the Apps Script threw `TypeError: Cannot read properties of undefined (reading 'type')` because the code tried to access `e.postData.type`.

**Root Cause:**  
When data is sent as `FormData` (multipart), Google Apps Script places the fields in `e.parameter` instead of `e.postData`. So `e.postData` can be `undefined`.

**Fix:**  
Updated `doPost()` to check `e.parameter.payload` first, then fall back to `e.postData.contents`:

```javascript
var data;
if (e.parameter && e.parameter.payload) {
  data = JSON.parse(e.parameter.payload);
} else if (e.postData && e.postData.contents) {
  data = JSON.parse(e.postData.contents);
} else {
  return jsonResponse({ result: 'error', error: 'No data received' });
}
```

---

## Things to Take Care of in the Future

### 1. Redeployment Required After Script Changes
Every time you edit `google_app_script.js` in the Apps Script editor, you MUST create a **new version** when deploying. Go to:  
**Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**  
Just saving the script is NOT enough.

### 2. Google Apps Script Quotas
- **Execution time limit:** 6 minutes per call.
- **URL Fetch calls:** 20,000/day (free tier).
- **Sheet read/write:** Can be slow with large datasets (1000+ rows). Consider pagination if the sheet grows large.

### 3. Security Considerations
- The credentials are hardcoded and static (`mnnit4abhishek@gmail.com` / `India@123`). This is fine for personal use but NOT suitable for production.
- The Web App URL is publicly accessible (set to "Anyone"). Anyone with the URL can call your API. Keep it private.
- Never commit the Apps Script Web App URL to a public repository.

### 4. Data Parsing Edge Cases
- If a user enters JSON-like characters in notes/titles (quotes, brackets), it could break `JSON.parse()`. The current code handles this for `links` and `intervals` fields but be cautious with free-text fields.

### 5. Sheet Structure Changes
- Do NOT rename, reorder, or delete columns in the Google Sheet manually. The script depends on exact column positions (A=entry_id, B=subject, etc.).
- If you need to add columns, always add them AFTER column J.

### 6. localStorage Dependency
- "Completed today" state is stored in the browser's localStorage. Clearing browser data resets daily progress tracking.
- Session persistence ("Remember this device") also uses localStorage.

### 7. Concurrent Access
- `LockService` is used to prevent race conditions, but if multiple users mark the same entry "done" simultaneously, the last write wins.
- The sheet-based approach doesn't support real-time collaboration well.

### 8. Timezone Issues
- Date comparisons for "today's revisions" use the browser's local timezone. If you travel across timezones, some entries might appear a day early/late.
- The Apps Script also uses your Google account's timezone (configurable in Apps Script project settings).

### 9. Error Handling
- Network failures show a generic toast. For better UX, consider adding retry logic or offline queue.
- If the Google Sheet is deleted or renamed, all API calls will fail silently.

### 10. Future Improvements to Consider
- Add a loading skeleton UI instead of just a spinner.
- Implement offline support with service workers.
- Add data export (CSV/JSON) feature.
- Add a "snooze" option to delay a revision by 1 day.
- Move to Firebase or Supabase if the app outgrows Google Sheets.
