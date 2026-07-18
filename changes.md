# Changes Log

## 2026-07-18

### Feature: Self Reflection Screen (New Tab)
- Added a dedicated "Self Reflection" tab in both desktop nav and mobile bottom nav.
- New screen allows users to save a daily reflection with:
  - Half-star rating (0.5 to 5)
  - Summary text
  - Reflection text (what went well / could improve)
  - "Did you notice?" optional observation
- "Previous Reflections" section shows all saved reflections from Google Sheet (newest first).
- List has vertical scroll and auto-loads when navigating to the tab.
- Data syncs with the `DayReflections` sheet in Google Sheets.
- Added `getAllReflections` GET endpoint in Apps Script.

## 2026-07-14

### Fix: Mobile Navigation Redesign — Bottom Nav Bar
- The old horizontal scrolling tab bar was hard to use on mobile (tabs hidden offscreen).
- Added a **fixed bottom navigation bar** visible only on mobile (below `md` breakpoint).
- Each tab has an icon + label for easy one-thumb access.
- "Add New" button has a raised floating circle (FAB style) for emphasis.
- Top tab bar is now hidden on mobile, visible only on `md+` (desktop/tablet).
- Added `pb-24` bottom padding to main content on mobile so content isn't hidden behind the bottom nav.
- Supports iOS safe area with `env(safe-area-inset-bottom)`.
- Both navs stay in sync — clicking either updates both active states.

### Feature: "How was your day?" — Star Rating & Day Summary
- Added a new section at the bottom of the Dashboard (Today) screen.
- Users can rate their day from 0.5 to 5 stars (half-star precision) with hover preview.
- Click left half of a star = half star, click right half = full star.
- Labels for each rating level (Terrible → Amazing day!).
- Optional short text summary of the day.
- Optional "Reflection" field — for deeper thoughts on what went well/could improve.
- Optional "Did you notice?" field for observations, patterns, or small wins.
- Reflection is saved per day in localStorage AND synced to Google Sheet (`DayReflections` tab).
- If already saved today, the form hides and shows the saved rating + all fields.
- Resets daily — fresh form appears each new day.

### Feature: Vertical Scroll for Large Record Lists
- Dashboard revision cards container: capped at `60vh` with vertical scroll.
- All Entries list: capped at `70vh` with vertical scroll.
- Captures list: capped at `60vh` with vertical scroll.
- Added subtle custom scrollbar styling (thin, rounded, gray thumb) for a clean look.
- `pr-1` padding-right added to prevent scrollbar overlapping content.
