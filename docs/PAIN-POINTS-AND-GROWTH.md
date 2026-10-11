# DocBrisk for Excel: why users did not stay, and what 3.2 changes

## 1. What Excel users struggle with (and pay to fix)

These are the complaints that come up again and again from finance, accounts, MIS and operations people. They are
based on the common questions in Excel forums and help sites, not on DocBrisk's own user data.

| # | Pain point | How often | DocBrisk before 3.2 | 3.2 |
|---|---|---|---|---|
| 1 | **#N/A from VLOOKUP / XLOOKUP**: trailing spaces, non-breaking spaces from web/SAP exports, numbers stored as text, duplicate keys taking the first match silently | Every day | Nothing direct (only FUZZYLOOKUP for names) | **Smart Lookup** + `=SMARTLOOKUP` (free): matches through these differences, labels every row, counts the repeated keys |
| 2 | Messy exports (SAP/Tally/bank): text numbers, `1,234.50-`, repeated headers, merged cells | Every day | Smart Clean, Report to Table (good) | Unchanged, now followed by a tip to Smart Analyst |
| 3 | Reconciling two lists (books vs statement, GSTR-2B, bank) | Weekly/monthly | Reconcile, Compare (Pro) | Compare Two Lists (free) now leads to Reconcile |
| 4 | Building the same monthly report and commentary again | Monthly | Dashboards (free), Smart Analyst (Pro) | Dashboard / Clean / Summary now lead to Smart Analyst / Why It Changed |
| 5 | Names that do not match (Pvt/Private, typos) | Weekly | Fuzzy Merge, Fuzzy Duplicates (Pro) | Smart Lookup estimates how many of the missing keys Fuzzy Merge would match |
| 6 | Sending each party its own rows | Monthly | Mail to Each Party (Pro) | Split to Sheets / Email Sheet / Web Dashboard lead to it |
| 7 | "I don't know which tool to use" (79 buttons) | Always | Assistant, Quick Find | My Impact shows the tools that fit the user's own work |

## 2. Why the add-in was not loved

1. **The value was invisible.** A tool ran and nothing said "this saved you 40 minutes". People forget what they did not
   have to do, so they never felt the add-in was worth money.
2. **No reason to come back.** Nothing changed from one day to the next: no progress, no streak, no reward.
3. **The free/Pro split felt like a wall, not a ladder.** Pro was only mentioned at the moment a tool refused to run
   (a block, which annoys). Users never had a week with Pro to get used to it, so there was nothing to lose.
4. **The headline free tools were rare jobs** (dashboards) while the daily job, the lookup, had no tool.
5. **It had never been compiled** when shipped (README of 3.0/3.1). Every crash on first use costs trust.

## 3. The habit loop in 3.2

```
 Trigger            Action                 Reward                          Investment
 -----------------  ---------------------  ------------------------------  ------------------------------
 Status bar at      Smart Lookup / Clean   "Saved ~12 min", streak +1,     Milestones and bonus runs
 Excel start:       / any tool             milestone card with bonus       collected; Pro Week usage
 streak, hours,                            Pro runs                        that would be lost
 bonus runs
```

- **Daily trigger:** the greeting on the status bar when Excel opens (streak, hours saved, Pro Week day, bonus runs).
- **Daily action that is free:** Smart Lookup is the most frequent job in Excel and stays free, so the add-in is opened
  every day.
- **Variable reward:** milestones arrive at different moments (runs, tools, streaks, hours), each with bonus Pro runs.
- **Investment:** the numbers in My Impact grow; the streak is something to keep.

## 4. The path from Free to Pro

1. **Pro Week (reverse trial), days 1-7:** every Pro tool open. People get used to Smart Analyst, Reconcile and Fuzzy
   Merge on their real files. The card the day before the end and the card after the end show *their own* Pro runs and
   minutes saved (loss aversion with real numbers, nothing invented).
2. **Free runs, then bonus runs:** 3 free runs per Pro tool, then bonus runs earned with milestones. Using DocBrisk more
   gives more Pro tastes, so habit leads to more Pro use.
3. **Pro advantages shown in between, at the right moment and rarely:**
   - after a free tool, the Pro tool that is the natural next step of *that* job, with a "Try" button
     (at most 2 a day, 3 runs apart, the same tip at most 3 times, "No tips" switch);
   - after Smart Lookup, a concrete number: "about 41 of the 70 keys not found look like the same names spelled
     differently";
   - on the last free runs of a Pro tool: what Pro members get, and how many minutes this run saved;
   - in the Pro window and My Impact: "your time saved is worth about ₹X; Pro costs ₹2.8 a day".
4. **Retention for Pro members:** at most once a week, after a Pro tool, the hours Pro has saved them, so the value is
   seen before the renewal date (the 7-day renewal reminder was already there).

## 5. Pricing notes

- Show the yearly plan per day (₹999 / year = ₹2.8 a day). This is now done everywhere the price is shown.
- Keep the monthly plan as the anchor that makes the yearly plan look good ("Saves ₹189 a year").
- Ideas not built (they need the website): a team plan (5 PCs), a discount code inside the "Pro Week has ended" card
  valid for 48 hours, and a referral reward (a friend installs -> both get bonus runs). These would need the Worker on
  docbrisk.com to issue the codes.

## 6. What to measure

DocBrisk does not send data anywhere, so measure on the website and by asking:

- Conversion from install to Pro key within 30 days, before and after 3.2 (count of keys issued / downloads).
- Share of renewals.
- A short optional question on the download page: "Which tool do you use most?"
- If you ever add opt-in usage statistics, the useful numbers are: day-7 and day-30 return rate, runs per active day,
  Pro tool runs during the Pro Week, and which tip leads to a purchase.

## 7. Next steps

1. Run 3.2 once in Excel on Windows (100% and 150% screen scaling): My Impact, the tip card, Smart Lookup on a sheet.
2. Update the website: lead with "VLOOKUP without the #N/A" (free) and "7 days of Pro on install".
3. Check the minutes per tool in `Habit.cs` against what users tell you; keep them on the careful side.
