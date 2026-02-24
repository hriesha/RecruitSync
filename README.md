# RecruitSync

Tired of losing track of where you applied? This script scans your Gmail for job application emails and automatically logs them to a Google Sheet — company, job title, date, and status. When a follow-up arrives (interview, rejection, offer), it updates the row automatically.

No installs, no APIs, no accounts. Just Google Apps Script.

---

## What it tracks

| Column | What it does |
|---|---|
| Company | Pulled from the email subject or sender domain |
| Job Title | Pulled from the subject/body (left blank if it can't find it) |
| Date Applied | Date of the first confirmation email |
| Status | `Applied` → `Interview` → `Offer` or `Rejected` |
| Thread ID | Used to avoid duplicates — you can hide this column |
| Last Updated | When the status last changed |

---

## Setup (takes ~2 minutes)

**1. Create a Google Sheet**

Go to [sheets.google.com](https://sheets.google.com) and make a new spreadsheet. Name it whatever you want.

**2. Open Apps Script**

In your spreadsheet, go to **Extensions → Apps Script**.

**3. Paste the code**

Delete everything in the editor (the default `myFunction` stub), then paste in the contents of `Code.gs` from this repo. Hit **Save** (Cmd+S / Ctrl+S).

**4. Run it**

Click the function dropdown (it'll say `scanJobApplications`) and hit **Run**. Google will ask you to approve access to Gmail and Sheets — that's expected, go ahead and allow it.

That's it. Go back to your spreadsheet and you'll see a "Job Applications" tab filled in.

---

## Tips

- **First run is slow** — it searches the last 180 days by default. After that, change `DAYS_BACK` at the top of `Code.gs` to something like `30` so it's faster.
- **Job title is blank sometimes** — not every confirmation email includes it in a parseable format. You can fill those in manually.
- **Status not updating?** — follow-up emails need to be in the same Gmail thread as the original confirmation. If a company starts a new thread, the script won't connect them automatically (yet).

---

## Enable auto-scanning (runs every morning at 8am)

After your first manual run, you can set it to scan automatically so you never have to think about it again.

In the Apps Script editor, select `createDailyTrigger` from the function dropdown and hit **Run**.

That's it. It'll scan your Gmail every morning at 8am and update the sheet in the background.

To turn it off, run `removeDailyTrigger` the same way.
