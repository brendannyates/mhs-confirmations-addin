# MHS Confirmations

An Excel add-in and a scheduled Power Automate flow. They find appointments on the `List` sheet that still need a confirmation between today and two business days out, group them by clinic, and email each opted-in clinic's help mailbox at 8:00 and 16:00 Pacific on weekdays.

Each email row shows: Md Rc (MRN), Pat Name as a LASFIR code, Event, Appt Dt, Beg Tm, Rendering, Create Date, the four outreach status columns (K to N), and a "View" link to the row.

This repo holds code only. It contains no patient data and no link to the workbook. The workbook link is stored only inside the workbook, in `Confirm Settings!G2`.

## Files
| File | Purpose |
|---|---|
| `index.html`, `app.js`, `core.js`, `style.css`, `assets/` | The task pane, served by GitHub Pages |
| `manifest.xml` | What you upload to Excel to install the pane |
| `office-scripts/ConfirmationsRun.ts` | The script the scheduled flow runs |

---

## Step-by-step integration

### Step 1: Check the hosting (2 minutes)
1. Open https://brendannyates.github.io/mhs-confirmations-addin/index.html in a browser. It should show a page titled "Confirmations needed". Outside Excel it stays on "Loading…", which is expected.
2. Open https://brendannyates.github.io/mhs-confirmations-addin/manifest.xml. It should show XML.
3. If either page shows a 404, go to the repo's Settings > Pages, confirm Source is "Deploy from a branch" with `main` and `/ (root)`, and wait two minutes.
4. Save `manifest.xml` to your computer (right-click the page > Save as), or download it from the repo. You need it in Step 2.

### Step 2: Install the pane in Excel
1. Open the workbook from SharePoint in **Excel on the web**, not a downloaded copy.
2. Go to Home > Add-ins > More Add-ins > My Add-ins > **Upload My Add-in**, and choose `manifest.xml`.
   - If Upload is missing, your tenant blocks it. Ask IT to deploy `manifest.xml` under M365 admin center > Settings > Integrated apps.
   - Uploading installs it for you only. Each Practice Leader who wants the pane repeats this step, or IT deploys it to the group.
3. Click **Confirmations** on the Home ribbon. The pane opens on the right.

### Step 3: First run of the pane
1. On first open the pane creates a `Confirm Settings` sheet: clinics and emails in A:C (Opt In all "No"), holidays in E, workbook link in G.
2. Check the top line: a date window (today to two business days out) and counts.
3. Read every yellow warning. They flag unmapped locations, unreadable dates, bad names, a missing workbook link and an out-of-date report.
4. Open one clinic card and click **Preview email**. Check that Md Rc is the first column, plus the To address and the subject.
5. The "Workbook link" box usually fills itself. If it's blank, paste the workbook's link (Share > Copy link). The pane saves it to `Confirm Settings!G2`.
6. Click a "View" link from a previewed email in a browser to confirm it opens the workbook. It may not jump to the exact row in Excel on the web.

### Step 4: Settings
1. On `Confirm Settings`, Practice Leaders change column C to **Yes** for each clinic that should get emails. Nothing sends while a clinic is "No".
2. Optional: list holidays in column E, one per row, as YYYY-MM-DD. They're skipped when counting business days.

### Step 5: Add the Office Script
1. In the workbook, go to **Automate > New Script**. If the Automate tab is missing, an admin has to enable Office Scripts.
2. Delete the starter code. Paste in all of `office-scripts/ConfirmationsRun.ts` (open it in the repo, click "Raw", select all, copy).
3. Rename the script exactly **Confirmations Run** and save.
4. Note: the script is stored in your OneDrive (Documents > Office Scripts), not in the workbook. The flow depends on your account.

### Step 6: Build the flow (make.powerautomate.com, signed in as the workbook owner)
1. **Create > Scheduled cloud flow.** Name it "Confirmations Run". Click Create.
2. **Recurrence** trigger: Interval 1, Frequency **Week**, Time zone **(UTC-08:00) Pacific Time**, On these days **Mon to Fri**, At these hours **8, 16**, At these minutes **0**.
3. **+ New step > Run script** (Excel Online (Business)):
   - Location: the SharePoint site. Document Library: the library. File: the workbook.
   - Script: **Confirmations Run**.
   - todayIso: open the Expression tab and paste
     `formatDateTime(convertTimeZone(utcNow(),'UTC','Pacific Standard Time'),'yyyy-MM-dd')`
4. **+ New step > Parse JSON**:
   - Content: `result` from Run script.
   - Schema: paste this.
     ```json
     {"type":"object","properties":{"status":{"type":"string"},"message":{"type":"string"},"ownerHtml":{"type":"string"},
     "warnings":{"type":"array","items":{"type":"string"}},
     "emails":{"type":"array","items":{"type":"object","properties":{"clinic":{"type":"string"},"to":{"type":"string"},"subject":{"type":"string"},"html":{"type":"string"},"count":{"type":"integer"}}}}}}
     ```
5. **+ New step > Condition**: `status` (from Parse JSON) **is equal to** `ok`.
6. In **If yes**:
   1. **Apply to each**, over `emails`. Inside it, **Send an email (V2)** (Office 365 Outlook): To = `to`, Subject = `subject`, Body = `html`. Under Advanced options, set Is HTML = **Yes** if shown. (To send from a shared mailbox, use "Send an email from a shared mailbox (V2)" instead.)
   2. After the loop (still in If yes), add another **Condition**. Left side, Expression: `length(body('Parse_JSON')?['warnings'])`. Operator **is greater than**. Right side `0`.
      In its If yes: **Send an email (V2)** to yourself. Subject `Confirmations Run: warnings`, Body = `ownerHtml`, Is HTML = Yes.
7. In **If no**: **Send an email (V2)** to yourself. Subject `Confirmations Run: nothing sent`, Body = `ownerHtml`, Is HTML = Yes.
8. Save.

If you renamed the Parse JSON step, use its name in `body('...')` with spaces replaced by underscores.

### Step 7: Test before going live
1. In the clinic send step (inside Apply to each), temporarily set To to **your own address**. Set one or two clinics to **Yes**.
2. Make sure `List` has current data (the data analyst maintains it).
3. In the flow, click **Test > Manually > Run flow**. Check the email: Md Rc, LASFIR codes, dates, times, and the View link.
5. **Coverage test:** paste a report that ends before the window's last day and run. You should get a "nothing sent" email.
6. **Warnings test:** add a row inside the window with a made-up Loc Name and run. You should get a warnings email naming it. Delete the row.
7. Set To back to the `to` value, save, and make sure the flow is turned on. Confirm the next 8:00 run arrives.

---

## Daily routine
Nothing to do. The data analyst maintains `List`, and each run (8:00 and 16:00) uses whatever is in `List` at that moment. To add or remove a clinic, change column C on `Confirm Settings`.

## Rules
- Window: today through WORKDAY(today, 2), skipping weekends and the holidays in column E.
- A row is skipped if K:N contain "confirmed"/"conf" or "rescheduled" (but not "unconfirmed"/"not confirmed"), or "cancel".
- Dates and times may be real Excel values or text (`10/14/2026`, `2026-10-14`, `9:30 AM`). Rows with unreadable dates are reported as warnings.
- LASFIR: first 3 letters of the last name + first 3 of the first name; anything in parentheses is ignored.
- Clinics and their emails come from `Confirm Settings` columns A and B. A List `Loc Name` matches a clinic when it starts with the clinic name minus its " - XX" state (whole word): "Castro - CA" matches "Castro Street", "Oakland - CA" matches "Oakland IV Ketamine Clinic". The longest match wins. Unmatched locations are reported in the warnings email, never silently dropped.
- The run refuses to send only if the List headers have moved. If List doesn't reach the window's end, it still sends and adds a warning.
- The 16:00 run re-sends anything still unconfirmed from 8:00. Run times are Pacific, so TX clinics get emails at 10:00 and 18:00 Central.

## Security
- Whatever is on `main` runs inside the workbook with read/write access, and pushes go live within minutes. Keep two-factor sign-in on for this GitHub account and limit who can push.
- Staff should not type patient names into K:N, since that text is copied into the emails as is.

## Updating later
- Pane changes: push to `main`; they go live by themselves.
- `manifest.xml` changes: re-upload it in Excel (Step 2).
- Script changes: re-paste `ConfirmationsRun.ts` into the existing **Confirmations Run** script and save.
