# How the Invoice reader works (as of 13 Oct 2026)

The Google Sheet "Invoice reader" (in Scans) reads scanned supplier invoices with Claude, turns them into Sofia journal lines, and exports a file Sofia can import. The code is `InvoiceReader.gs` plus `ExampleSeed.gs` (starting examples; sample data in this repo) in the sheet's Apps Script.

---

## 1. What a person does

**Every day (staff):**
1. Scan invoices. The scanner saves them straight into **Scans → Invoice reader → Inbox**.
2. Open the **Needs Check** tab and clear each row:
   - **Journal row:** fix the value in the Journal, then set its Status to **OK**.
   - **Invoices row** (statement, duplicate, unreadable file): deal with it, then change its Status (e.g. "Checked").
   - **Missing code:** add it in Suppliers / Expense accounts / Restaurants, then **Invoices → Refresh Sofia codes**.
3. **Invoices → Export for Sofia**, then import the file into Sofia.

**About once a week (you):**
- **Lessons tab:** read any **Pending** lessons and set Use to **Yes** or **No**. It takes about a minute.

**When something new comes up (you):**
- A new supplier, account or code: add it to the code tabs, or tell Claude.
- **Still open:** the "Other" account, car numbers (T3), how to spot CSS invoices, restaurant stamp photos, the fixed-asset amount (now €1,000), and the Sofia import test.

That's all. Everything below happens by itself.

---

## 2. What happens automatically

1. **Every hour** (or every 2 hours, whichever you started) the script picks up new scans from the Inbox (nothing else is checked).
2. **Opus 5.5 reads each scan.** One file can hold many invoices. For each invoice it gives:
   - supplier, number, date and restaurant;
   - net/VAT per category and the total.
3. **For each invoice, it always tries for the invoice number.** If unsure, it guesses the most likely number, marks it "best guess" and the cell turns yellow. A guessed invoice still goes to Sofia.
4. **Second read:** if any field is unsure, **Fable 5.1** reads those pages again without seeing Opus's answer.
   - If they agree, the invoice is OK.
   - If they disagree, it goes to Needs Check showing both readings. An invoice number that both had to guess is the exception: the first guess is kept and it goes through.
5. **Checks:**
   - net + VAT = total;
   - not a duplicate (same supplier + same number);
   - the restaurant is clear;
   - every Sofia code is known;
   - equipment over €1,000 is labelled FIXED ASSET;
   - car expenses need a car number.
6. **Codes are filled in:**
   - supplier account;
   - expense account (some suppliers are locked to one, e.g. Bolt is always 84001);
   - T1 restaurant code; storage/kitchen use 801 and the 32xxx accounts;
   - VAT code;
   - company code: SHR, or SHE for EXPRESS.
7. **Results:**
   - everything goes into **Journal** and **Invoices**;
   - ready invoices go to the **Sofia** tab;
   - everything else goes to **Needs Check**.
8. **The scan file moves** to Done or Needs check, so it is never read twice.
9. **On export:** references are shortened to Sofia's 15 characters, and the export is blocked if a company code is missing or TEST- codes remain. Exported rows are stamped and never exported twice.
10. **Learning:**
    - every hand correction is logged in **Corrections**;
    - fixing a Needs Check invoice (Status → OK) makes it a **verified example** in **Examples** automatically;
    - once a day, Claude suggests **lessons** from the corrections, which wait for approval;
    - if staff move a supplier to the same account twice, that supplier is locked to that account.

---

## 3. The tabs

| Tab | What it is | Who touches it |
|---|---|---|
| **Help** | What every tab, colour and message means | Read only |
| **Needs Check** | Everything not going to Sofia yet, with where to fix it. Red tab = something is waiting | Staff, daily |
| **Sofia** | Ready invoices in Sofia's import layout | Look over before export |
| **Journal** | Accounting lines per invoice: expenses DR, VAT 42001 DR, supplier CR. Fix values here | Staff, when fixing |
| **Invoices** | One row per document read: what Claude read, status, notes. **Verified** tick box at the end | Read; tick is optional |
| **Suppliers** | Supplier name → Sofia account. Column E "Always this category" locks a supplier to one expense account | Add new suppliers |
| **Expense accounts** | Category → Sofia account, plus the storage (32xxx) account | Add missing codes |
| **Restaurants** | Restaurant → T1 code, how Claude recognises it, company code (column D) | Rarely |
| **VAT codes** | 0, 5, 9, 19 | Never |
| **Examples** | Trusted, human-checked invoices: the correct reading, Claude's original, what was corrected. At most 5 active per supplier. Starts with the 197 you checked on 9 Oct | Rarely: set Status to Retired if one is wrong |
| **Corrections** | Automatic log of every hand edit in Journal/Invoices | Nobody |
| **Lessons** | Lessons Claude suggests from corrections. Only **Yes** ones are used | You, weekly |
| **Log** | Every read: time, file, cost in $, result. "API limit reached (429)" means the key hit its limit | When something goes wrong |
| **Test run** | Results of a test re-read (only after you run one) | You, when testing |
| Checked scans *(hidden)* | Which Scans files were already looked at | Nobody |

**Tab order and hidden tabs:** the accountant sees only Needs Check, Journal, Invoices, Sofia, Help, Suppliers, Expense accounts and Restaurants, in that order. Lessons, Examples, Corrections, Log and VAT codes are hidden; use **Invoices → Setup → Show admin tabs** (for example to approve lessons), then **Hide admin tabs**. Set up sheet and folders re-applies this.

**There's no "Supplier profiles" tab.** The profiles are built from the Examples tab each time a run starts (one line per supplier: number format, usual category, VAT rate, past corrections, notes) and given to Claude in its instructions. To see or change what a profile says, look at or edit that supplier's rows in **Examples**.

**Status column colours (Journal and Invoices):**
- **Green "Exported …":** already sent to Sofia.
- **Blue "OK":** ready, waiting for the next export.
- **Red "Needs check …":** waiting for a person.
- **Grey (anything else, e.g. "Checked"):** dealt with by hand.

**Cell colours:**
- **Yellow:** guessed invoice number (still goes to Sofia).
- **Grey:** missing Sofia code.
- **Red:** couldn't read it, or needs checking.

---

## 4. The menu (Invoices)

**Top level (daily use):**

| Item | What it does |
|---|---|
| Help | Opens the Help tab |
| Read new invoices now | Runs immediately instead of waiting for the timer |
| Refresh Sofia codes | After adding codes, fills them into existing Journal lines |
| Update Sofia and Needs Check tabs | Rebuilds both from the Journal |
| Export for Sofia | Saves the CSV in Invoice reader → Sofia exports and stamps the rows as exported |

**Setup** (rarely):

| Item | What it does |
|---|---|
| Set up sheet and folders | Creates or updates tabs and folders. Run it after every script update |
| Set Claude API key | Stores the key in the script (never in a cell or chat) |
| Set Sofia company code | Only if it isn't SHR |
| Show / Hide admin tabs | Shows or hides Lessons, Examples, Corrections, Log, VAT codes |
| Start automatic reading (every hour / every 2 hours), Stop | The timer |
| Learn from corrections now | Runs the daily learning step now |
| Supplier profiles: switch on / off | Whether Claude gets the supplier profiles. **Off** until the Test run shows it helps |

**Testing and one-off:**

| Item | What it does |
|---|---|
| Test run: re-read Done and Needs check, Stop the test run | Re-reads old scans into the Test run tab only. Changes nothing else |
| Load the Sofia codes Claude has | Fills the code tabs with everything collected so far (only empty cells) |
| Load starting examples (9 Oct answer key) | One-off: loads the 197 checked invoices into Examples |
| Fill / Remove test codes | Fake codes for previewing only. Remove before exporting |
| Start fresh | Empties Journal, Invoices, Sofia, Needs Check, Log, Corrections, Examples, Lessons, Test run. Keeps the code tabs, settings and the scan files. Asks you to type RESET |

---

## 5. Costs

- **Reading:** roughly 2–4 cents per invoice with Opus. Fable only reads unsure invoices, and costs about 2.5× Opus for those.
- **Profiles:** under 1 cent extra per file (cached).
- **Learning:** a few cents a day.
- **A Test run:** about $5–15 (re-reads ~250 invoices).
- The **Log** tab shows the exact cost of every call.

---

## 6. Installing an update

1. Apps Script: paste the new `InvoiceReader.gs` over the old one. `ExampleSeed.gs` only needs adding once.
2. Reload the sheet.
3. **Invoices → Setup → Set up sheet and folders**, then **Invoices → Testing and one-off → Load the Sofia codes Claude has**.
4. First time only: **Load starting examples**, and **Remove test codes** if the TEST- codes are still there.
