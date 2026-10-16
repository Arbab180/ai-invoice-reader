# AI Invoice Reader

An automation that reads scanned supplier invoices with AI and turns them into ready-to-import accounting entries, built for a group of restaurants that receives hundreds of paper invoices a month.

It runs entirely inside Google Workspace (Drive + Sheets + Apps Script) and uses the Claude API to read the scans.

> **Note:** this is a public copy for my portfolio. Supplier names, restaurant names, account codes and folder IDs have been replaced with sample data. The logic is the real system.

## The problem

Every invoice used to be typed into the accounting system (Sofia) by hand: supplier, invoice number, date, restaurant, VAT split, expense account. The scans are messy: Greek and English, handwritten notes, stamps over numbers, several invoices in one PDF, the same head-office address on every branch's invoices.

## What it does

```
Scanner → Drive "Inbox" → AI first read → rule checks → AI second read (only if unsure)
       → supplier / restaurant / account codes → Journal → "Export for Sofia"
                                             ↘ Needs Check (a person fixes it) → learns from the fix
```

- **Reads every scan** with a strong model and returns structured JSON: supplier, invoice no., type (invoice / cash invoice / credit note), dates, restaurant, lines by category with net, VAT % and VAT.
- **Checks its own work.** Rules flag anything off: unknown supplier, totals that don't add up, VAT rates that don't exist, unclear restaurant. Only those fields get a *blind* second read by a different model. Agree = OK, disagree = Needs Check with both readings shown.
- **Works out which restaurant** from receiving stamps, handwritten notes and delivery lines, not just the billing address.
- **Fills accounting codes** from lookup tabs: supplier account, restaurant (T1) code, expense account, separate stock accounts for the central store.
- **One-click export** in the accounting system's import layout.
- **Learns from corrections.** Hand edits are logged; repeated fixes become lessons and verified examples that go into future prompts, so the same mistake isn't made twice.
- **Keeps people in the loop only where needed.** Staff only look at the Needs Check tab.
- Runs on a timer, logs cost per run, handles API rate limits, and has a test mode that re-reads past invoices to measure accuracy.

## Tech

- Google Apps Script (JavaScript), Google Sheets as the UI and database, Google Drive for files
- Claude API (vision + structured output), two-model cross-check, prompt caching
- Time-driven triggers, `onEdit` change tracking, Script Properties for the API key

## Files

| File | What it is |
|---|---|
| [`src/InvoiceReader.gs`](src/InvoiceReader.gs) | The main script (menu, reading, checks, coding, export, learning loop) |
| [`src/SplitInvoices.gs`](src/SplitInvoices.gs) | Early helper that split multi-invoice PDFs and sorted them by restaurant |
| [`src/ExampleSeed.gs`](src/ExampleSeed.gs) | Starting examples (sample data here) |
| [`docs/how-it-works.md`](docs/how-it-works.md) | The guide for the people who use the sheet |

## How I built it

I designed the workflow, the accounting rules and the checks with the business's accounts team, and built and tested it iteratively on real invoice batches, using Claude as a coding partner. Accuracy was measured against a hand-checked answer key of about 200 invoices, and each round of fixes came from what the test runs got wrong.

## Try it yourself

1. Create a Google Sheet, open **Extensions → Apps Script**, and paste in the three files from `src/`.
2. Set `PARENT_FOLDER_ID` at the top of `InvoiceReader.gs` to a Drive folder of yours.
3. Reload the sheet, then in the **Invoices** menu run **Setup → Set up sheet and folders**, then **Setup → Set Claude API key** (stored in Script Properties, never in the code).
4. Drop a scanned invoice PDF into the Inbox folder and run **Invoices → Read new invoices now**.
