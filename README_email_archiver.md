# Email Archiver (Outlook → Filesystem) — v2.6

A Windows-only Python utility that routes and archives Outlook messages from one or more mailboxes to a filesystem structure. Rules are defined in an Excel "route map". Messages are saved as Unicode `.msg` files or, optionally, only their attachments are extracted. The tool also categorizes processed Outlook items and produces a run summary.

Script: `email_archiver_v26.py`


## Key Features

- Subject-only `.msg` names: uses sanitized `Subject.msg` (no timestamps, no EntryID).
- Attachment titles preserved: attachment-only routes save files with original names.
- Uniqueness: on collision, appends `_01`, `_02`, … before the extension.
- Monthly folders when known: `RootPath\\YYYY\\MM-MonthName` (e.g., `2025\\06-June`).
- Period from content: month/year derived from subject and attachment titles only; fallback is current year (no month).
- Year scaffold: when saving to a month folder, the script also creates remaining month folders for that year; for the current year this is capped at the current month (e.g., in September 2025, creating `09-September` also creates `10-October` to `09-September` only — capped at current month).
- Outlook categories additive: new category is appended, existing categories are preserved.
- Multi-mailbox: processes multiple mailboxes by display name or SMTP.
- Summary workbook + unknown sender CSV for operational follow-up.


## Quick Start

1) Install prerequisites

- Windows with Microsoft Outlook installed and configured for the target mailboxes.
- Python 3.10+ (uses modern typing like `Tuple[...] | None`).
- Packages: `pywin32`, `pandas`, `python-dateutil`, `openpyxl`.

PowerShell example:

```
py -3.11 -m pip install --upgrade pip
py -3.11 -m pip install pywin32 pandas python-dateutil openpyxl
```

2) Configure the script

Open `email_archiver_v26.py` and review the user-config block:

- `MAP_PATH`: UNC/local path to the route map Excel.
- `UNKNOWN_CSV`: CSV path for unmapped senders.
- `DEFAULT_SAVE_PATH`: fallback root for unrecognized senders.
- `SUMMARY_PATH`: Excel path where the run log is written.
- `MAILBOXES`: list of mailbox display names or SMTP addresses to process.
- `SAVE_TYPE`: Outlook `OlSaveAsType` (default `3` → Unicode `.msg`).
- `MAX_PATH_LENGTH`: guard against long Windows paths (default `200`).
- `CAT_SAVED`, `CAT_DEFAULT`, `CAT_SKIPPED`: Outlook category names to apply.

3) Create/verify the route map workbook (see “Route Map Schema” below).

4) Run it

PowerShell examples:

```
# Archive yesterday for all configured mailboxes
py -3.11 .\email_archiver_v26.py --yesterday

# Archive a specific date
py -3.11 .\email_archiver_v26.py --date 2025-08-20

# Archive an inclusive range
py -3.11 .\email_archiver_v26.py --range 2025-08-01 2025-08-15
```


## Route Map Schema (Excel at `MAP_PATH`)

The route map drives where emails go and whether to save the full message or only attachments. The script reads it via `pandas.read_excel()` (all values treated as strings; blanks become empty).

Required columns:

| Column         | Type   | Allowed values | Purpose                                                                                   |
|----------------|--------|----------------|-------------------------------------------------------------------------------------------|
| `SenderEmail`  | text   | any email      | Exact sender email for specific mapping, or a representative domain sender (see below).   |
| `GenericSender`| text   | `yes`/`no`     | `no` = exact sender mapping; `yes` = domain-level/generic mapping (use `SubjectKey`).     |
| `SubjectKey`   | text   | comma list     | Comma-separated keywords matched against the email subject (lowercased).                  |
| `RootPath`     | path   | UNC/local path | Base folder where the message/attachments should be saved.                                |
| `Attachment`   | text   | `yes`/`no`     | `yes` = save attachments only, skip `.msg`; `no` = save the `.msg` file.                 |

How resolution works:

- Exact match first: a lowercased `SenderEmail` that equals the actual sender wins, returning `(RootPath, Attachment)`.
- Generic (domain) match next: rows with `GenericSender = 'yes'` are grouped by domain; a row applies when any `SubjectKey` token is contained in the lowercased subject. Multiple generic rows can exist per domain, each with their own `RootPath` and `Attachment` choice.
- Domain fallback: if a domain has at least one non-generic (`GenericSender = 'no'`) row, that domain’s first specific row becomes the default for unmapped specific senders from that domain.
- Otherwise subject-key fallback: for rows in the domain, if any `SubjectKey` matches, use their `RootPath`.
- No match: route resolution returns `None` and the message is treated as unknown (see below).

Notes:

- The code maintains an in-memory append (`add_row`) for newly encountered specific senders inferred from a domain rule. Persisting that back to Excel, if desired, should be done in the archive engine (see “Engine Flow”).
- Values are case-insensitive and trimmed; blanks are treated as empty strings.

Example rows:

| SenderEmail              | GenericSender | SubjectKey                 | RootPath                         | Attachment |
|--------------------------|---------------|----------------------------|----------------------------------|------------|
| invoices@vendor.com      | no            |                            | \\share\AP\VendorA              | yes        |
| @vendor.com              | yes           | statement,balance          | \\share\AP\Statements           | no         |
| alerts@fundadmin.co      | no            |                            | \\share\Ops\NAV\FundAdmin       | no         |


## Other Files Produced/Used

- `UNKNOWN_CSV`: A CSV ledger of completely unmapped senders to help expand the route map. Typical fields include: `Mailbox, SenderEmail, Subject, ReceivedTime`. The exact columns may vary by your local engine; include at least `SenderEmail` and timestamp to be useful.
- `SUMMARY_PATH`: An Excel workbook containing a run summary. Typical columns: `Mailbox, SenderEmail, Subject, Saved, SavedPath(s), AttachCount, Category, Error, StartedAt, FinishedAt`. Adjust to your needs.


## Foldering and Filenames

- Target folder:
  - If month extracted: `RootPath\YYYY\MM-MonthName` (e.g., `06-June`). Also creates remaining month folders for the year to aid future filing.
  - If month not extracted: `RootPath\YYYY` only (no month) so an analyst can move it into the correct month later.
- `.msg` name: `Subject.msg` (sanitized; illegal characters replaced; no timestamps or EntryID).
- Attachment names (attachment-only routes): original attachment file names (sanitized), saved under the determined folder (year-only if no month).
- Uniqueness: `_unique_path()` appends `_01`, `_02`, … before the suffix when a name already exists.
- Path length guard: `shorten_filename()` ensures the full path length stays within `MAX_PATH_LENGTH` (default 200). If needed, the base name is truncated (minimum ~30 characters where possible).


## Date Detection (Period Resolver)

To determine the archive period (subject/attachment titles only), the extractor recognizes:

- Month name + 4-digit year: `September 2025`, `Sept-2025`, `Sep2025`.
- 4-digit year + month name: `2025 September`.
- Month name + 2-digit year: `Sep '25`, `September-25` (maps `25` → `2025`).
- Numeric with separators: `2025-06`, `2025/6`, `06/2025`, `6 2025`.
- Contiguous numeric: `202506` (YYYYMM), `062025` (MMYYYY).
- Day-inclusive forms (month inferred from the date): `31/08/2025`, `2025-08-31`, `31-Aug-2025`, `31Aug25`, `31August2025`, `31.08.25`, `310825`, `31082025`.
- Quarter forms mapped to month-end: `Q3 2025`, `2025 Q3`, `3Q 2025` → September 2025.

Selection rules: prefer explicit month names and 4-digit years; if multiple candidates exist, pick the latest by year+month; slight preference to dates found in attachment names over subject.

Future-date control: if a parsed period is in the future (beyond the current year/month), it’s treated as invalid and placement falls back to the current year (no month). This keeps items accessible for analysts while preventing misfiles.

Two-digit year policy: two-digit years map to `20YY` (e.g., `’25` → `2025`). The future-date control prevents filing into future years.


## Outlook Categories

If you pre-create categories in Outlook with names matching the constants, the script can label processed items. Category assignment is additive and does not overwrite existing categories.

- `Saved`: Message saved (or attachments extracted) successfully.
- `Not Saved`: Message routed to default or otherwise not archived per rules.
- `Skipped`: Attachment-only rule but message had no attachments; or message intentionally skipped.


## Command Line Interface

Flags (mutually exclusive):

- `--yesterday`: Archives messages from the previous calendar day.
- `--date YYYY-MM-DD`: Archives messages from that specific date.
- `--range FROM TO`: Archives an inclusive date range; the two dates are sorted so order doesn’t matter.

Review and dry-run:

- `--dry-run`: Simulates the run and prints the planned destinations and filenames without writing or categorizing.
- `--interactive`: If no date flag is provided, prompts to choose Yesterday/Date/Range and shows a dry-run preview.
- `--yes`: After the preview, proceed to the live run without asking for confirmation.

Examples:

```
# Preview yesterday’s actions (no writes)
py -3.11 .\email_archiver_v26.py --yesterday --dry-run

# Interactive: pick date(s), preview, then confirm to run
py -3.11 .\email_archiver_v26.py --interactive

# One-shot: preview range, then run without prompt
py -3.11 .\email_archiver_v26.py --range 2025-08-01 2025-08-15 --yes
```


## Engine Flow (Architecture)

High-level logic intended by v2.6 (with `archive_window(start, end)`):

1) Load route map (`pandas.read_excel(MAP_PATH)`), normalize strings/blanks.
2) Build two structures:
   - `exact`: `{ sender_email_lower: (root_path, attach_only_bool) }` for `GenericSender = 'no'` rows.
   - `generic`: `{ domain: [ (subject_keys_list, root_path, attach_only_bool), ... ] }` for `GenericSender = 'yes'` rows.
3) Connect to Outlook via `win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")` and bind each mailbox in `MAILBOXES`.
4) For each mailbox and message in the date window:
   - Resolve route: `resolve_route(sender, subject, df, exact, generic)`.
   - If route is `None`: log to `UNKNOWN_CSV`, optionally save to `DEFAULT_SAVE_PATH`, and set category `Not Saved`.
   - If route found and `Attachment = yes`:
     - If the message has attachments: call `save_attachments(msg, root)` (saves under `YYYY/MM-MonthName` if known, else `YYYY`) and set category `Saved`.
     - If no attachments: set category `Skipped`.
   - If route found and `Attachment = no`: call `save_message(msg, root)` and set category `Saved`.
   - Append to in-memory summary rows (mailbox, sender, subject, saved path(s), etc.).
5) Write the summary to `SUMMARY_PATH` (Excel) and ensure `UNKNOWN_CSV` is appended/created with headers.
6) Optionally persist route map enhancements produced by `add_row()` back to `MAP_PATH`.

Note: In this repository, `archive_window()` is intentionally abbreviated (placeholder comment). Port the full engine body from v2.5 if you have it, keeping the new v2.6 save/name behavior.


## Changelog Highlights (v2.6)

- Removed Outlook EntryID usage in filenames.
- Collision policy: `_01`, `_02`, … appended before the extension.
- `.msg` filename now uses Subject-only (no timestamps).
- Attachment-only routes save attachments with original titles under `YYYY/MM-MonthName`.
- Month folder naming `MM-MonthName` enforced everywhere.
- Period derives from content (subject/attachments); fallback is the current month.


## Requirements & Environment

- OS: Windows (uses Outlook COM automation via `pywin32`).
- Outlook: Must be installed, openable, and profiles configured to access `MAILBOXES`.
- Python: 3.10+ recommended.
- Python packages: `pywin32`, `pandas`, `python-dateutil`, `openpyxl`.


## Operational Notes & Limits

- Path length: Capped by `MAX_PATH_LENGTH` (default 200). Long subjects/paths are truncated safely.
- Illegal characters: Windows-forbidden path characters are replaced with `-`.
- No timestamps in filenames: names are based on subject or original attachment title.
- Idempotency: Safe to re-run; exact duplicates become `_01`, `_02`, … rather than overwriting.
- Network shares: Ensure the account running the script has read/write permissions.
- Categories: Create `Saved`, `Not Saved`, `Skipped` in Outlook ahead of time if you want labeling (assignments are appended, not overwritten).


## Troubleshooting

- ImportError for `win32com.client`: `py -3.11 -m pip install pywin32` and run `py -3.11 -m pywin32_postinstall install` if needed.
- `PermissionError` or cannot write: Verify `RootPath`, `SUMMARY_PATH`, and `UNKNOWN_CSV` share permissions.
- Excel read/write errors: Install `openpyxl` and close any Excel file that might be locking the workbook.
- No messages processed: Verify `MAILBOXES` names, date flags, and Outlook profile.


## Security

- The script writes files to paths determined by the route map. Treat `MAP_PATH` as trusted configuration.
- Attachments are saved as-is; consider scanning outputs with AV/EDR where appropriate.


## Extending

- Add more matchers (e.g., regex on subject) before the fallback to `ReceivedTime`.
- Persist `add_row()` updates back into the route map to “learn” new exact senders from domains.
- Expand the summary workbook with more operational metrics.


---

If you need me to also produce a `requirements.txt` or a starter `RouteMap.xlsx` template, say the word and I’ll add them.
