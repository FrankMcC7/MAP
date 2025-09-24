#!/usr/bin/env python3
# Email Archiver – v2.6  (Aug 2025)
#
# Changes in v2.6
#   • Removed any use of Outlook EntryID in filenames (kept from v2.5 policy).
#   • When a filename collision occurs, append _01, _02, ... (two-digit, zero-padded)
#     immediately before the extension instead of _(2), _(3).
#
# From v2.5
#   • Saved .msg filenames use a clean pattern:
#       <YYYYMMDD_HHMMSS>_<Subject>.msg
#   • Month folders use “MM-MonthName” (e.g. 06-June)
#   • New Excel column “Attachment” (Yes / No):
#       – Yes  ➜ save only the attachments into RootPath (skip the e-mail itself)
#                – if no attachment exists the message is *skipped*
#       – No   ➜ save the whole message under  <RootPath>\YYYY\MM-MonthName
#   • Optional Outlook categories: Saved, Not Saved, Skipped.
# -----------------------------------------------------------

from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import sys
from pathlib import Path
from typing import Any, List, Tuple

import pandas as pd
import win32com.client as win32
from dateutil import parser as nlp

# ─────────────────────────── USER CONFIG ────────────────────────────
MAP_PATH          = r"\\share\EmailRoutes.xlsx"        # routing / rule book
UNKNOWN_CSV       = r"\\share\unknown_senders.csv"     # completely unmapped
DEFAULT_SAVE_PATH = r"\\share\DefaultArchive"          # fallback root
SUMMARY_PATH      = r"\\share\ArchiveSummary.xlsx"     # run log
MAILBOXES         = ["Funds Ops", "NAV Alerts"]        # display names / SMTP
SAVE_TYPE         = 3                                  # olMsg (Unicode)
MAX_PATH_LENGTH   = 200                                # path length guard
# Outlook categories you’ve pre-created:
CAT_SAVED         = "Saved"
CAT_DEFAULT       = "Not Saved"
CAT_SKIPPED       = "Skipped"

# ──────────────────────────── REGEX HELPERS ─────────────────────────
ISO_RGX = re.compile(r"\b(20\d{2})[-_.]?([01]?\d)\b")
MON_RGX = re.compile(
    r"(?i)\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[-_/\. ]*(20\d{2})"
)
ILLEGAL_FS = re.compile(r'[\\/:\*\?"<>|]')


def _to_bool(s: str) -> bool:
    return str(s).strip().lower() in {"y", "yes", "true", "1"}

# ─────────────────────────── PATH UTILITIES ─────────────────────────
def shorten_filename(target_dir: str, base: str, ext: str) -> str:
    """Truncate so full path ≤ MAX_PATH_LENGTH."""
    full = os.path.join(target_dir, base + ext)
    if len(full) <= MAX_PATH_LENGTH:
        return base + ext
    allow = MAX_PATH_LENGTH - len(target_dir) - len(os.sep) - len(ext)
    return (base[: max(allow, 30)] + ext) if allow > 0 else ("trunc" + ext)


def month_folder(month: int) -> str:
    return f"{month:02d}-{dt.date(1900, month, 1).strftime('%B')}"


def _ensure_remaining_months(year_dir: Path, archive_year: int, from_month: int) -> None:
    """Create month folders from (from_month+1) forward, capped by current month for current year.

    Does not overwrite files; only ensures directories exist.
    """
    today = dt.date.today()
    cap_month = 12 if archive_year < today.year else today.month
    for m in range(max(1, from_month + 1), cap_month + 1):
        (year_dir / month_folder(m)).mkdir(parents=True, exist_ok=True)


def _unique_path(base_path: Path) -> Path:
    """
    Ensure the path is unique by appending _01, _02, ... before the suffix if needed.
    base_path already includes the suffix (.msg or otherwise).
    """
    if not base_path.exists():
        return base_path
    stem, suffix = base_path.stem, base_path.suffix
    n = 1
    while True:
        candidate = base_path.with_name(f"{stem}_{n:02d}{suffix}")
        if not candidate.exists():
            return candidate
        n += 1


# ─────────────────────────── SAVE ROUTINES ──────────────────────────
def save_message(msg, root: str, forced_day: dt.date | None = None):
    """Save the whole e-mail as .msg using Subject-only naming with uniqueness."""
    root_path = Path(root)
    is_default_root = os.path.normcase(os.path.normpath(root)) == os.path.normcase(os.path.normpath(DEFAULT_SAVE_PATH))
    if forced_day is not None and is_default_root:
        year_dir = root_path / str(forced_day.year)
        year_dir.mkdir(parents=True, exist_ok=True)
        month_dir = year_dir / month_folder(forced_day.month)
        month_dir.mkdir(parents=True, exist_ok=True)
        folder = month_dir / forced_day.isoformat()
        folder.mkdir(parents=True, exist_ok=True)
    else:
        yr, mo = detect_period(msg)
        year_dir = root_path / str(yr)
        year_dir.mkdir(parents=True, exist_ok=True)
        if mo is None:
            folder = year_dir
        else:
            folder = year_dir / month_folder(mo)
            folder.mkdir(parents=True, exist_ok=True)
            _ensure_remaining_months(year_dir, yr, mo)

    subj = ILLEGAL_FS.sub("-", (msg.Subject or "No Subject")).strip()[:150]
    base = subj if subj else "No Subject"
    fname = shorten_filename(str(folder), base, ".msg")
    path = _unique_path(folder / fname)

    msg.SaveAs(str(path), SAVE_TYPE)

def save_attachments(msg, root: str, forced_day: dt.date | None = None):
    """Save attachments only using original attachment titles.

    Foldering:
    - If month extracted: Root/Year/MM-Month (and scaffold remaining months).
    - If not: Root/Year only, so analysts can place into correct month later.
    """
    root_path = Path(root)
    is_default_root = os.path.normcase(os.path.normpath(root)) == os.path.normcase(os.path.normpath(DEFAULT_SAVE_PATH))
    if forced_day is not None and is_default_root:
        year_dir = root_path / str(forced_day.year)
        year_dir.mkdir(parents=True, exist_ok=True)
        month_dir = year_dir / month_folder(forced_day.month)
        month_dir.mkdir(parents=True, exist_ok=True)
        folder = month_dir / forced_day.isoformat()
        folder.mkdir(parents=True, exist_ok=True)
    else:
        yr, mo = detect_period(msg)
        year_dir = root_path / str(yr)
        year_dir.mkdir(parents=True, exist_ok=True)
        if mo is None:
            folder = year_dir
        else:
            folder = year_dir / month_folder(mo)
            folder.mkdir(parents=True, exist_ok=True)
            _ensure_remaining_months(year_dir, yr, mo)

    for att in msg.Attachments:
        fname_raw = ILLEGAL_FS.sub("-", att.FileName)
        name = shorten_filename(str(folder), fname_raw, "")
        path = _unique_path(folder / name)
        att.SaveAsFile(str(path))


def _plan_paths_for_message(msg, root: str, attachment_only: bool, forced_day: dt.date | None = None) -> Tuple[Path, List[Path]]:
    root_path = Path(root)
    is_default_root = os.path.normcase(os.path.normpath(root)) == os.path.normcase(os.path.normpath(DEFAULT_SAVE_PATH))
    if forced_day is not None and is_default_root:
        year_dir = root_path / str(forced_day.year)
        month_dir = year_dir / month_folder(forced_day.month)
        folder = month_dir / forced_day.isoformat()
    else:
        yr, mo = detect_period(msg)
        year_dir = root_path / str(yr)
        folder = year_dir if mo is None else (year_dir / month_folder(mo))

    out_paths: List[Path] = []
    if attachment_only:
        try:
            atts = list(msg.Attachments)
        except Exception:
            atts = []
        for att in atts:
            fname_raw = ILLEGAL_FS.sub("-", att.FileName)
            name = shorten_filename(str(folder), fname_raw, "")
            out_paths.append(_unique_path(folder / name))
    else:
        base = _sanitize_subject(msg)
        name = shorten_filename(str(folder), base, ".msg")
        out_paths.append(_unique_path(folder / name))
    return folder, out_paths




def detect_period(item) -> Tuple[int, int | None]:
    """Extract (year, month) from subject/attachment titles only; fallback to (current_year, None).

    Recognizes MonthName-Year, YYYY-MM/MM-YYYY, contiguous YYYYMM/MMYYYY,
    and day-inclusive forms like 31/08/2025, 2025-08-31, 31Aug25, 31August2025, 31.08.25, 310825.
    """

    def norm_two_digit_year(yy: int) -> int:
        # Policy: always interpret YY as 20YY (e.g., '25' -> 2025)
        return 2000 + yy

    def mon_from_word(s: str) -> int | None:
        try:
            return dt.datetime.strptime((s[:3]).title(), "%b").month
        except Exception:
            return None

    def add_candidate(cands: list[tuple[int, int, int]], y: int, m: int, w: int):
        if 1 <= m <= 12 and 1900 <= y <= 2100:
            cands.append((y, m, w))

    subject = item.Subject or ""
    try:
        att_names = [att.FileName for att in item.Attachments]
    except Exception:
        att_names = []

    candidates: list[tuple[int, int, int]] = []
    texts = [(subject, 0)] + [(name, 5) for name in att_names]

    # Month name alternatives (include 'sept')
    MON = (
        "jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|"
        "aug|august|sep|sept|september|oct|october|nov|november|dec|december"
    )
    # Use non-letter boundaries so underscores count as separators too (e.g., Paribas_June 2025)
    pat_mon_yyyy = re.compile(
        rf"(?<![A-Za-z])({MON})(?![A-Za-z])[-_/\.\s]*'?((?:19|20)\d{{2}})(?!\d)", re.I
    )
    pat_yyyy_mon = re.compile(
        rf"\b((?:19|20)\d{{2}})\b[-_/\.\s]*'?({MON})(?![A-Za-z])", re.I
    )
    pat_mon_yy = re.compile(
        rf"(?<![A-Za-z])({MON})(?![A-Za-z])[-_/\.\s]*'?(\d{{2}})(?!\d)", re.I
    )
    pat_dd_mon_yyyy = re.compile(
        rf"\b([0-3]?\d)[-_/\.\s]*({MON})[-_/\.\s]*'?((?:19|20)?\d{{2}})\b", re.I
    )
    pat_mon_dd_yyyy = re.compile(
        rf"\b({MON})[-_/\.\s]*([0-3]?\d)[-_/\.\s]*'?((?:19|20)?\d{{2}})\b", re.I
    )
    pat_ddmonyyyy_contig = re.compile(
        rf"(?<!\d)([0-3]?\d)({MON})((?:19|20)?\d{{2}})(?!\d)", re.I
    )

    # Quarter variants (mapped to month: Q1->03, Q2->06, Q3->09, Q4->12)
    q_to_month = {1: 3, 2: 6, 3: 9, 4: 12}
    pat_qn_yyyy = re.compile(r"\bQ([1-4])[-_/\.\s]*((?:19|20)?\d{2})\b", re.I)
    pat_yyyy_qn = re.compile(r"\b((?:19|20)?\d{2})[-_/\.\s]*Q([1-4])\b", re.I)
    pat_nq_yyyy = re.compile(r"\b([1-4])Q[-_/\.\s]*((?:19|20)?\d{2})\b", re.I)

    # Numeric variants
    pat_yyyy_mm = re.compile(r"\b((?:19|20)\d{2})[-_/\.\s]([01]?\d)\b")
    pat_mm_yyyy = re.compile(r"\b([01]?\d)[-_/\.\s]((?:19|20)\d{2})\b")
    pat_yyyymm = re.compile(r"(?<!\d)((?:19|20)\d{2})(1[0-2]|0[1-9])(?!\d)")
    pat_mmyyyy = re.compile(r"(?<!\d)(1[0-2]|0[1-9])((?:19|20)\d{2})(?!\d)")
    pat_dd_mm_yyyy = re.compile(r"\b([0-3]?\d)[-\./]([01]?\d)[-\./]((?:19|20)\d{2})\b")
    pat_yyyy_mm_dd = re.compile(r"\b((?:19|20)\d{2})[-\./]([01]?\d)[-\./]([0-3]?\d)\b")
    pat_ddmmyyyy = re.compile(r"(?<!\d)([0-3]\d)([01]\d)((?:19|20)\d{2})(?!\d)")
    pat_ddmmyy = re.compile(r"(?<!\d)([0-3]\d)([01]\d)(\d{2})(?!\d)")

    for text, bonus in texts:
        if not text:
            continue
        text_clean = text.replace("’", "'")

        for match in pat_mon_yyyy.finditer(text_clean):
            mon = mon_from_word(match.group(1))
            yr = int(match.group(2))
            add_candidate(candidates, yr, mon or 0, 100 + bonus)

        for match in pat_yyyy_mon.finditer(text_clean):
            yr = int(match.group(1))
            mon = mon_from_word(match.group(2))
            add_candidate(candidates, yr, mon or 0, 95 + bonus)

        for match in pat_mon_yy.finditer(text_clean):
            mon = mon_from_word(match.group(1))
            yr = norm_two_digit_year(int(match.group(2)))
            add_candidate(candidates, yr, mon or 0, 90 + bonus)

        for match in pat_dd_mon_yyyy.finditer(text_clean):
            day = int(match.group(1))
            mon = mon_from_word(match.group(2))
            yr_raw = match.group(3)
            yr = int(yr_raw) if len(yr_raw) == 4 else norm_two_digit_year(int(yr_raw))
            add_candidate(candidates, yr, mon or 0, 80 + bonus + (2 if day <= 12 else 0))

        for match in pat_mon_dd_yyyy.finditer(text_clean):
            mon = mon_from_word(match.group(1))
            day = int(match.group(2))
            yr_raw = match.group(3)
            yr = int(yr_raw) if len(yr_raw) == 4 else norm_two_digit_year(int(yr_raw))
            add_candidate(candidates, yr, mon or 0, 78 + bonus + (2 if day <= 12 else 0))

        for match in pat_ddmonyyyy_contig.finditer(text_clean):
            day = int(match.group(1))
            mon = mon_from_word(match.group(2))
            yr_raw = match.group(3)
            yr = int(yr_raw) if len(yr_raw) == 4 else norm_two_digit_year(int(yr_raw))
            add_candidate(candidates, yr, mon or 0, 75 + bonus + (2 if day <= 12 else 0))

        for match in pat_qn_yyyy.finditer(text_clean):
            q = int(match.group(1))
            yr_raw = match.group(2)
            yr = int(yr_raw) if len(yr_raw) == 4 else norm_two_digit_year(int(yr_raw))
            add_candidate(candidates, yr, q_to_month.get(q, 0), 70 + bonus)

        for match in pat_yyyy_qn.finditer(text_clean):
            yr_raw = match.group(1)
            yr = int(yr_raw) if len(yr_raw) == 4 else norm_two_digit_year(int(yr_raw))
            q = int(match.group(2))
            add_candidate(candidates, yr, q_to_month.get(q, 0), 68 + bonus)

        for match in pat_nq_yyyy.finditer(text_clean):
            q = int(match.group(1))
            yr_raw = match.group(2)
            yr = int(yr_raw) if len(yr_raw) == 4 else norm_two_digit_year(int(yr_raw))
            add_candidate(candidates, yr, q_to_month.get(q, 0), 66 + bonus)

        for match in pat_yyyy_mm.finditer(text_clean):
            yr = int(match.group(1))
            mon = int(match.group(2))
            add_candidate(candidates, yr, mon, 60 + bonus)

        for match in pat_mm_yyyy.finditer(text_clean):
            mon = int(match.group(1))
            yr = int(match.group(2))
            add_candidate(candidates, yr, mon, 58 + bonus)

        for match in pat_yyyymm.finditer(text_clean):
            yr = int(match.group(1))
            mon = int(match.group(2))
            add_candidate(candidates, yr, mon, 56 + bonus)

        for match in pat_mmyyyy.finditer(text_clean):
            mon = int(match.group(1))
            yr = int(match.group(2))
            add_candidate(candidates, yr, mon, 54 + bonus)

        for match in pat_dd_mm_yyyy.finditer(text_clean):
            day = int(match.group(1))
            mon = int(match.group(2))
            yr = int(match.group(3))
            add_candidate(candidates, yr, mon, 50 + bonus + (2 if day <= 12 else 0))

        for match in pat_yyyy_mm_dd.finditer(text_clean):
            yr = int(match.group(1))
            mon = int(match.group(2))
            day = int(match.group(3))
            add_candidate(candidates, yr, mon, 48 + bonus + (2 if day <= 12 else 0))

        for match in pat_ddmmyyyy.finditer(text_clean):
            day = int(match.group(1))
            mon = int(match.group(2))
            yr = int(match.group(3))
            add_candidate(candidates, yr, mon, 46 + bonus + (2 if day <= 12 else 0))

        for match in pat_ddmmyy.finditer(text_clean):
            day = int(match.group(1))
            mon = int(match.group(2))
            yr = norm_two_digit_year(int(match.group(3)))
            add_candidate(candidates, yr, mon, 44 + bonus + (2 if day <= 12 else 0))

        for match in ISO_RGX.finditer(text_clean):
            yr = int(match.group(1))
            mon = int(match.group(2))
            add_candidate(candidates, yr, mon, 40 + bonus)

        for mon_word in MON.split('|'):
            mon_idx = mon_from_word(mon_word)
            if mon_idx and mon_word.lower() in text_clean.lower():
                today_hint = dt.date.today()
                add_candidate(candidates, today_hint.year, mon_idx, 10 + bonus)

    if candidates:
        candidates.sort(key=lambda t: (t[2], t[0], t[1]))
        today = dt.date.today()
        y, m, _ = candidates[-1]
        if y > today.year:
            return today.year, None
        if y == today.year and m is not None and m > today.month:
            return today.year, None
        return y, m

    today = dt.date.today()
    return today.year, None


def set_category(itm: Any, cat: str):
    """Append the category without overwriting existing categories."""
    try:
        existing = itm.Categories or ""
    except Exception:
        existing = ""
    cats = [c.strip() for c in existing.split(';') if c.strip()]
    if cat and cat not in cats:
        cats.append(cat)
        itm.Categories = '; '.join(cats)
        try:
            itm.Save()
        except Exception:
            pass


def add_row(df: pd.DataFrame, sender: str, root: str):
    df.loc[len(df)] = [sender, 'no', '', root, 'no']


def resolve_route(sender: str, subject: str, df, exact, generic) -> Tuple[str, bool] | None:
    key = _normalize_sender_key(sender)
    if key in exact:
        return exact[key]

    subject_lc = (subject or '').strip().lower()
    domain = _domain_from_sender(sender) or key
    if not domain:
        return None

    if domain in generic:
        for keys, root, attach in generic[domain]:
            if any(k and k in subject_lc for k in keys):
                return root, attach

    if 'SenderEmail' not in df.columns:
        return None

    sender_series = df['SenderEmail'].astype(str).str.strip().str.lower()
    cand = df[sender_series.str.endswith(domain)]
    if cand.empty:
        return None

    if 'GenericSender' in cand.columns:
        non_gen = cand[~cand['GenericSender'].apply(_to_bool)]
    else:
        non_gen = cand

    if not non_gen.empty:
        root = str(non_gen.iloc[0].get('RootPath', '')).strip()
        if root:
            add_row(df, sender, root)
            return root, False
        return None

    if 'SubjectKey' in cand.columns:
        for _, r in cand.iterrows():
            keys = [k.strip().lower() for k in str(r.get('SubjectKey', '')).split(',') if k.strip()]
            if any(subject_lc == key for key in keys):
                root = str(r.get('RootPath', '')).strip()
                if root:
                    add_row(df, sender, root)
                    return root, False
    return None


def _normalize_sender_key(sender: str) -> str:
    return (sender or '').strip().lower()


def _domain_from_sender(sender: str) -> str | None:
    key = _normalize_sender_key(sender)
    if '@' not in key:
        return None
    domain = key.split('@')[-1]
    return domain or None


def _match_template_row(df: pd.DataFrame, sender: str, subject: str):
    if 'SenderEmail' not in df.columns:
        return None
    domain = _domain_from_sender(sender)
    if not domain:
        return None

    sender_col = df['SenderEmail'].astype(str).str.strip().str.lower()
    domain_mask = sender_col.str.endswith(domain)
    subset = df.loc[domain_mask]
    if subset.empty:
        return None

    if 'GenericSender' in subset.columns:
        non_generic_subset = subset[~subset['GenericSender'].apply(_to_bool)]
    else:
        non_generic_subset = subset

    if not non_generic_subset.empty:
        return non_generic_subset.iloc[0]

    if 'GenericSender' not in subset.columns:
        return None

    generic_subset = subset[subset['GenericSender'].apply(_to_bool)]
    if generic_subset.empty:
        return None

    subject_lc = (subject or '').strip().lower()
    if not subject_lc:
        return None

    for _, row in generic_subset.iterrows():
        keys = [k.strip().lower() for k in str(row.get('SubjectKey', '')).split(',') if k.strip()]
        if any(subject_lc == key for key in keys):
            return row
    return None


def _preprocess_sender_map(ns, df: pd.DataFrame, start: dt.date, end: dt.date):
    added_rows: list[dict[str, Any]] = []
    if 'SenderEmail' not in df.columns:
        return df, added_rows

    existing: set[str] = set()
    for val in df['SenderEmail'].tolist():
        key = _normalize_sender_key(str(val))
        if key and key != 'nan':
            existing.add(key)

    for mailbox in MAILBOXES:
        try:
            items = _iter_messages_in_window(ns, mailbox, start, end)
        except Exception:
            continue
        for msg in items:
            try:
                sender = getattr(msg, 'SenderEmailAddress', '') or ''
                subject = getattr(msg, 'Subject', '') or ''
            except Exception:
                continue

            sender_key = _normalize_sender_key(sender)
            if not sender_key or sender_key in existing:
                continue

            template = _match_template_row(df, sender, subject)
            if template is None:
                continue

            row_dict = {}
            for col in df.columns:
                if col == 'SenderEmail':
                    row_dict[col] = sender
                else:
                    row_dict[col] = template.get(col, '')
            df.loc[len(df)] = row_dict
            added_rows.append(row_dict.copy())
            existing.add(sender_key)

    return df, added_rows


def _update_summary_workbook(summary_rows: List[dict[str, Any]], start: dt.date, end: dt.date) -> None:
    if not summary_rows:
        return

    try:
        df = pd.DataFrame(summary_rows)
    except Exception:
        return

    if df.empty or 'Date' not in df.columns:
        return

    try:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    except Exception:
        return

    df['Date'] = df['Date'].fillna(pd.Timestamp(start))
    df['Date'] = df['Date'].dt.date

    saved_mask = df['Action'].isin({'message', 'attachments'})
    default_mask = saved_mask & (df['Root'] == DEFAULT_SAVE_PATH)

    summary = (
        df.assign(
            SavedCount=saved_mask.astype(int),
            DefaultPathCount=default_mask.astype(int),
        )
        .groupby('Date', dropna=False)[['SavedCount', 'DefaultPathCount']]
        .sum()
        .reset_index()
    )

    if summary.empty:
        return

    summary[['SavedCount', 'DefaultPathCount']] = summary[['SavedCount', 'DefaultPathCount']].astype(int)
    summary['RunStart'] = start.isoformat()
    summary['RunEnd'] = end.isoformat()
    summary['LoggedAt'] = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    summary = summary.sort_values('Date').reset_index(drop=True)

    dest = Path(SUMMARY_PATH)
    try:
        if dest.exists():
            previous = pd.read_excel(dest)
        else:
            previous = pd.DataFrame()
    except Exception:
        previous = pd.DataFrame()

    combined = pd.concat([previous, summary], ignore_index=True)

    try:
        combined.to_excel(dest, index=False)
    except Exception as exc:
        print(f"[WARN] Could not update summary workbook: {exc}")


def _build_route_index(df: pd.DataFrame):
    exact: dict[str, Tuple[str, bool]] = {}
    generic: dict[str, List[Tuple[List[str], str, bool]]] = {}
    for _, row in df.iterrows():
        sender = str(row.get('SenderEmail', '')).strip()
        root = str(row.get('RootPath', '')).strip()
        gflag = str(row.get('GenericSender', '')).strip().lower()
        attach = _to_bool(row.get('Attachment', 'no'))
        subj_keys = [k.strip().lower() for k in str(row.get('SubjectKey', '')).split(',') if k.strip()]
        if not sender or not root:
            continue
        if gflag == 'no':
            exact[sender.lower()] = (root, attach)
        else:
            domain = sender.split('@')[-1].lower()
            generic.setdefault(domain, []).append((subj_keys, root, attach))
    return exact, generic


def _format_outlook_datetime(d: dt.datetime) -> str:
    return d.strftime('%m/%d/%Y %I:%M %p')


def _iter_messages_in_window(ns, mailbox: str, start: dt.date, end: dt.date):
    store = ns.Folders.Item(mailbox)
    inbox = store.Folders.Item('Inbox')
    items = inbox.Items
    items.Sort('[ReceivedTime]', True)
    start_dt = dt.datetime.combine(start, dt.time.min)
    end_dt = dt.datetime.combine(end, dt.time.max)
    flt = (
        f"[ReceivedTime] >= '{_format_outlook_datetime(start_dt)}' AND "
        f"[ReceivedTime] <= '{_format_outlook_datetime(end_dt)}'"
    )
    restricted = items.Restrict(flt)
    for itm in restricted:
        yield itm


def _sanitize_subject(msg) -> str:
    return ILLEGAL_FS.sub('-', (msg.Subject or 'No Subject').strip())[:150] or 'No Subject'

def archive_window(start: dt.date, end: dt.date, dry_run: bool = False, interactive_confirm: bool = False):
    df = pd.read_excel(MAP_PATH, dtype=str).fillna("")

    outlook = win32.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    df, added_rows = _preprocess_sender_map(ns, df, start, end)

    if added_rows:
        if dry_run:
            print(f"[INFO] Pre-scan identified {len(added_rows)} new sender(s) (not persisted during dry run).")
        else:
            try:
                df.to_excel(MAP_PATH, index=False)
                print(f"[INFO] Pre-scan added {len(added_rows)} new sender(s) to routing map.")
            except Exception as e:
                print(f"[WARN] Could not write updated routes to '{MAP_PATH}': {e}")

    df = df.fillna("")
    exact, generic = _build_route_index(df)

    summary_rows = []
    unknown_rows = []

    for mailbox in MAILBOXES:
        try:
            items = _iter_messages_in_window(ns, mailbox, start, end)
        except Exception as e:
            print(f"[WARN] Cannot open mailbox '{mailbox}': {e}")
            continue

        for msg in items:
            try:
                sender = getattr(msg, "SenderEmailAddress", "") or ""
                subject = getattr(msg, "Subject", "") or ""
                received_raw = getattr(msg, "ReceivedTime", None)
            except Exception:
                continue

            received_date = None
            if received_raw:
                try:
                    received_date = pd.Timestamp(received_raw).date()
                except Exception:
                    received_date = None

            route = resolve_route(sender, subject, df, exact, generic)
            if route is None:
                root = DEFAULT_SAVE_PATH
                attach_only = False
                planned_cat = CAT_DEFAULT
                unknown_rows.append((mailbox, sender, subject))
            else:
                root, attach_only = route
                if attach_only:
                    try:
                        has_att = getattr(msg, "Attachments").Count > 0
                    except Exception:
                        has_att = False
                    planned_cat = CAT_SAVED if has_att else CAT_SKIPPED
                else:
                    planned_cat = CAT_SAVED

            is_default_root = os.path.normcase(os.path.normpath(root)) == os.path.normcase(os.path.normpath(DEFAULT_SAVE_PATH))
            if is_default_root:
                forced_day = received_date or start
            else:
                forced_day = None

            folder, file_paths = _plan_paths_for_message(msg, root, attach_only, forced_day)

            action = "attachments" if attach_only else "message"
            if attach_only:
                try:
                    has_att = getattr(msg, "Attachments").Count > 0
                except Exception:
                    has_att = False
                if not has_att:
                    action = "skip"

            folder_path = Path(folder)
            root_path = Path(root)
            try:
                rel_parts = folder_path.relative_to(root_path).parts
            except Exception:
                rel_parts = ()
            if rel_parts:
                year_value = rel_parts[0]
                month_value = rel_parts[1] if len(rel_parts) > 1 else ""
            else:
                year_value = folder_path.name
                month_value = folder_path.name if folder_path.name and folder_path.name[:2].isdigit() else ""

            summary_rows.append(
                {
                    "Date": received_date.isoformat() if received_date else start.isoformat(),
                    "Mailbox": mailbox,
                    "Sender": sender,
                    "Subject": subject,
                    "Year": year_value,
                    "MonthFolder": month_value,
                    "Root": root,
                    "Folder": str(folder),
                    "Files": "; ".join(str(p) for p in file_paths) if file_paths else "",
                    "Action": action,
                    "Category": planned_cat,
                }
            )

            if not dry_run:
                try:
                    if action == "attachments":
                        save_attachments(msg, root, forced_day)
                    elif action == "message":
                        save_message(msg, root, forced_day)
                    set_category(msg, planned_cat)
                except Exception as e:
                    print(f"[ERROR] Saving item failed: {e}")


    if not dry_run and unknown_rows:
        try:
            import csv
            file_exists = Path(UNKNOWN_CSV).exists()
            with open(UNKNOWN_CSV, "a", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                if not file_exists:
                    w.writerow(["Mailbox", "SenderEmail", "Subject"])
                for r in unknown_rows:
                    w.writerow(r)
        except Exception as e:
            print(f"[WARN] Could not write unknown senders: {e}")

    if not dry_run:
        _update_summary_workbook(summary_rows, start, end)

    if dry_run or interactive_confirm:
        print("Preview of planned actions:")
        for row in summary_rows[:200]:
            print(
                f"- [{row['Mailbox']}] {row['Subject']!r} -> {row['Action']} in {row['Folder']}\n  Files: {row['Files']}\n  Category: {row['Category']}"
            )

    return summary_rows



# ─────────────────────────── CLI ─────────────────────────
def parse_date(s: str) -> dt.date:
    return dt.datetime.strptime(s, "%Y-%m-%d").date()


def main():
    ap = argparse.ArgumentParser(description="Archive Outlook mailboxes")
    g = ap.add_mutually_exclusive_group(required=False)
    g.add_argument("--yesterday", action="store_true", help="Archive yesterday's messages")
    g.add_argument("--date", type=parse_date, metavar="YYYY-MM-DD", help="Archive a specific date")
    g.add_argument("--range", nargs=2, type=parse_date, metavar=("FROM", "TO"), help="Archive a date range (inclusive)")
    ap.add_argument("--dry-run", action="store_true", help="Simulate only; show planned folders/files without writing")
    ap.add_argument("--yes", action="store_true", help="Proceed without interactive confirmation in dry-run review")
    ap.add_argument("--interactive", action="store_true", help="Use interactive prompts if no date is provided")
    args = ap.parse_args()

    def interactive_pick():
        print("Select date mode: [Y]esterday, [D]ate, [R]ange")
        choice = input("> ").strip().lower()
        if choice.startswith("y"):
            d = dt.date.today() - dt.timedelta(days=1)
            return d, d
        if choice.startswith("d"):
            s = input("Enter date (YYYY-MM-DD): ").strip()
            d = parse_date(s)
            return d, d
        if choice.startswith("r"):
            s1 = input("From (YYYY-MM-DD): ").strip()
            s2 = input("To   (YYYY-MM-DD): ").strip()
            a, b = parse_date(s1), parse_date(s2)
            a, b = sorted([a, b])
            return a, b
        raise SystemExit("Invalid selection")

    if not (args.yesterday or args.date or args.range):
        if args.interactive:
            start, end = interactive_pick()
        else:
            ap.print_help()
            return
    else:
        if args.yesterday:
            start = dt.date.today() - dt.timedelta(days=1)
            end = start
        elif args.date:
            start = end = args.date
        else:
            start, end = sorted(args.range)

    # Always show preview for review
    _ = archive_window(start, end, dry_run=True, interactive_confirm=True)
    if args.dry_run:
        return

    if not args.yes:
        resp = input("Proceed with live run? [y/N]: ").strip().lower()
        if resp not in {"y", "yes"}:
            print("Aborted by user.")
            return

    archive_window(start, end, dry_run=False, interactive_confirm=False)


if __name__ == "__main__":
    main()

