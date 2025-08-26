#!/usr/bin/env python3
# Email Archiver – v2.8  (Aug 2025)
#
# Changes vs v2.7:
#   • Never falls back to ReceivedTime for foldering.
#   • Exhaustive (subject/attachment) parsing for Year+Month only.
#   • If NO (Year, Month) is found, save under <RootPath>\<CURRENT_YEAR>\ (no month).
#
# Still includes:
#   • No EntryID in filenames; collision-safe with _01, _02...
#   • Pre-enrichment of new non-generic senders into EmailRoutes.xlsx BEFORE saving.
#   • Append-only ArchiveSummary.xlsx logging.
#   • Attachments saved directly under RootPath when Attachment=Yes (per prior spec).

from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import win32com.client as win32

# ─────────────────────────── USER CONFIG ────────────────────────────
MAP_PATH          = r"\\share\EmailRoutes.xlsx"        # routing / rule book
UNKNOWN_CSV       = r"\\share\unknown_senders.csv"     # unmapped/errors log (CSV append)
DEFAULT_SAVE_PATH = r"\\share\DefaultArchive"          # fallback root for inferred routes
SUMMARY_PATH      = r"\\share\ArchiveSummary.xlsx"     # run log (append-only)
MAILBOXES         = ["Funds Ops", "NAV Alerts"]        # display names / SMTP as in Outlook
SAVE_TYPE         = 3                                   # olMsg (Unicode)
MAX_PATH_LENGTH   = 200                                 # full path length guard

# Outlook categories (optional; pre-create in Outlook):
CAT_SAVED         = "Saved"
CAT_DEFAULT       = "Not Saved"
CAT_SKIPPED       = "Skipped"

# ─────────────────────────── Helpers & Regex ─────────────────────────
ILLEGAL_FS = re.compile(r'[\\/:\*\?"<>|]')

# Month name mapping (case-insensitive, accepts both long & short)
_MONTH_MAP = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

# Patterns for explicit Year+Month (no ambiguity, no defaults)
# 1) MonthName + Year  (e.g., Aug 2025, August-2025, Aug2025)
_RX_MN_YR = re.compile(
    r'(?i)\b('
    r'jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|'
    r'aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)'
    r'[\s\-_./]*'
    r'((?:20)\d{2}|\d{2})\b'
)

# 2) Year + MonthName  (e.g., 2025 Aug, 2025-Aug, 2025August)
_RX_YR_MN = re.compile(
    r'(?i)\b((?:20)\d{2}|\d{2})[\s\-_./]*('
    r'jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|'
    r'aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b'
)

# 3) YYYY[sep]MM  (e.g., 2025-08, 2025/08, 202508)
_RX_YYYY_MM = re.compile(r'\b(20\d{2})[\s\-_./]?([01]\d)\b')

# 4) MM[sep]YYYY  (e.g., 08-2025, 08/2025, 082025)
_RX_MM_YYYY = re.compile(r'\b([01]\d)[\s\-_./]?(20\d{2})\b')

def month_folder(month: int) -> str:
    return f"{month:02d}-{dt.date(1900, month, 1).strftime('%B')}"

def _to_full_year(ystr: str) -> int:
    # Map 2-digit years to 2000–2079; else keep 4-digit year
    if len(ystr) == 2:
        yy = int(ystr)
        return 2000 + yy if yy <= 79 else 1900 + yy  # conservative cutoff
    return int(ystr)

def shorten_filename(target_dir: str, base: str, ext: str) -> str:
    """Truncate so full path ≤ MAX_PATH_LENGTH, keeping ≥30 chars of base."""
    full = os.path.join(target_dir, base + ext)
    if len(full) <= MAX_PATH_LENGTH:
        return base + ext
    allow = MAX_PATH_LENGTH - len(target_dir) - len(os.sep) - len(ext)
    return (base[: max(allow, 30)] + ext) if allow > 0 else ("trunc" + ext)

def _unique_path(base_path: Path) -> Path:
    """Append _01, _02 ... before suffix until path is unique."""
    if not base_path.exists():
        return base_path
    stem, suffix = base_path.stem, base_path.suffix
    n = 1
    while True:
        candidate = base_path.with_name(f"{stem}_{n:02d}{suffix}")
        if not candidate.exists():
            return candidate
        n += 1

def _attachment_names_safe(item) -> List[str]:
    names: List[str] = []
    try:
        cnt = item.Attachments.Count
        for i in range(1, cnt + 1):
            try:
                names.append(item.Attachments.Item(i).FileName)
            except Exception:
                pass
    except Exception:
        pass
    return names

def _extract_year_month_from_text(text: str) -> Optional[Tuple[int, int]]:
    """Return (year, month) if a clear Month+Year is present; else None."""
    if not text:
        return None

    # 1) MonthName + Year
    m = _RX_MN_YR.search(text)
    if m:
        mname = m.group(1).lower()
        year  = _to_full_year(m.group(2))
        month = _MONTH_MAP.get(mname[:3], None)
        if month and 1 <= month <= 12 and 2000 <= year <= 2099:
            return year, month

    # 2) Year + MonthName
    m = _RX_YR_MN.search(text)
    if m:
        year  = _to_full_year(m.group(1))
        mname = m.group(2).lower()
        month = _MONTH_MAP.get(mname[:3], None)
        if month and 1 <= month <= 12 and 2000 <= year <= 2099:
            return year, month

    # 3) YYYY[sep]MM (allows contiguous YYYYMM)
    m = _RX_YYYY_MM.search(text)
    if m:
        year  = int(m.group(1))
        month = int(m.group(2))
        if 1 <= month <= 12 and 2000 <= year <= 2099:
            return year, month

    # 4) MM[sep]YYYY (allows contiguous MMYYYY)
    m = _RX_MM_YYYY.search(text)
    if m:
        month = int(m.group(1))
        year  = int(m.group(2))
        if 1 <= month <= 12 and 2000 <= year <= 2099:
            return year, month

    return None

def detect_year_month_from_item(item) -> Optional[Tuple[int, int]]:
    """
    Parse only from Subject and Attachment file names. No fallback to ReceivedTime.
    Order: Subject first, then each attachment in order. Returns (year, month) or None.
    """
    # Subject
    ym = _extract_year_month_from_text(item.Subject or "")
    if ym:
        return ym
    # Attachments
    for name in _attachment_names_safe(item):
        ym = _extract_year_month_from_text(name or "")
        if ym:
            return ym
    return None

# ─────────────────────────── CATEGORY HELPER ────────────────────────
def set_category(itm: Any, cat: Optional[str]):
    if not cat:
        return
    try:
        itm.Categories = cat
        itm.Save()
    except Exception:
        pass

# ─────────────────────────── SENDER EMAIL (SMTP) ────────────────────
def get_smtp(item) -> str:
    """Best-effort resolution of SMTP address for an Outlook MailItem."""
    try:
        if item.SenderEmailType == "EX":  # Exchange
            try:
                exuser = item.Sender.GetExchangeUser()
                if exuser and exuser.PrimarySmtpAddress:
                    return exuser.PrimarySmtpAddress.lower()
            except Exception:
                pass
        addr = (item.SenderEmailAddress or "").lower()
        if addr:
            return addr
        return (item.Sender and item.Sender.Address or "").lower()
    except Exception:
        return (item.Sender and item.Sender.Address or "").lower()

# ─────────────────────────── ROUTING MAP UTILITIES ──────────────────
REQUIRED_COLS = ["SenderEmail", "GenericSender", "SubjectKey", "RootPath", "Attachment"]

def load_route_map(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str).fillna("")
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"EmailRoutes.xlsx missing columns: {missing}")
    df["SenderEmail_lc"]   = df["SenderEmail"].str.lower()
    df["GenericSender_lc"] = df["GenericSender"].str.lower()
    df["Attachment_lc"]    = df["Attachment"].str.lower()
    return df

def write_route_map(df: pd.DataFrame, path: str):
    out = df.drop(columns=[c for c in df.columns if c.endswith("_lc")], errors="ignore")
    try:
        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as xw:
            out.to_excel(xw, index=False)
    except Exception as e:
        print(f"[WARN] Could not update EmailRoutes.xlsx (is it open?): {e}")

def infer_root_for(sender: str, subject: str, df: pd.DataFrame) -> Tuple[str, bool]:
    """
    Infer a RootPath/Attachment for a previously unseen sender.
    Priority:
      1) Domain-specific non-generic row.
      2) Domain-generic rows whose SubjectKey matches.
      3) DEFAULT_SAVE_PATH with attachments=False.
    """
    domain = sender.split("@")[-1]
    # 1) Domain-specific non-generic row
    mask_domain = df["SenderEmail_lc"].str.endswith(domain)
    domain_rows = df[mask_domain]
    non_generic = domain_rows[domain_rows["GenericSender_lc"] == "no"]
    if not non_generic.empty:
        r = non_generic.iloc[0]
        return r["RootPath"], (r["Attachment_lc"] == "yes")
    # 2) Domain-specific generic rows that match SubjectKey
    generic = domain_rows[domain_rows["GenericSender_lc"] == "yes"]
    sub_lc = (subject or "").lower()
    for _, r in generic.iterrows():
        keys = re.split(r"[;,]", r["SubjectKey"]) if r["SubjectKey"] else []
        keys = [k.strip().lower() for k in keys if k.strip()]
        if any(k in sub_lc for k in keys):
            return r["RootPath"], (r["Attachment_lc"] == "yes")
    # 3) Fallback
    return DEFAULT_SAVE_PATH, False

def ensure_sender_row(sender: str, root: str, attach: bool, df: pd.DataFrame) -> pd.DataFrame:
    """Add a concrete non-generic row for sender if not present."""
    mask = df["SenderEmail_lc"] == sender.lower()
    if mask.any() and (df.loc[mask, "GenericSender_lc"] == "no").any():
        return df
    new = pd.DataFrame([{
        "SenderEmail": sender,
        "GenericSender": "no",
        "SubjectKey": "",
        "RootPath": root,
        "Attachment": "yes" if attach else "no"
    }])
    new["SenderEmail_lc"]   = sender.lower()
    new["GenericSender_lc"] = "no"
    new["Attachment_lc"]    = "yes" if attach else "no"
    return pd.concat([df, new], ignore_index=True)

# ─────────────────────────── SAVE ROUTINES ──────────────────────────
def save_message(msg, root: str) -> Path:
    """
    Save e-mail as .msg.
    Folder rules:
      • If (Year, Month) parsed from subject/attachments → <root>\Year\MM-MonthName\...
      • Else → <root>\<CURRENT_YEAR>\ (no month) for analyst manual assignment.
    """
    parsed = detect_year_month_from_item(msg)

    if parsed:
        year, month = parsed
        folder = Path(root) / str(year) / month_folder(month)
    else:
        folder = Path(root) / str(dt.date.today().year)  # no month
    folder.mkdir(parents=True, exist_ok=True)

    ts   = msg.ReceivedTime.strftime("%Y%m%d_%H%M%S")  # timestamp for uniqueness
    subj = ILLEGAL_FS.sub("-", msg.Subject or "")[:80]
    base = f"{ts}_{subj}" if subj else ts
    fname = shorten_filename(str(folder), base, ".msg")
    path  = _unique_path(Path(folder) / fname)

    msg.SaveAs(str(path), SAVE_TYPE)
    return path

def save_attachments(msg, root: str) -> List[Path]:
    """Save attachments directly under RootPath (spec unchanged)."""
    folder = Path(root)
    folder.mkdir(parents=True, exist_ok=True)

    ts   = msg.ReceivedTime.strftime("%Y%m%d_%H%M%S")
    subj = ILLEGAL_FS.sub("-", msg.Subject or "")[:40]
    saved: List[Path] = []

    cnt = msg.Attachments.Count
    for idx in range(1, cnt + 1):
        att = msg.Attachments.Item(idx)
        raw = ILLEGAL_FS.sub("-", att.FileName)
        base_name, ext = os.path.splitext(raw)
        base = f"{ts}_{subj}_{idx}_{base_name}" if base_name else f"{ts}_{subj}_{idx}"
        fname = shorten_filename(str(folder), base, ext)
        path  = _unique_path(folder / fname)
        att.SaveAsFile(str(path))
        saved.append(path)
    return saved

# ─────────────────────────── OUTLOOK HELPERS ────────────────────────
def ns():
    return win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

def open_inbox(display_name: str):
    try:
        root = ns().Folders[display_name]
        return root.Folders["Inbox"]
    except Exception:
        print(f"[WARN] Mailbox not found: {display_name}")
        return None

def restrict_items(inbox, start: dt.date, end: dt.date):
    start_dt = dt.datetime.combine(start, dt.time.min)
    end_dt   = dt.datetime.combine(end, dt.time.max)
    fmt = "%m/%d/%Y %I:%M %p"
    restriction = f"[ReceivedTime] >= '{start_dt.strftime(fmt)}' AND [ReceivedTime] <= '{end_dt.strftime(fmt)}'"
    items = inbox.Items.Restrict(restriction)
    items.Sort("[ReceivedTime]", True)
    return items

# ─────────────────────────── SUMMARY WRITER ─────────────────────────
def write_summary(rows: List[Dict[str, Any]]):
    try:
        new_df = pd.DataFrame(rows)
        if Path(SUMMARY_PATH).exists():
            old = pd.read_excel(SUMMARY_PATH)
            out = pd.concat([old, new_df], ignore_index=True)
        else:
            out = new_df
        with pd.ExcelWriter(SUMMARY_PATH, engine="openpyxl", mode="w") as xw:
            out.to_excel(xw, index=False)
    except Exception as e:
        print(f"[WARN] Could not write summary to {SUMMARY_PATH}: {e}")

# ─────────────────────────── ARCHIVE ENGINE ─────────────────────────
REQUIRED_COLS = ["SenderEmail", "GenericSender", "SubjectKey", "RootPath", "Attachment"]

def archive_window(start: dt.date, end: dt.date):
    df = load_route_map(MAP_PATH)
    unknown_rows: List[Dict[str, Any]] = []
    summary_rows: List[Dict[str, Any]] = []
    run_id = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for mbox in MAILBOXES:
        counts = {
            "total": 0, "saved_msg": 0, "saved_att": 0,
            "default_saves": 0, "skipped_no_attach": 0, "errors": 0,
        }

        inbox = open_inbox(mbox)
        if not inbox:
            summary_rows.append({
                "RunAt": run_id, "Mailbox": mbox, "Start": start, "End": end,
                "Total": 0, "SavedMsg": 0, "SavedAtt": 0,
                "DefaultSaves": 0, "SkippedNoAttachment": 0, "Errors": 0
            })
            continue

        items = restrict_items(inbox, start, end)
        itm = items.GetFirst()
        while itm:
            try:
                if itm.Class != 43:  # olMail
                    itm = items.GetNext(); continue

                counts["total"] += 1

                sender  = get_smtp(itm) or (itm.SenderName or "").lower()
                subject = itm.Subject or ""

                # Ensure a concrete non-generic row exists BEFORE deciding save behavior
                root, attach = infer_root_for(sender, subject, df)
                df = ensure_sender_row(sender, root, attach, df)

                if attach:
                    if itm.Attachments.Count > 0:
                        save_attachments(itm, root)
                        set_category(itm, CAT_SAVED)
                        counts["saved_att"] += 1
                        if root == DEFAULT_SAVE_PATH:
                            counts["default_saves"] += 1
                    else:
                        set_category(itm, CAT_SKIPPED)
                        counts["skipped_no_attach"] += 1
                else:
                    save_message(itm, root)
                    set_category(itm, CAT_SAVED)
                    counts["saved_msg"] += 1
                    if root == DEFAULT_SAVE_PATH:
                        counts["default_saves"] += 1

            except Exception as e:
                counts["errors"] += 1
                try:
                    unknown_rows.append({"RunAt": run_id, "Mailbox": mbox, "Sender": sender, "Subject": subject, "Error": str(e)})
                except Exception:
                    unknown_rows.append({"RunAt": run_id, "Mailbox": mbox, "Sender": "", "Subject": "", "Error": str(e)})
                set_category(itm, CAT_DEFAULT)
            finally:
                itm = items.GetNext()

        summary_rows.append({
            "RunAt": run_id, "Mailbox": mbox, "Start": start, "End": end,
            "Total": counts["total"], "SavedMsg": counts["saved_msg"], "SavedAtt": counts["saved_att"],
            "DefaultSaves": counts["default_saves"], "SkippedNoAttachment": counts["skipped_no_attach"],
            "Errors": counts["errors"],
        })

    # Persist any new sender rows back to EmailRoutes.xlsx
    write_route_map(df, MAP_PATH)

    # Dump unknowns/errors
    if unknown_rows:
        try:
            pd.DataFrame(unknown_rows).to_csv(
                UNKNOWN_CSV, index=False, mode="a", header=not Path(UNKNOWN_CSV).exists()
            )
        except Exception as e:
            print(f"[WARN] Could not write unknown_senders.csv: {e}")

    # Append summary
    write_summary(summary_rows)

# ─────────────────────────── CLI ─────────────────────────
def parse_date(s: str) -> dt.date:
    return dt.datetime.strptime(s, "%Y-%m-%d").date()

def main():
    ap = argparse.ArgumentParser(description="Archive Outlook mailboxes")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--yesterday", action="store_true")
    g.add_argument("--date", type=parse_date, metavar="YYYY-MM-DD")
    g.add_argument("--range", nargs=2, type=parse_date, metavar=("FROM", "TO"))
    args = ap.parse_args()

    if args.yesterday:
        d = dt.date.today() - dt.timedelta(days=1)
        archive_window(d, d)
    elif args.date:
        archive_window(args.date, args.date)
    else:
        s, e = sorted(args.range)
        archive_window(s, e)

if __name__ == "__main__":
    main()