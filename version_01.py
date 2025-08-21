#!/usr/bin/env python3
# Email Archiver – v2.5  (Aug 2025)
#
# Changes in v2.5
#   • Saved .msg filenames now use a clean pattern:
#       <YYYYMMDD_HHMMSS>_<Subject>.msg
#     (previously appended Outlook EntryID)
#   • Collisions are handled by appending _(2), _(3), ... before the extension.
#
# From v2.4
#   • Month folders use “MM-MonthName” (e.g. 06-June)
#   • New Excel column “Attachment” (Yes / No):
#       – Yes  ➜ save only the attachments into RootPath
#                (skip the e-mail itself)
#                – if no attachment exists the message is *skipped*
#       – No   ➜ save the whole message under  <RootPath>\YYYY\MM-MonthName
#   • Optional Outlook category “Skipped” applied when
#     Attachment = Yes but the message has no attachments.
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
CAT_SKIPPED       = "Skipped"      # ← new

# ──────────────────────────── REGEX HELPERS ─────────────────────────
ISO_RGX = re.compile(r"\b(20\d{2})[-_.]?([01]?\d)\b")
MON_RGX = re.compile(
    r"(?i)\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[-_/\. ]*(20\d{2})"
)
ILLEGAL_FS = re.compile(r'[\\/:\*\?"<>|]')

# ─────────────────────────── ROUTE MAP LOADER ───────────────────────
def load_map_sheet() -> pd.DataFrame:
    df = pd.read_excel(MAP_PATH, dtype=str).fillna("")
    expected = {
        "SenderEmail",
        "GenericSender",
        "SubjectKey",
        "RootPath",
        "Attachment",
    }
    miss = expected.difference(df.columns)
    if miss:
        raise ValueError(f"EmailRoutes.xlsx missing columns: {', '.join(miss)}")
    df["GenericSender"] = df["GenericSender"].str.lower()
    df["Attachment"] = df["Attachment"].str.lower().map({"yes": "yes"}).fillna("no")
    return df


def build_maps(
    df: pd.DataFrame,
) -> Tuple[dict[str, Tuple[str, bool]], dict[str, List[Tuple[List[str], str, bool]]]]:
    """Return (exact, generic) where each value holds (root_path, attach_flag)"""
    exact: dict[str, Tuple[str, bool]] = {}
    generic: dict[str, List[Tuple[List[str], str, bool]]] = {}
    for _, row in df.iterrows():
        sender = row["SenderEmail"].lower()
        root = row["RootPath"]
        attach = row["Attachment"] == "yes"
        if row["GenericSender"] == "yes":
            keys = [k.strip().lower() for k in row["SubjectKey"].split(",") if k.strip()]
            generic.setdefault(sender, []).append((keys, root, attach))
        else:
            exact[sender] = (root, attach)
    return exact, generic


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


def _unique_path(base_path: Path) -> Path:
    """
    Ensure the path is unique by appending _(n) before the suffix if needed.
    base_path already includes the suffix (.msg or otherwise).
    """
    if not base_path.exists():
        return base_path
    stem, suffix = base_path.stem, base_path.suffix
    n = 2
    while True:
        candidate = base_path.with_name(f"{stem}_({n}){suffix}")
        if not candidate.exists():
            return candidate
        n += 1


# ─────────────────────────── SAVE ROUTINES ──────────────────────────
def save_message(msg, root: str):
    """
    Save the whole e-mail as .msg using a clean, unique filename:
    <YYYYMMDD_HHMMSS>_<Subject>.msg
    """
    yr, mo = detect_period(msg)
    folder = Path(root) / str(yr) / month_folder(mo)
    folder.mkdir(parents=True, exist_ok=True)

    ts = msg.ReceivedTime.strftime("%Y%m%d_%H%M%S")
    subj = ILLEGAL_FS.sub("-", msg.Subject or "")[:80]

    base = f"{ts}_{subj}" if subj else ts
    fname = shorten_filename(str(folder), base, ".msg")
    path = _unique_path(folder / fname)

    msg.SaveAs(str(path), SAVE_TYPE)


def save_attachments(msg, root: str):
    Path(root).mkdir(parents=True, exist_ok=True)
    ts = msg.ReceivedTime.strftime("%Y%m%d_%H%M%S")
    subj = ILLEGAL_FS.sub("-", msg.Subject or "")[:40]

    for idx, att in enumerate(msg.Attachments, start=1):
        fname_raw = ILLEGAL_FS.sub("-", att.FileName)
        base = f"{ts}_{subj}_{idx}_{fname_raw}"
        name = shorten_filename(root, base, "")
        path = Path(root) / name
        # Attachments rarely collide, but make it robust:
        path = _unique_path(path.with_suffix(path.suffix))
        att.SaveAsFile(str(path))


# ─────────────────────────── DATE / NLP HELPERS ─────────────────────
def detect_period(item) -> Tuple[int, int]:
    """Try to infer accounting period from subject / attachment names."""
    blob = " ".join([item.Subject or ""] + [att.FileName for att in item.Attachments])
    try:
        d = nlp.parse(blob, fuzzy=True, default=item.ReceivedTime)
        return d.year, d.month
    except Exception:
        pass
    for rgx in (MON_RGX, ISO_RGX):
        m = rgx.search(blob)
        if m:
            if rgx is MON_RGX:
                month = dt.datetime.strptime(m.group(1)[:3], "%b").month
                return int(m.group(2)), month
            return int(m.group(1)), int(m.group(2))
    return item.ReceivedTime.year, item.ReceivedTime.month


# ─────────────────────────── CATEGORY HELPER ────────────────────────
def set_category(itm: Any, cat: str):
    """Overwrite category as per previous versions to avoid behaviour change."""
    itm.Categories = cat
    itm.Save()


# ─────────────────────────── ROUTE RESOLVER ────────────────────────
def add_row(df: pd.DataFrame, sender: str, root: str):
    """Persist newly inferred exact route (Attachment = ‘no’ by default)."""
    df.loc[len(df)] = [sender, "no", "", root, "no"]


def resolve_route(
    sender: str,
    subject: str,
    df: pd.DataFrame,
    exact,
    generic,
) -> Tuple[str, bool] | None:
    key = sender.lower()
    if key in exact:
        return exact[key]

    if key in generic:
        for keys, root, attach in generic[key]:
            if any(k in subject.lower() for k in keys):
                return root, attach

    # Infer from domain → default Attachment = no
    domain = sender.split("@")[-1].lower()
    cand = df[df["SenderEmail"].str.lower().str.endswith(domain)]
    non_gen = cand[cand["GenericSender"] == "no"]
    if not non_gen.empty:
        root = non_gen.iloc[0]["RootPath"]
        add_row(df, sender, root)
        return root, False
    for _, r in cand.iterrows():
        if any(k in subject.lower() for k in r["SubjectKey"].split(",")):
            add_row(df, sender, r["RootPath"])
            return r["RootPath"], False
    return None


# ─────────────────────────── ARCHIVE ENGINE ─────────────────────────
def archive_window(start: dt.date, end: dt.date):
    df = load_map_sheet()
    exact, generic = build_maps(df)
    init_len = len(df)

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    flt = f"[ReceivedTime] >= '{start.strftime('%m/%d/%Y')}' AND [ReceivedTime] <= '{end.strftime('%m/%d/%Y')} 11:59 PM'"

    stats = dict(total=0, mapped=0, fallback=0, skipped=0, errors=0)
    unresolved: List[Tuple[str, str]] = []

    for mbx in MAILBOXES:
        try:
            inbox = ns.Folders(mbx).Folders("Inbox")
        except Exception:
            rcpt = ns.CreateRecipient(mbx)
            rcpt.Resolve()
            inbox = ns.GetSharedDefaultFolder(rcpt, win32.constants.olFolderInbox)

        for itm in inbox.Items.Restrict(flt):
            stats["total"] += 1
            try:
                route = resolve_route(
                    itm.SenderEmailAddress or "", itm.Subject or "", df, exact, generic
                )
                if route:
                    root, attach_only = route
                    if attach_only:
                        if itm.Attachments.Count:
                            save_attachments(itm, root)
                            set_category(itm, CAT_SAVED)
                            stats["mapped"] += 1
                        else:
                            # Attachment = Yes but no files → skip
                            set_category(itm, CAT_SKIPPED)
                            stats["skipped"] += 1
                    else:
                        save_message(itm, root)
                        set_category(itm, CAT_SAVED)
                        stats["mapped"] += 1
                else:
                    # Default path saving
                    save_message(itm, DEFAULT_SAVE_PATH)
                    set_category(itm, CAT_DEFAULT)
                    stats["fallback"] += 1
                    unresolved.append((itm.SenderEmailAddress or "", itm.Subject or ""))
            except Exception:
                stats["errors"] += 1

    # Persist inferred routes
    if len(df) > init_len:
        df.to_excel(MAP_PATH, index=False, engine="openpyxl")

    # Unknown senders
    if unresolved:
        pd.DataFrame(unresolved, columns=["Sender", "Subject"]).to_csv(
            UNKNOWN_CSV, mode="a", header=False, index=False
        )

    # Summary
    pd.DataFrame(
        [
            {
                "RunDate": dt.date.today(),
                "Total": stats["total"],
                "SavedMapped": stats["mapped"],
                "SavedDefault": stats["fallback"],
                "Skipped": stats["skipped"],
                "Errors": stats["errors"],
            }
        ]
    ).to_excel(SUMMARY_PATH, index=False)


# ─────────────────────────── CLI / ENTRY ────────────────────────────
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
