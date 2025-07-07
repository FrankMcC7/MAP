#!/usr/bin/env python3
from __future__ import annotations
import argparse
import datetime as dt
import os
import re
import sys
from typing import Tuple, List, Any

import pandas as pd
import win32com.client as win32
from dateutil import parser as nlp

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
    from tkcalendar import DateEntry
except ImportError:
    tk = None  # GUI is optional

# ───────────────────────────────────────── USER CONFIG ─────────────────────────────────────────
MAP_PATH            = r"\\share\\EmailRoutes.xlsx"
UNKNOWN_CSV         = r"\\share\\unknown_senders.csv"
DEFAULT_SAVE_PATH   = r"\\share\\DefaultArchive"
SUMMARY_PATH        = r"\\share\\ArchiveSummary.xlsx"  # summary of each run
# For shared mailboxes, specify either display name or SMTP address
MAILBOXES           = ["Funds Ops", "NAV Alerts"]
SAVE_TYPE           = 3   # olMsg Unicode
MAX_PATH_LENGTH     = 200  # max full path length before truncation

# ───────────────────────────────────────── INTERNAL REGEX ──────────────────────────────────────
ISO_RGX = re.compile(r"\b(20\d{2})[-_.]?([01]?\d)\b")
MON_RGX = re.compile(r"(?i)\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[-_/\\. ]*(20\d{2})")

# ───────────────────────────────────────── MAP HELPERS ─────────────────────────────────────────
def load_map_sheet() -> pd.DataFrame:
    df = pd.read_excel(MAP_PATH, dtype=str).fillna("")
    expected = {"SenderEmail", "GenericSender", "SubjectKey", "RootPath"}
    missing = expected.difference(df.columns)
    if missing:
        raise ValueError(f"EmailRoutes.xlsx missing columns: {', '.join(missing)}")
    df["GenericSender"] = df["GenericSender"].str.lower()
    return df


def build_maps(df: pd.DataFrame) -> Tuple[dict, dict]:
    exact: dict[str, str] = {}
    generic: dict[str, List[Tuple[List[str], str]]] = {}
    for _, row in df.iterrows():
        sender = row["SenderEmail"].lower()
        if row["GenericSender"] == "yes":
            keys = [k.strip().lower() for k in row["SubjectKey"].split(",") if k.strip()]
            generic.setdefault(sender, []).append((keys, row["RootPath"]))
        else:
            exact[sender] = row["RootPath"]
    return exact, generic


def add_row(df: pd.DataFrame, sender: str, root: str) -> bool:
    if sender.lower() in df["SenderEmail"].str.lower().values:
        return False
    df.loc[len(df)] = [sender, "no", "", root]
    return True


def infer_route(sender: str, subject: str, df: pd.DataFrame,
                exact: dict, generic: dict) -> str | None:
    domain = sender.split("@")[-1].lower()
    candidates = df[df["SenderEmail"].str.lower().str.endswith(domain)]
    non_generic = candidates[candidates["GenericSender"] == "no"]
    if not non_generic.empty:
        root = non_generic.iloc[0]["RootPath"]
        if add_row(df, sender, root):
            exact[sender.lower()] = root
        return root
    for _, r in candidates.iterrows():
        keys = [k.strip().lower() for k in r["SubjectKey"].split(",") if k.strip()]
        if any(k in subject.lower() for k in keys):
            root = r["RootPath"]
            if add_row(df, sender, root):
                exact[sender.lower()] = root
            return root
    return None


def resolve_path(sender: str, subject: str, df: pd.DataFrame,
                 exact: dict, generic: dict) -> str | None:
    sender_lc = sender.lower()
    if sender_lc in exact:
        return exact[sender_lc]
    if sender_lc in generic:
        for keywords, root in generic[sender_lc]:
            if any(k in subject.lower() for k in keywords):
                return root
    return infer_route(sender, subject, df, exact, generic)


def parse_cli_date(s: str) -> dt.date:
    return dt.datetime.strptime(s, "%Y-%m-%d").date()


def outlook_restrict_clause(start: dt.date, end: dt.date) -> str:
    fmt = "%m/%d/%Y %I:%M %p"
    s = dt.datetime.combine(start, dt.time.min).strftime(fmt)
    e = dt.datetime.combine(end,   dt.time.max).strftime(fmt)
    return f"[ReceivedTime] >= '{s}' AND [ReceivedTime] <= '{e}'"


def detect_period(item) -> Tuple[int, int]:
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

# ───────────────────────────────────────── PATH SHORTENER ────────────────────────────────────
def shorten_filename(target_dir: str, base_name: str, ext: str) -> str:
    full_path = os.path.join(target_dir, base_name + ext)
    if len(full_path) <= MAX_PATH_LENGTH:
        return base_name + ext
    allowed = MAX_PATH_LENGTH - len(target_dir) - len(os.sep) - len(ext)
    if allowed <= 0:
        return base_name[:50] + ext
    return base_name[:allowed] + ext

# ───────────────────────────────────────── SAVE MESSAGE ─────────────────────────────────────────
def save_message(item, root: str):
    year, month = detect_period(item)
    month_name = dt.date(1900, month, 1).strftime("%B")
    target_dir = os.path.join(root, str(year), month_name)
    os.makedirs(target_dir, exist_ok=True)

    safe_subject = re.sub(r'[\\/:*?"<>|]', "-", item.Subject or "")[:80]
    base_name = f"{safe_subject}_{item.EntryID}"
    ext = ".msg"
    fname = shorten_filename(target_dir, base_name, ext)
    item.SaveAs(os.path.join(target_dir, fname), SAVE_TYPE)

# ───────────────────────────────────────── CATEGORY HELPER ────────────────────────────────────
def set_category(item: Any, category: str):
    item.Categories = category
    item.Save()

# ───────────────────────────────────────── ARCHIVE CORE ─────────────────────────────────────────
def archive_window(start: dt.date, end: dt.date):
    df = load_map_sheet()
    exact, generic = build_maps(df)
    initial_rows = len(df)

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    flt = outlook_restrict_clause(start, end)

    total = saved_mapped = saved_default = errors = 0
    unresolved: List[Tuple[str, str]] = []

    for mbox in MAILBOXES:
        # Try local store, else shared mailbox
        try:
            store = ns.Folders(mbox)
            inbox_folder = store.Folders("Inbox")
        except Exception:
            recipient = ns.CreateRecipient(mbox)
            recipient.Resolve()
            if not recipient.Resolved:
                raise ValueError(f"Mailbox '{mbox}' not found or not resolved")
            inbox_folder = ns.GetSharedDefaultFolder(recipient, win32.constants.olFolderInbox)
        items = inbox_folder.Items.Restrict(flt)

        for itm in items:
            total += 1
            try:
                path = resolve_path(
                    itm.SenderEmailAddress or "", itm.Subject or "", df, exact, generic
                )
                if path:
                    save_message(itm, path)
                    set_category(itm, "Saved")
                    saved_mapped += 1
                else:
                    save_message(itm, DEFAULT_SAVE_PATH)
                    set_category(itm, "Not Saved")
                    saved_default += 1
                    unresolved.append((itm.SenderEmailAddress or "", itm.Subject or ""))
            except Exception:
                errors += 1

    if len(df) > initial_rows:
        df.to_excel(MAP_PATH, index=False, engine="openpyxl")

    if unresolved:
        pd.DataFrame(unresolved, columns=["Sender", "Subject"]).to_csv(
            UNKNOWN_CSV, mode="a", header=False, index=False
        )

    summary_df = pd.DataFrame([{  
        "RunDate": dt.date.today(),
        "TotalEmails": total,
        "SavedMapped": saved_mapped,
        "SavedDefault": saved_default,
        "ErrorsSkipped": errors
    }])
    summary_df.to_excel(SUMMARY_PATH, index=False)

# ───────────────────────────────────────── GUI FRONT END ────────────────────────────────────────
def launch_gui():
    if tk is None:
        print("tkcalendar not installed", file=sys.stderr)
        sys.exit(1)

    root = tk.Tk()
    root.title("Email Archiver – Select Dates")

    mode = tk.StringVar(value="single")
    def toggle(): end_entry.state(["!disabled"] if mode.get()=="range" else ["disabled"])
    def run():
        try:
            if mode.get()=="single": d=start_entry.get_date(); archive_window(d,d)
            else: s,e=start_entry.get_date(),end_entry.get_date(); archive_window(min(s,e),max(s,e))
            messagebox.showinfo("Done","Archiving complete.")
        except Exception as err:
            messagebox.showerror("Error",str(err))

    ttk.Radiobutton(root,text="Single day",variable=mode,value="single",command=toggle).grid(row=0,column=0,sticky="w")
    ttk.Radiobutton(root,text="Date range",variable=mode,value="range",command=toggle ).grid(row=1,column=0,sticky="w")
    ttk.Label(root,text="Start").grid(row=0,column=1,padx=4)
    start_entry=DateEntry(root); start_entry.grid(row=0,column=2,padx=4)
    ttk.Label(root,text="End").grid(row=1,column=1,padx=4)
    end_entry=DateEntry(root); end_entry.state(["disabled"]); end_entry.grid(row=1,column=2,padx=4)
    ttk.Button(root,text="Run",command=run).grid(row=2,column=0,columnspan=3,pady=8)
    root.mainloop()

# ───────────────────────────────────────── CLI ENTRY ────────────────────────────────────────────
def main():
    parser=argparse.ArgumentParser(description="Archive shared‑mailbox e‑mail")
    group=parser.add_mutually_exclusive_group()
    group.add_argument("--yesterday",action="store_true",help="Previous calendar day")
    group.add_argument("--date",type=parse_cli_date,metavar="YYYY-MM-DD",help="Single date")
    group.add_argument("--range",nargs=2,type=parse_cli_date,metavar=("FROM","TO"),help="Inclusive date range")
    group.add_argument("--ui",action="store_true",help="Tkinter calendar picker")
    args=parser.parse_args()
    if args.ui: launch_gui(); return
    if args.yesterday: d=dt.date.today()-dt.timedelta(days=1); archive_window(d,d)
    elif args.date: archive_window(args.date,args.date)
    elif args.range: s,e=min(args.range),max(args.range); archive_window(s,e)
    else: parser.error("Specify --ui, --yesterday, --date, or --range")

if __name__=="__main__": main()
