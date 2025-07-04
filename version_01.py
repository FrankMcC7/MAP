#!/usr/bin/env python
# email_archiver.py  –  tested Outlook 365 / Python 3.11
import argparse, datetime as dt, os, re, sys, pandas as pd
import win32com.client as win32
from dateutil import parser as nlp                                # fuzzy parse
try:  # optional GUI
    import tkinter as tk; from tkinter import ttk, messagebox
    from tkcalendar import DateEntry                              # pip install tkcalendar
except ImportError:
    tk = None

MAP   = r"\\share\\EmailRoutes.xlsx"
ULOG  = r"\\share\\unknown_senders.csv"
BOXES = ["Funds Ops", "NAV Alerts"]                               # shared mailboxes
SAVE_TYPE = 3                                                     # olMsg
ISO_RGX  = re.compile(r"\b(20\d{2})[-_.]?([01]?\d)\b")
MON_RGX  = re.compile(r"(?i)\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[-_/\. ]*(20\d{2})")

# ── map helpers ─────────────────────────────────────────────────────────────
def load_map_sheet():
    df = pd.read_excel(MAP, dtype=str).fillna("")
    df["GenericSender"] = df["GenericSender"].str.lower()
    return df

def build_dicts(df):
    exact, generic = {}, {}
    for _, r in df.iterrows():
        sender = r["SenderEmail"].lower()
        if r["GenericSender"] == "yes":
            keys = [k.strip().lower() for k in r["SubjectKey"].split(",") if k.strip()]
            generic.setdefault(sender, []).append((keys, r["RootPath"]))
        else:
            exact[sender] = r["RootPath"]
    return exact, generic

def add_map_row(df, sender, root):
    if df["SenderEmail"].str.lower().eq(sender.lower()).any():
        return False                    # already present
    df.loc[len(df)] = [sender, "no", "", root]
    return True

# ── routing logic ──────────────────────────────────────────────────────────
def infer_route(sender, subject, df, exact, generic):
    domain = sender.split("@")[-1]
    rows   = df[df["SenderEmail"].str.lower().str.endswith(domain.lower())]

    # clone first non-generic row for domain
    ng = rows[rows["GenericSender"] == "no"]
    if not ng.empty:
        root = ng.iloc[0]["RootPath"]
        if add_map_row(df, sender, root): exact[sender] = root
        return root

    # match generic rows by keyword
    for _, r in rows.iterrows():
        keys = [k.strip().lower() for k in r["SubjectKey"].split(",") if k.strip()]
        if any(k in subject.lower() for k in keys):
            root = r["RootPath"]
            if add_map_row(df, sender, root): exact[sender] = root
            return root
    return None

def resolve_path(sender, subject, df, exact, generic):
    sender_lc = sender.lower()
    if sender_lc in exact:
        return exact[sender_lc]
    if sender_lc in generic:            # exact match to a generic row
        for keys, p in generic[sender_lc]:
            if any(k in subject.lower() for k in keys):
                return p
    return infer_route(sender_lc, subject, df, exact, generic)

# ── date window & filters ──────────────────────────────────────────────────
def win_date(s): return dt.datetime.strptime(s, "%Y-%m-%d").date()
def build_filter(start, end):
    fmt = "%m/%d/%Y %I:%M %p"                           # Outlook Restrict format :contentReference[oaicite:0]{index=0}
    s = dt.datetime.combine(start, dt.time.min).strftime(fmt)
    e = dt.datetime.combine(end,   dt.time.max).strftime(fmt)
    return f"[ReceivedTime] >= '{s}' AND [ReceivedTime] <= '{e}'"

# ── period derivation ──────────────────────────────────────────────────────
def extract_period(msg):
    blob = " ".join([msg.Subject or "", *[a.FileName for a in msg.Attachments]])
    try: d = nlp.parse(blob, fuzzy=True, default=msg.ReceivedTime)     # :contentReference[oaicite:1]{index=1}
    except Exception: d = None
    if d: return d.year, d.month
    for rgx in (MON_RGX, ISO_RGX):
        m = rgx.search(blob)
        if m:
            if rgx is MON_RGX:
                return int(m.group(2)), dt.datetime.strptime(m.group(1)[:3], "%b").month
            return int(m.group(1)), int(m.group(2))
    return msg.ReceivedTime.year, msg.ReceivedTime.month

def save_msg(item, root):
    y, m = extract_period(item)
    path = os.path.join(root, str(y), dt.date(1900, m, 1).strftime("%B"))
    os.makedirs(path, exist_ok=True)
    safe = re.sub(r'[\\/:*?"<>|]', "-", item.Subject)[:80]
    item.SaveAs(os.path.join(path, f"{safe}_{item.EntryID}.msg"), SAVE_TYPE)

# ── archiver core ──────────────────────────────────────────────────────────
def archive(start, end):
    df = load_map_sheet()
    exact, generic = build_dicts(df)
    updated, unknown = False, []

    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    flt = build_filter(start, end)

    for box in BOXES:
        inbox = outlook.Folders(box).Folders("Inbox").Items.Restrict(flt)
        for itm in inbox:
            path = resolve_path(itm.SenderEmailAddress, itm.Subject, df, exact, generic)
            if path:
                save_msg(itm, path)
            else:
                unknown.append([itm.SenderEmailAddress, itm.Subject])

    if updated: df.to_excel(MAP, index=False)           # overwrite safely :contentReference[oaicite:2]{index=2}
    if unknown:
        pd.DataFrame(unknown).to_csv(ULOG, mode="a", header=False, index=False)

# ── CLI + GUI front-ends ───────────────────────────────────────────────────
def run_gui():
    if tk is None:
        print("tkcalendar not installed"); sys.exit(1)
    root = tk.Tk(); root.title("Select Archive Dates")
    mode = tk.StringVar(value="single")
    def swap(): end.state(["!disabled"] if mode.get()=="range" else ["disabled"])
    def go():
        try:
            if mode.get()=="single":
                d = start.get_date(); archive(d, d)
            else:
                s, e = start.get_date(), end.get_date(); archive(min(s, e), max(s, e))
            messagebox.showinfo("Done", "Archiving complete.")
        except Exception as err:
            messagebox.showerror("Error", str(err))
    ttk.Radiobutton(root, text="Single day", variable=mode, value="single", command=swap).grid(row=0, column=0, sticky="w")
    ttk.Radiobutton(root, text="Date range", variable=mode, value="range",  command=swap).grid(row=1, column=0, sticky="w")
    start = DateEntry(root); end = DateEntry(root); end.state(["disabled"])
    ttk.Label(root, text="Start").grid(row=0, column=1); start.grid(row=0, column=2)
    ttk.Label(root, text="End").grid(row=1, column=1);   end.grid(row=1, column=2)
    ttk.Button(root, text="Run", command=go).grid(row=2, column=0, columnspan=3, pady=8)
    root.mainloop()                                      # :contentReference[oaicite:3]{index=3}

def cli():
    a = argparse.ArgumentParser()
    g = a.add_mutually_exclusive_group()
    g.add_argument("--yesterday", action="store_true")
    g.add_argument("--date",  type=win_date)
    g.add_argument("--range", nargs=2, type=win_date)
    g.add_argument("--ui", action="store_true")
    opts = a.parse_args()
    if opts.ui:       run_gui()
    elif opts.yesterday:
        d = dt.date.today() - dt.timedelta(1); archive(d, d)
    elif opts.date:   archive(opts.date, opts.date)
    elif opts.range:  s, e = min(opts.range), max(opts.range); archive(s, e)
    else:             print("Use --ui, --yesterday, --date or --range"); sys.exit(1)

if __name__ == "__main__":
    cli()
