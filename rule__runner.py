"""
outlook_rule_runner.py  –  Outlook Rule Runner (multi-SenderMatch, TZ-safe)
"""

import argparse, logging, re, sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import win32com.client as win32  # pywin32

# ───────────── CONFIG ─────────────
RULEBOOK_PATH = Path(__file__).with_name("Outlook_Rule_Template.xlsx")
LOG_PATH = Path(__file__).with_suffix(".log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_PATH, "a", "utf-8"), logging.StreamHandler()],
)

# ───────── Outlook helpers ────────
def get_ns(): return win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
def open_root(name):
    try: return get_ns().Folders[name]
    except Exception: logging.error(f"Mailbox '{name}' not found."); return None
def ensure_folder(mbx, path):
    root = open_root(mbx)
    if not root: return None
    f = root
    for p in re.split(r"[\\/]+", path.strip("\\/")):
        try: f = f.Folders[p]
        except Exception: f = f.Folders.Add(p)
    return f
def categories_of(mail):
    return [c.strip().casefold() for c in (mail.Categories or "").split(",") if c]

# ───────── Predicate helpers ──────
def _clean(v):
    if v is None: return None
    if isinstance(v, float) and pd.isna(v): return None
    if isinstance(v, str) and v.strip() == "": return None
    return str(v)

def build_eq_pred(v):
    v = _clean(v)
    if v is None: return lambda *_: True
    if v.startswith("/") and v.endswith("/"):
        pat = re.compile(v[1:-1], re.IGNORECASE)
        return lambda s,*_: bool(pat.search(s or ""))
    targets = [t.casefold() for t in v.split(",") if t.strip()]
    return lambda s,*_: (s or "").casefold() in targets

def build_kw_regex(v):
    v = _clean(v)
    if v is None: return lambda *_: True
    if v.startswith("/") and v.endswith("/"):
        pat = re.compile(v[1:-1], re.IGNORECASE)
        return lambda s,*_: bool(pat.search(s or ""))
    kws = [k.casefold() for k in v.split(",") if k.strip()]
    return lambda s,*_: any(k in (s or "").casefold() for k in kws)

# ───────── TZ helper ──────────────
def _naive(dt: datetime) -> datetime:
    return dt.replace(tzinfo=None) if dt.tzinfo else dt

def received_pred(period, start=None, end=None):
    if period == "all": return lambda *_: True
    if period == "yesterday":
        today = datetime.now().date()
        start = datetime.combine(today - timedelta(days=1), datetime.min.time())
        end   = datetime.combine(today - timedelta(days=1), datetime.max.time())
    elif period == "thismonth":
        today = datetime.now().date()
        start = datetime(today.year, today.month, 1)
        end   = datetime.now()
    elif period == "custom":
        pass
    else: raise ValueError("bad period")

    s_n, e_n = _naive(start), _naive(end)
    return lambda rt,*_: s_n <= _naive(rt) <= e_n

# ───────── Load rules ─────────────
def load_rules(xls):
    df = pd.read_excel(xls, sheet_name="Rules").where(pd.notnull, None)
    if "Enabled" in df: df = df[~df.Enabled.str.lower().eq("no")]
    for col in ("Mailbox","RuleName","ActionMoveTo"):
        if col not in df.columns: raise ValueError(f"Missing {col}")
    return df.reset_index(drop=True)

# ───────── Core mailbox runner ────
def run_mailbox(df, src, time_ok):
    root = open_root(src);  inbox = root and root.Folders["Inbox"]
    if not inbox: return
    rules=[dict(
        RuleName=r.RuleName,
        Sender=build_eq_pred(r.SenderMatch),
        Subj = build_kw_regex(r.SubjectContains),
        Cat  = build_kw_regex(r.Category),
        Target=r.ActionMoveTo) for _,r in df.iterrows()]
    for mail in list(inbox.Items):
        try:
            if not time_ok(mail.ReceivedTime): continue
            sender, subj, cats = mail.SenderName, mail.Subject, ",".join(categories_of(mail))
            for R in rules:
                if R["Sender"](sender) and R["Subj"](subj) and R["Cat"](cats):
                    dest_desc = R["Target"] or f"Archive/{R['RuleName']}"
                    parts = re.split(r"[\\/]+", dest_desc, 1)
                    tgt_mbx, path = (src, parts[0]) if len(parts)==1 else parts
                    dest = ensure_folder(tgt_mbx, path)
                    if dest:
                        mail.Move(dest); mail.UnRead=False
                        logging.info(f"{src}|{R['RuleName']}→{tgt_mbx}/{path}")
                    break
        except Exception as e: logging.error(f"Error mail in {src}: {e}")

# ───────── CLI / interactive ──────
def choose_period():
    m={"1":"yesterday","2":"thismonth","3":"all","4":"custom"}
    print("1 Yesterday  2 This month  3 All  4 Custom"); c=input("> ").strip()
    per=m.get(c,"all"); s=e=None
    if per=="custom":
        s=input("Start YYYY-MM-DD: "); e=input("End YYYY-MM-DD: ")
        try:
            s=datetime.strptime(s,"%Y-%m-%d")
            e=datetime.strptime(e,"%Y-%m-%d")+timedelta(days=1,seconds=-1)
        except: print("Bad date, using all"); per="all"
    return per,s,e

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("-p","--period",choices=["yesterday","thismonth","all","custom"])
    ap.add_argument("--start"); ap.add_argument("--end"); a=ap.parse_args()
    if a.period:
        if a.period=="custom":
            if not (a.start and a.end): sys.exit("need --start & --end")
            s=datetime.strptime(a.start,"%Y-%m-%d")
            e=datetime.strptime(a.end,"%Y-%m-%d")+timedelta(days=1,seconds=-1)
        else: s=e=None; 
        per=a.period
    else: per,s,e=choose_period()

    logging.info(f"=== Outlook Rule Runner period={per} ===")
    df=load_rules(RULEBOOK_PATH); pred=received_pred(per,s,e)
    for mbx,g in df.groupby("Mailbox"): run_mailbox(g, mbx, pred)
    logging.info("=== done ===")

if __name__=="__main__": main()
