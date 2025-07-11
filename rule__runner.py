"""
outlook_rule_runner.py  –  TZ-safe, multi-SenderMatch, fast date filter
"""

import argparse, logging, re, sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import win32com.client as win32          # pywin32

# ───────────── CONFIG ─────────────
RULEBOOK_PATH = Path(__file__).with_name("Outlook_Rule_Template.xlsx")
LOG_PATH      = Path(__file__).with_suffix(".log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_PATH, "a", "utf-8"), logging.StreamHandler()],
)

# ───────── Outlook helpers ────────
def ns(): return win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
def open_root(disp): 
    try: return ns().Folders[disp]
    except Exception: logging.error(f"Mailbox '{disp}' not found."); return None
def ensure_folder(mbx, path):
    root=open_root(mbx); 
    if not root: return None
    f=root
    for p in re.split(r"[\\/]+", path.strip("\\/")):
        try: f=f.Folders[p]
        except Exception: f=f.Folders.Add(p)
    return f
def categories_of(mail):
    return [c.strip().casefold() for c in (mail.Categories or "").split(",") if c]

# ─────── Predicate builders ───────
def _clean(v):
    if v is None: return None
    if isinstance(v,float) and pd.isna(v): return None
    if isinstance(v,str) and v.strip()=="": return None
    return str(v)
def eq_pred(v):
    v=_clean(v)
    if v is None: return lambda *_:True
    if v.startswith("/") and v.endswith("/"):
        pat=re.compile(v[1:-1],re.I); return lambda s,*_: bool(pat.search(s or ""))
    targets=[t.casefold() for t in v.split(",") if t.strip()]
    return lambda s,*_: (s or "").casefold() in targets
def kw_regex_pred(v):
    v=_clean(v)
    if v is None: return lambda *_:True
    if v.startswith("/") and v.endswith("/"):
        pat=re.compile(v[1:-1],re.I); return lambda s,*_: bool(pat.search(s or ""))
    kws=[k.casefold() for k in v.split(",") if k.strip()]
    return lambda s,*_: any(k in (s or "").casefold() for k in kws)

# ─────── Date helpers ─────────────
def naive(dt): return dt.replace(tzinfo=None) if dt and getattr(dt,"tzinfo",None) else dt
def outlook_date(dt:datetime)->str: return dt.strftime("%m/%d/%Y %I:%M %p")
def build_time_window(period,start=None,end=None):
    if period=="all": return None,None
    if period=="yesterday":
        today=datetime.now().date()
        start=datetime.combine(today-timedelta(days=1),datetime.min.time())
        end  =datetime.combine(today-timedelta(days=1),datetime.max.time())
    elif period=="thismonth":
        today=datetime.now().date()
        start=datetime(today.year,today.month,1)
        end=datetime.now()
    elif period=="custom":
        pass
    else: raise ValueError("bad period")
    return start,end

# ───────── Load rules ─────────────
def load_rules(xls):
    df=pd.read_excel(xls,sheet_name="Rules").where(pd.notnull, None)
    if "Enabled" in df: df=df[~df.Enabled.str.lower().eq("no")]
    for col in ("Mailbox","RuleName","ActionMoveTo"):
        if col not in df.columns: raise ValueError(f"Missing {col}")
    return df.reset_index(drop=True)

# ─── Core mailbox processing (fast) ───
def run_mb(df, src, period,start,end):
    root=open_root(src); inbox=root and root.Folders["Inbox"]
    if not inbox: return

    # COM-side date filter
    if period!="all":
        restriction = f"[ReceivedTime] >= '{outlook_date(start)}' AND [ReceivedTime] <= '{outlook_date(end)}'"
        items=inbox.Items.Restrict(restriction)
    else:
        items=inbox.Items
    # sort newest-first for speed
    items.Sort("[ReceivedTime]", True)

    rules=[dict(
        Rule=row.RuleName,
        Sender=eq_pred(row.SenderMatch),
        Sub=kw_regex_pred(row.SubjectContains),
        Cat=kw_regex_pred(row.Category),
        Target=row.ActionMoveTo) for _,row in df.iterrows()]

    itm=items.GetFirst()
    while itm:
        try:
            if itm.Class!=43:       # 43 = olMail
                itm=items.GetNext(); continue
            rt=getattr(itm,"ReceivedTime",None)
            if period!="all" and not (start<=naive(rt)<=end):
                itm=items.GetNext(); continue

            snd=itm.SenderName; subj=itm.Subject; cats=",".join(categories_of(itm))
            for R in rules:
                if R["Sender"](snd) and R["Sub"](subj) and R["Cat"](cats):
                    dest_desc=R["Target"] or f"Archive/{R['Rule']}"
                    parts=re.split(r"[\\/]+",dest_desc,1)
                    tgt, path=(src,parts[0]) if len(parts)==1 else parts
                    dest=ensure_folder(tgt,path)
                    if dest:
                        itm.Move(dest); itm.UnRead=False
                        logging.info(f"{src}|{R['Rule']}→{tgt}/{path}")
                    break
        except Exception as e:
            logging.error(f"Error mail in {src}: {e}")
        itm=items.GetNext()

# ───────── CLI / interactive ──────
def pick():
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
    p=argparse.ArgumentParser(); p.add_argument("-p","--period",choices=["yesterday","thismonth","all","custom"])
    p.add_argument("--start"); p.add_argument("--end"); a=p.parse_args()
    if a.period:
        per=a.period
        if per=="custom":
            if not (a.start and a.end): sys.exit("need --start --end")
            s=datetime.strptime(a.start,"%Y-%m-%d"); e=datetime.strptime(a.end,"%Y-%m-%d")+timedelta(days=1,seconds=-1)
        else: s=e=None
    else: per,s,e=pick()

    s,e = build_time_window(per,s,e)
    logging.info(f"=== Outlook Rule Runner period={per} ===")
    df=load_rules(RULEBOOK_PATH)
    for mbx,g in df.groupby("Mailbox"):
        run_mb(g,mbx,per,s,e)
    logging.info("=== done ===")

if __name__=="__main__": main()
