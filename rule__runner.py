"""
outlook_rule_runner.py
--------------------------------------------------------------
• SenderMatch   :  "name1;name2" (OR)  or  /regex/
• SubjectContains: "kw1;kw2"    (OR)  or  /regex/
• Category      :  "CatA;CatB"  (AND) or  /regex/
• ActionMoveTo  :  <TargetMailbox>/<Folder Path>
• Date windows  :  all | yesterday | thismonth | custom (CLI/menu)
• Fast scan     :  Outlook Items.Restrict + skips non-MailItem
--------------------------------------------------------------
Dependencies:
    pip install pandas pywin32 openpyxl
"""

import argparse
import logging
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import win32com.client as win32  # pywin32

# ─────────────────────────── CONFIG ───────────────────────────
RULEBOOK_PATH = Path(__file__).with_name("Outlook_Rule_Template.xlsx")
LOG_PATH = Path(__file__).with_suffix(".log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_PATH, "a", "utf-8"), logging.StreamHandler()],
)

# ───────────────────── Outlook helpers ────────────────────────
def ns():
    return win32.Dispatch("Outlook.Application").GetNamespace("MAPI")


def open_root(display_name: str):
    try:
        return ns().Folders[display_name]
    except Exception:
        logging.error(f"Mailbox '{display_name}' not found in Outlook profile.")
        return None


def ensure_folder(mailbox: str, path: str):
    root = open_root(mailbox)
    if not root:
        return None
    folder = root
    for part in re.split(r"[\\/]+", path.strip("\\/")):
        try:
            folder = folder.Folders[part]
        except Exception:
            folder = folder.Folders.Add(part)
    return folder


def categories_of(mail):
    return [c.strip().casefold() for c in (mail.Categories or "").split(",") if c]


# ───────────────────── Predicate builders ─────────────────────
def _clean(value):
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, str) and value.strip() == "":
        return None
    return str(value)


def eq_pred(value):
    """Match display name/e-mail: OR across semicolon list; supports regex."""
    value = _clean(value)
    if value is None:
        return lambda *_: True

    if value.startswith("/") and value.endswith("/"):
        pattern = re.compile(value[1:-1], re.IGNORECASE)
        return lambda s, *_: bool(pattern.search(s or ""))

    targets = [t.casefold() for t in re.split(r"\s*;\s*", value) if t.strip()]
    return lambda s, *_: (s or "").casefold() in targets


def kw_regex_pred(value):
    """Keyword OR logic for SubjectContains."""
    value = _clean(value)
    if value is None:
        return lambda *_: True

    if value.startswith("/") and value.endswith("/"):
        pattern = re.compile(value[1:-1], re.IGNORECASE)
        return lambda s, *_: bool(pattern.search(s or ""))

    keywords = [k.casefold() for k in re.split(r"\s*;\s*", value) if k.strip()]
    return lambda s, *_: any(k in (s or "").casefold() for k in keywords)


def category_pred(value):
    """AND logic for Categories (all tokens must be present)."""
    value = _clean(value)
    if value is None:
        return lambda *_: True

    if value.startswith("/") and value.endswith("/"):
        pattern = re.compile(value[1:-1], re.IGNORECASE)
        return lambda lst, *_: bool(pattern.search(",".join(lst)))

    tokens = [t.casefold() for t in re.split(r"\s*;\s*", value) if t.strip()]
    return lambda lst, *_: all(tok in lst for tok in tokens)


# ───────────────────── Date helpers ───────────────────────────
def naive(dt):
    return dt.replace(tzinfo=None) if dt and getattr(dt, "tzinfo", None) else dt


def ol_date(dt):
    """Outlook date string for Restrict filter."""
    return dt.strftime("%m/%d/%Y %I:%M %p")


def build_window(period, start=None, end=None):
    if period == "all":
        return None, None

    if period == "yesterday":
        today = datetime.now().date()
        start = datetime.combine(today - timedelta(days=1), datetime.min.time())
        end = datetime.combine(today - timedelta(days=1), datetime.max.time())
    elif period == "thismonth":
        today = datetime.now().date()
        start = datetime(today.year, today.month, 1)
        end = datetime.now()
    elif period == "custom":
        pass
    else:
        raise ValueError("Bad period")

    return start, end


# ───────────────────── Load rules ─────────────────────────────
def load_rules(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Rules").where(pd.notnull, None)
    if "Enabled" in df:
        df = df[~df.Enabled.str.lower().eq("no")]
    for col in ("Mailbox", "RuleName", "ActionMoveTo"):
        if col not in df.columns:
            raise ValueError(f"RuleBook missing column: {col}")
    return df.reset_index(drop=True)


# ───────────────────── Mailbox runner ────────────────────────
def run_mailbox(df: pd.DataFrame, src: str, period: str, start, end):
    root = open_root(src)
    inbox = root and root.Folders["Inbox"]
    if not inbox:
        return

    # Outlook-side Restrict for speed
    if period == "all":
        items = inbox.Items
    else:
        restriction = (
            f"[ReceivedTime] >= '{ol_date(start)}' AND [ReceivedTime] <= '{ol_date(end)}'"
        )
        items = inbox.Items.Restrict(restriction)

    items.Sort("[ReceivedTime]", True)

    rules = [
        dict(
            Name=row.RuleName,
            Sender=eq_pred(row.SenderMatch),
            Subject=kw_regex_pred(row.SubjectContains),
            Cat=category_pred(row.Category),
            Target=row.ActionMoveTo,
        )
        for _, row in df.iterrows()
    ]

    itm = items.GetFirst()
    while itm:
        try:
            if itm.Class != 43:  # not a MailItem
                itm = items.GetNext()
                continue

            sender = itm.SenderName
            subject = itm.Subject
            cats = categories_of(itm)

            for R in rules:
                if R["Sender"](sender) and R["Subject"](subject) and R["Cat"](cats):
                    dest_desc = R["Target"] or f"Archive/{R['Name']}"
                    parts = re.split(r"[\\/]+", dest_desc, 1)
                    tgt_mbx, path = (src, parts[0]) if len(parts) == 1 else parts
                    dest = ensure_folder(tgt_mbx, path)
                    if dest:
                        itm.Move(dest)
                        itm.UnRead = False
                        logging.info(f"{src} | {R['Name']} → {tgt_mbx}/{path}")
                    break
        except Exception as e:
            logging.error(f"Error processing mail in {src}: {e}")
        itm = items.GetNext()


# ───────────────────── CLI / Interactive ─────────────────────
def choose():
    mapping = {"1": "yesterday", "2": "thismonth", "3": "all", "4": "custom"}
    print("1 Yesterday   2 This month   3 All   4 Custom")
    choice = input("> ").strip()
    period = mapping.get(choice, "all")
    start = end = None
    if period == "custom":
        s = input("Start YYYY-MM-DD: ").strip()
        e = input("End   YYYY-MM-DD: ").strip()
        try:
            start = datetime.strptime(s, "%Y-%m-%d")
            end = datetime.strptime(e, "%Y-%m-%d") + timedelta(days=1, seconds=-1)
        except ValueError:
            print("Invalid date, defaulting to All")
            period = "all"
    return period, start, end


def main():
    ap = argparse.ArgumentParser(description="Outlook Rule Runner")
    ap.add_argument(
        "-p", "--period", choices=["yesterday", "thismonth", "all", "custom"]
    )
    ap.add_argument("--start")
    ap.add_argument("--end")
    args = ap.parse_args()

    if args.period:
        period = args.period
        if period == "custom":
            if not (args.start and args.end):
                sys.exit("Custom period requires --start and --end YYYY-MM-DD")
            start = datetime.strptime(args.start, "%Y-%m-%d")
            end = datetime.strptime(args.end, "%Y-%m-%d") + timedelta(
                days=1, seconds=-1
            )
        else:
            start = end = None
    else:
        period, start, end = choose()

    start, end = build_window(period, start, end)

    logging.info(f"=== Outlook Rule Runner period={period} ===")
    rules_df = load_rules(RULEBOOK_PATH)
    for mbx, grp in rules_df.groupby("Mailbox"):
        run_mailbox(grp, mbx, period, start, end)
    logging.info("=== Processing completed ===")


if __name__ == "__main__":
    main()