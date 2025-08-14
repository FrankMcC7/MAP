
"""
outlook_rule_runner.py  —  v1.3 (Aug 2025)

Fixes & Improvements
• Fix COM error (-2147467262, "No such interface supported") by not touching a moved item.
  - We cache the next pointer BEFORE moving and never access itm after Move().
  - We use the MOVE RETURN VALUE for post-move operations.
• SenderMatch now accepts display name OR email (and supports /regex/).
• Stronger logging (EntryID, subject, rule) and defensive error handling.
• Safer destination resolution and folder creation across mailboxes.
• Rule Category column uses AND logic (unchanged per your spec).
• Works with Inbox date filtering via Items.Restrict to keep it fast.
--------------------------------------------------------------
CLI:
  python outlook_rule_runner.py -p all
  python outlook_rule_runner.py -p yesterday
  python outlook_rule_runner.py -p custom --start 2025-07-01 --end 2025-07-11
"""

import argparse
import logging
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Iterable, List, Optional, Tuple

import pandas as pd
import win32com.client as win32
from pywintypes import com_error  # type: ignore

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
        try:
            # Try via shared default folder resolution (for shared mailboxes)
            rcpt = ns().CreateRecipient(display_name)
            rcpt.Resolve()
            store = ns().GetSharedDefaultFolder(rcpt, win32.constants.olFolderInbox).Parent
            return store
        except Exception:
            logging.error(f"Mailbox '{display_name}' not found/resolvable in Outlook profile.")
            return None


def ensure_folder(mailbox: str, path: str):
    """Create path under 'mailbox' if missing and return MAPIFolder."""
    root = open_root(mailbox)
    if not root:
        return None
    folder = root
    parts = [p for p in re.split(r"[\\/]+", path.strip("\\/")) if p]
    for part in parts:
        try:
            folder = folder.Folders[part]
        except Exception:
            try:
                folder = folder.Folders.Add(part)
            except Exception as e:
                logging.error(f"Failed to create/open subfolder '{part}' under {mailbox}: {e}")
                return None
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


def sender_pred(value) -> Callable[[str, str], bool]:
    """
    SenderMatch – OR logic over semicolons, or /regex/.
    Matches against BOTH display name and email.
    """
    value = _clean(value)
    if value is None:
        return lambda *_: True

    if value.startswith("/") and value.endswith("/"):
        pattern = re.compile(value[1:-1], re.IGNORECASE)
        return lambda name, email: bool(pattern.search(name or "") or pattern.search(email or ""))

    targets = [t.casefold() for t in re.split(r"\s*;\s*", value) if t.strip()]
    return lambda name, email: (name or "").casefold() in targets or (email or "").casefold() in targets


def kw_regex_pred(value) -> Callable[[str], bool]:
    """SubjectContains – OR logic over semicolons, or /regex/."""
    value = _clean(value)
    if value is None:
        return lambda *_: True

    if value.startswith("/") and value.endswith("/"):
        pattern = re.compile(value[1:-1], re.IGNORECASE)
        return lambda s: bool(pattern.search(s or ""))

    keywords = [k.casefold() for k in re.split(r"\s*;\s*", value) if k.strip()]
    return lambda s: any(k in (s or "").casefold() for k in keywords)


def category_pred(value) -> Callable[[Iterable[str]], bool]:
    """Category – AND logic over semicolons, or /regex/."""
    value = _clean(value)
    if value is None:
        return lambda *_: True

    if value.startswith("/") and value.endswith("/"):
        pattern = re.compile(value[1:-1], re.IGNORECASE)
        return lambda lst: bool(pattern.search(",".join(lst)))

    tokens = [t.casefold() for t in re.split(r"\s*;\s*", value) if t.strip()]
    return lambda lst: all(tok in lst for tok in tokens)


# ───────────────────── Date helpers ───────────────────────────
def ol_date(dt: datetime) -> str:
    return dt.strftime("%m/%d/%Y %I:%M %p")  # Outlook Restrict format


def build_window(period: str, start: Optional[datetime] = None, end: Optional[datetime] = None):
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
        # start/end already supplied
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
        logging.error(f"Could not open Inbox for mailbox: {src}")
        return

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
            Sender=sender_pred(row.SenderMatch),
            Subject=kw_regex_pred(row.SubjectContains),
            Cat=category_pred(row.Category),
            Target=row.ActionMoveTo,
        )
        for _, row in df.iterrows()
    ]

    itm = items.GetFirst()
    while itm:
        try:
            # Cache the next pointer FIRST. We will set itm=next_itm at loop end.
            next_itm = items.GetNext()

            # Skip non-mail items
            if getattr(itm, "Class", None) != 43:  # 43 = olMail
                itm = next_itm
                continue

            name = getattr(itm, "SenderName", "") or ""
            email = getattr(itm, "SenderEmailAddress", "") or ""
            subject = getattr(itm, "Subject", "") or ""
            cats = categories_of(itm)
            entry = getattr(itm, "EntryID", "") or ""

            matched = False
            for R in rules:
                if R["Sender"](name, email) and R["Subject"](subject) and R["Cat"](cats):
                    matched = True
                    dest_desc = R["Target"] or f"Archive/{R['Name']}"
                    parts = re.split(r"[\\/]+", dest_desc, 1)
                    tgt_mbx, path = (src, parts[0]) if len(parts) == 1 else parts
                    dest = ensure_folder(tgt_mbx, path)
                    if dest:
                        try:
                            moved = itm.Move(dest)  # returns the item in the destination
                            try:
                                moved.UnRead = False
                            except Exception:
                                pass
                            logging.info(f"{src} | {R['Name']} → {tgt_mbx}/{path} | {entry} | {subject}")
                        except com_error as ce:
                            logging.error(f"Move failed in {src} → {tgt_mbx}/{path} | {entry} | {subject} | {ce}")
                    else:
                        logging.error(f"Destination not available: {tgt_mbx}/{path} for rule '{R['Name']}'")
                    break  # stop at first matched rule

            if not matched:
                logging.info(f"{src} | NoRule | {entry} | {subject}")
        except com_error as e:
            logging.error(f"COM error in {src}: {e}")
        except Exception as e:
            logging.error(f"Error processing mail in {src}: {e}")
        finally:
            # Advance using the cached pointer to avoid touching moved item
            itm = next_itm

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
