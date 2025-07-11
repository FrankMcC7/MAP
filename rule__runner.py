"""
outlook_rule_runner.py  —  Outlook Rule Runner (enhanced)

• Reads an Excel “RuleBook” (sheet: Rules) that lists custom rules
  Columns (template already provided):
      Mailbox (required)
      RuleName (required)
      ActionMoveTo (required)
      Enabled           (yes / no, optional — default yes)
      SenderMatch       (optional)
      SubjectContains   (optional — comma keywords OR /regex/)
      Category          (optional — comma list)

• Behaviour checklist implemented:
      – Empty Subject/Category → no extra filter
      – Comma or /regex/ filters for SubjectContains & Category
      – Case-insensitive comparisons
      – Default archive path “Archive\<RuleName>” if ActionMoveTo blank
      – Recursive folder creation in any mailbox
      – Detailed logging of every move

• Date-window filters:
      --period  yesterday  → previous calendar day
      --period  thismonth  → from 1st of current month to now
      --period  all        → whole inbox
      --period  custom --start YYYY-MM-DD --end YYYY-MM-DD

• If run with NO arguments, an interactive menu appears.

• Designed to run head-less via Windows Task Scheduler, e.g.:
      python outlook_rule_runner.py --period yesterday

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

# ────────────────────────────────────────────────────────────
# CONFIGURATION
# ────────────────────────────────────────────────────────────
RULEBOOK_PATH = Path(__file__).with_name("Outlook_Rule_Template.xlsx")
LOG_PATH = Path(__file__).with_suffix(".log")
LOG_LEVEL = logging.INFO

logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_PATH, "a", "utf-8"), logging.StreamHandler()],
)

# ────────────────────────────────────────────────────────────
# Outlook helpers
# ────────────────────────────────────────────────────────────
def get_mapi_ns():
    return win32.Dispatch("Outlook.Application").GetNamespace("MAPI")


def open_mailbox_root(display_name: str):
    ns = get_mapi_ns()
    try:
        return ns.Folders[display_name]
    except Exception:
        logging.error(f"Mailbox '{display_name}' not found in Outlook profile.")
        return None


def ensure_folder(target_mbx: str, folder_path: str):
    """Create (if needed) and return Folder given <Mailbox>/<Path>."""
    root = open_mailbox_root(target_mbx)
    if not root:
        return None

    parts = re.split(r"[\\/]+", folder_path.strip("\\/"))
    folder = root
    for part in parts:
        try:
            folder = folder.Folders[part]
        except Exception:
            folder = folder.Folders.Add(part)
    return folder


def categories_of(mail):
    cats = mail.Categories
    if not cats:
        return []
    return [c.strip().casefold() for c in cats.split(",")]

# ────────────────────────────────────────────────────────────
# Predicate builders
# ────────────────────────────────────────────────────────────
def build_eq_pred(value: str):
    if not value or value.strip() == "":
        return lambda *_: True
    target = value.casefold()
    return lambda text, *_: (text or "").casefold() == target


def build_kw_or_regex_pred(value: str):
    if not value or value.strip() == "":
        return lambda *_: True

    s = value.strip()
    if s.startswith("/") and s.endswith("/"):
        pattern = re.compile(s[1:-1], re.IGNORECASE)
        return lambda text, *_: bool(pattern.search(text or ""))
    else:
        keywords = [k.casefold() for k in s.split(",") if k.strip()]
        return lambda text, *_: any(k in (text or "").casefold() for k in keywords)


def received_time_pred(period: str, start_dt=None, end_dt=None):
    """Return predicate comparing mail.ReceivedTime to date window."""
    if period == "all":
        return lambda *_: True

    if period == "yesterday":
        today = datetime.now().date()
        start = datetime.combine(today - timedelta(days=1), datetime.min.time())
        end = datetime.combine(today - timedelta(days=1), datetime.max.time())
    elif period == "thismonth":
        today = datetime.now().date()
        start = datetime(today.year, today.month, 1)
        end = datetime.now()
    elif period == "custom" and start_dt and end_dt:
        start, end = start_dt, end_dt
    else:
        raise ValueError("Invalid period definition")

    return lambda rt, *_: start <= rt <= end

# ────────────────────────────────────────────────────────────
# Core processing
# ────────────────────────────────────────────────────────────
def load_rules(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Rules")

    if "Enabled" in df.columns:
        df = df[~df["Enabled"].str.lower().eq("no")]

    required = {"Mailbox", "RuleName", "ActionMoveTo"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"RuleBook missing columns: {', '.join(missing)}")

    return df.reset_index(drop=True)


def process_mailbox(df: pd.DataFrame, src_mbx: str, time_pred):
    root = open_mailbox_root(src_mbx)
    if not root:
        return
    inbox = root.Folders["Inbox"]
    items = list(inbox.Items)  # enumerate once

    # Build rules
    rules = []
    for _, row in df.iterrows():
        rules.append(
            {
                "RuleName": row["RuleName"],
                "SenderPred": build_eq_pred(row.get("SenderMatch", "")),
                "SubjectPred": build_kw_or_regex_pred(row.get("SubjectContains", "")),
                "CategoryPred": build_kw_or_regex_pred(row.get("Category", "")),
                "Target": row["ActionMoveTo"],
            }
        )

    for mail in items:
        try:
            if not time_pred(mail.ReceivedTime):
                continue

            sender_name = getattr(mail, "SenderName", "")
            subject = getattr(mail, "Subject", "")
            cat_list = categories_of(mail)

            for r in rules:
                if (
                    r["SenderPred"](sender_name)
                    and r["SubjectPred"](subject)
                    and r["CategoryPred"](",".join(cat_list))
                ):
                    # Resolve destination
                    descriptor = r["Target"] or f"Archive/{r['RuleName']}"
                    parts = re.split(r"[\\/]+", descriptor, maxsplit=1)
                    if len(parts) == 1:
                        target_mbx, folder_path = src_mbx, parts[0]
                    else:
                        target_mbx, folder_path = parts[0], parts[1]

                    dest = ensure_folder(target_mbx, folder_path)
                    if dest:
                        moved = mail.Move(dest)
                        moved.UnRead = False
                        logging.info(
                            f"{src_mbx} | {r['RuleName']} | moved to {target_mbx}/{folder_path}"
                        )
                    else:
                        logging.error(
                            f"{src_mbx} | {r['RuleName']} | FAILED to resolve {descriptor}"
                        )
                    break
        except Exception as e:
            logging.error(f"Error processing mail in {src_mbx}: {e}")
            continue


def run(period: str, start_dt=None, end_dt=None):
    logging.info(f"=== Outlook Rule Runner – period: {period} ===")
    df = load_rules(RULEBOOK_PATH)
    time_filter = received_time_pred(period, start_dt, end_dt)

    for mbx, grp in df.groupby("Mailbox"):
        process_mailbox(grp, mbx, time_filter)

    logging.info("=== Processing completed ===")

# ────────────────────────────────────────────────────────────
# CLI & interactive menu
# ────────────────────────────────────────────────────────────
def interactive_prompt():
    print("Select processing window:")
    print(" 1. Yesterday")
    print(" 2. This month")
    print(" 3. Complete inbox")
    print(" 4. Custom range")
    choice = input("Enter choice [1-4]: ").strip()
    mapping = {"1": "yesterday", "2": "thismonth", "3": "all", "4": "custom"}
    period = mapping.get(choice, "all")

    if period == "custom":
        s = input("Start date (YYYY-MM-DD): ").strip()
        e = input("End   date (YYYY-MM-DD): ").strip()
        try:
            start_dt = datetime.strptime(s, "%Y-%m-%d")
            end_dt = datetime.strptime(e, "%Y-%m-%d") + timedelta(days=1, seconds=-1)
        except ValueError:
            print("Invalid date, defaulting to complete inbox.")
            period, start_dt, end_dt = "all", None, None
    else:
        start_dt = end_dt = None

    return period, start_dt, end_dt


def main():
    ap = argparse.ArgumentParser(description="Outlook Rule Runner")
    ap.add_argument(
        "-p",
        "--period",
        choices=["yesterday", "thismonth", "all", "custom"],
        help="Date window to process",
    )
    ap.add_argument("--start", help="Custom start date YYYY-MM-DD")
    ap.add_argument("--end", help="Custom end date YYYY-MM-DD")
    args = ap.parse_args()

    if args.period:
        if args.period == "custom":
            if not (args.start and args.end):
                print("Custom period requires --start and --end YYYY-MM-DD")
                sys.exit(1)
            s_dt = datetime.strptime(args.start, "%Y-%m-%d")
            e_dt = datetime.strptime(args.end, "%Y-%m-%d") + timedelta(days=1, seconds=-1)
            run("custom", s_dt, e_dt)
        else:
            run(args.period)
    else:
        per, s_dt, e_dt = interactive_prompt()
        run(per, s_dt, e_dt)


if __name__ == "__main__":
    main()
