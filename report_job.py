#!/usr/bin/env python3
"""
report_job.py - run several SQL Server stored procedures, put each result in its
own sheet of ONE Excel workbook, and email the workbook.

    SQL Server (stored procedures) -> DataFrames -> one .xlsx (a sheet per procedure) -> email

Usage
    python report_job.py                                   # previous full calendar month
    python report_job.py --from-date 2026-07-01 --to-date 2026-12-31
    python report_job.py --dry-run                         # build the file, do not send
    python report_job.py --only "Co cau tin dung"          # run only some sheets (repeatable)
    python report_job.py --due                             # run only the reports whose own schedule is due today
    python report_job.py --due --today 2024-01-01 --dry-run   # pretend today is 2024-01-01 (testing, back-filling)

What runs is described in reports.json (copy reports.example.json and edit it):
which procedures, in which sheet, with which parameters.
Connection details, mail server and recipients come from environment variables
or from a .env file next to this script (see .env.example).

Standard form: every procedure takes the same two parameters, @ngay_batdau and @ngay_ketthuc (start and end
date). A report in reports.json then needs no "params" at all; the program passes {from_date} and {to_date}.
Extra fixed values go in "extra_params"; "params" is only for procedures that do not follow the standard.

Each report can have its own schedule ("schedule" in reports.json). One daily trigger (Task Scheduler / cron)
runs "python report_job.py --due"; the program then runs only the reports that are due that day, each with the
period that fits its schedule (previous week, previous month, previous quarter...). See README.md.

Parameter tokens usable in reports.json (all derived from the reporting period)
    "{from_date}"      start of the period (date)
    "{to_date}"        end of the period (date)
    "{year}"           year of {to_date} (int)
    "{year_start}"     1 January of that year       "{year_end}"   31 December of that year
    "{prev_year_end}"  31 December of the year before
    "{day_before_from}" the day before {from_date} = last day of the previous month/quarter (opening snapshot)
    "{month_start}"    first day of {to_date}'s month    "{month_end}"  last day of that month
    "{today}"          today's date
A report can have its own period: "period": {"from": "2023-07-01", "to": "2023-12-31"}
(each value is a YYYY-MM-DD date or one of the date tokens above); the tokens then refer to that period.
Any other parameter value (a number, a text, a YYYY-MM-DD date) is passed as written.
"""
from __future__ import annotations

import argparse
import datetime as dt
import decimal
import json
import logging
import os
import re
import smtplib
import ssl
import sys
import traceback
import unicodedata
from dataclasses import dataclass
from email.message import EmailMessage
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableStyleInfo

try:  # optional: load variables from a .env file next to this script
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass

BASE_DIR = Path(__file__).resolve().parent
log = logging.getLogger("report_job")

INFO_SHEET = "Thong tin"
FONT_NAME = "Arial"
XLSX_SUBTYPE = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"

PROC_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*){0,2}$")
PARAM_RE = re.compile(r"^@[A-Za-z_][A-Za-z0-9_]*$")
TOKEN_RE = re.compile(r"^\{(\w+)\}$")
BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
WEEKDAYS = {"monday": 1, "tuesday": 2, "wednesday": 3, "thursday": 4, "friday": 5, "saturday": 6, "sunday": 7}
FREQUENCIES = ("daily", "weekly", "monthly", "quarterly", "yearly")
PERIOD_RULES = ("yesterday", "previous_week", "previous_month", "previous_quarter", "previous_year",
                "month_to_date", "year_to_date", "year_over_year")
DEFAULT_RULE = {"daily": "yesterday", "weekly": "previous_week", "monthly": "previous_month",
                "quarterly": "previous_quarter", "yearly": "previous_year"}
STANDARD_FROM = "@ngay_batdau"
STANDARD_TO = "@ngay_ketthuc"
ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
TABLE_STYLE_RE = re.compile(r"^TableStyle(Light|Medium|Dark)\d{1,2}$")


class ReportError(Exception):
    """Expected failure (empty result, bad config, file too large...)."""


# --------------------------------------------------------------------------
# Configuration (.env / environment)
# --------------------------------------------------------------------------
def _env(name: str, default: str | None = None, required: bool = False) -> str | None:
    """Read a setting. A missing OR empty value (e.g. "OUTPUT_DIR=" in .env) falls back to the default."""
    value = os.environ.get(name)
    value = value.strip() if value is not None else None
    if not value:
        value = default
    if required and not value:
        raise ReportError(f"Missing required setting: {name}")
    return value or None


def _split(value: str | None) -> list[str]:
    return [p.strip() for p in (value or "").replace(";", ",").split(",") if p.strip()]


def _flag(name: str, default: bool) -> bool:
    raw = _env(name)
    if raw is None:
        return default
    return raw.lower() in {"1", "true", "yes", "y"}


@dataclass
class Config:
    sql_server: str
    sql_database: str
    sql_user: str | None
    sql_password: str | None
    sql_trusted: bool
    sql_driver: str
    sql_encrypt: str
    sql_trust_cert: str
    sql_query_timeout: int
    smtp_host: str
    smtp_port: int
    smtp_starttls: bool
    smtp_user: str | None
    smtp_password: str | None
    mail_from: str
    mail_to: list[str]
    mail_cc: list[str]
    alert_to: list[str]
    output_dir: Path
    reports_file: Path
    allow_empty: bool
    max_attach_mb: int

    @classmethod
    def from_env(cls) -> "Config":
        trusted = _flag("SQL_TRUSTED", False)
        user = _env("SQL_USER")
        password = _env("SQL_PASSWORD")
        if not trusted and not (user and password):
            raise ReportError("Set SQL_USER and SQL_PASSWORD, or SQL_TRUSTED=true (Windows login)")
        return cls(
            sql_server=_env("SQL_SERVER", required=True),
            sql_database=_env("SQL_DATABASE", required=True),
            sql_user=user,
            sql_password=password,
            sql_trusted=trusted,
            sql_driver=_env("SQL_DRIVER", "ODBC Driver 18 for SQL Server"),
            sql_encrypt=_env("SQL_ENCRYPT", "yes"),
            sql_trust_cert=_env("SQL_TRUST_CERT", "yes"),
            sql_query_timeout=int(_env("SQL_QUERY_TIMEOUT", "300")),
            smtp_host=_env("SMTP_HOST", required=True),
            smtp_port=int(_env("SMTP_PORT", "587")),
            smtp_starttls=_flag("SMTP_STARTTLS", True),
            smtp_user=_env("SMTP_USER"),
            smtp_password=_env("SMTP_PASSWORD"),
            mail_from=_env("MAIL_FROM", required=True),
            mail_to=_split(_env("MAIL_TO", required=True)),
            mail_cc=_split(_env("MAIL_CC")),
            alert_to=_split(_env("ALERT_TO")),
            output_dir=Path(_env("OUTPUT_DIR", str(BASE_DIR / "output"))),
            reports_file=Path(_env("REPORTS_FILE", str(BASE_DIR / "reports.json"))),
            allow_empty=_flag("ALLOW_EMPTY", False),
            max_attach_mb=int(_env("MAX_ATTACH_MB", "20")),
        )


def setup_logging() -> None:
    log_dir = BASE_DIR / "logs"
    log_dir.mkdir(exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    file_handler = RotatingFileHandler(
        log_dir / "report_job.log", maxBytes=1_000_000, backupCount=5, encoding="utf-8"
    )
    file_handler.setFormatter(fmt)
    console = logging.StreamHandler()
    console.setFormatter(fmt)
    log.setLevel(logging.INFO)
    log.handlers.clear()
    log.addHandler(file_handler)
    log.addHandler(console)


# --------------------------------------------------------------------------
# Reports definition (reports.json)
# --------------------------------------------------------------------------
@dataclass
class Job:
    sheet: str
    proc: str
    params: Any = None  # dict (named, keys like "@year") or list (positional); filled with the standard dates if omitted
    title: str | None = None  # title shown at the top of the sheet (default: the sheet name)
    columns: list[str] | None = None  # optional: rename the columns by position
    table: bool = True  # format the result as an Excel table (filter buttons, banded rows)
    table_style: str = "TableStyleMedium2"
    hierarchy: bool | None = None  # True/False forces the "level 1 rows in bold" layout; None = auto
    period: dict | None = None  # optional own period: {"from": ..., "to": ...}
    period_label: str | None = None  # filled in while running, shown in the sheet subtitle
    schedule: dict | None = None  # when the report is due (used by --due)
    period_rule: str | None = None  # which period a scheduled run covers (default depends on the schedule)
    group: str | None = None  # reports of the same group that are due together share one workbook and one email
    mail_to: list[str] | None = None  # recipients of this report (default: MAIL_TO)
    base_ctx: dict | None = None  # filled in while running a scheduled report
    name_template: str | None = None  # name of the group's file and email, with {period}; from "groups" in reports.json


def load_jobs(path: Path) -> list[Job]:
    if not path.exists():
        raise ReportError(f"Reports file not found: {path} (copy reports.example.json to reports.json and edit it)")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ReportError(f"{path.name} is not valid JSON: {exc}") from exc

    items = data.get("reports") if isinstance(data, dict) else None
    if not items or not isinstance(items, list):
        raise ReportError(f'{path.name} must contain a non-empty "reports" list')

    standard = data.get("standard_params", {}) if isinstance(data, dict) else {}
    from_name = standard.get("from", STANDARD_FROM) if isinstance(standard, dict) else STANDARD_FROM
    to_name = standard.get("to", STANDARD_TO) if isinstance(standard, dict) else STANDARD_TO
    for name in (from_name, to_name):
        if not (isinstance(name, str) and PARAM_RE.match(name)):
            raise ReportError(f'{path.name}: "standard_params" must give parameter names like "@ngay_batdau"')

    groups_cfg = data.get("groups", {}) if isinstance(data, dict) else {}
    if not isinstance(groups_cfg, dict):
        raise ReportError(f'{path.name}: "groups" must be an object')
    for group_name, group_cfg in groups_cfg.items():
        template = group_cfg.get("name") if isinstance(group_cfg, dict) else None
        if not isinstance(group_cfg, dict) or set(group_cfg) - {"name"} or not isinstance(template, str):
            raise ReportError(f'{path.name}: group "{group_name}" must look like {{"name": "Báo cáo ... {{period}}"}}')
        leftover = template.replace("{period}", "")
        if "{period}" not in template or "{" in leftover or "}" in leftover:
            raise ReportError(f'{path.name}: group "{group_name}": "name" must contain {{period}} and no other braces')

    jobs: list[Job] = []
    seen: set[str] = set()
    for number, item in enumerate(items, start=1):
        if not isinstance(item, dict):
            raise ReportError(f"{path.name}: entry #{number} must be an object")
        sheet = str(item.get("sheet", "")).strip()
        proc = str(item.get("proc", "")).strip()
        params = item.get("params")
        extra = item.get("extra_params")
        if params is None:  # standard form: the two dates, plus any extra fixed values
            params = {from_name: "{from_date}", to_name: "{to_date}"}
            if extra is not None:
                if not isinstance(extra, dict):
                    raise ReportError(f'{path.name}: entry #{number} "extra_params" must be an object')
                params.update(extra)
        elif extra is not None:
            raise ReportError(f'{path.name}: entry #{number} has both "params" and "extra_params"; use only one')

        if not sheet or len(sheet) > 31 or BAD_SHEET_CHARS.search(sheet) or sheet.lower() == INFO_SHEET.lower():
            raise ReportError(
                f'{path.name}: entry #{number} has an invalid sheet name "{sheet}" '
                f'(max 31 characters, none of [ ] : * ? / \\, and not "{INFO_SHEET}")'
            )
        if sheet.lower() in seen:
            raise ReportError(f'{path.name}: sheet name "{sheet}" is used twice')
        seen.add(sheet.lower())
        if not PROC_RE.match(proc):
            raise ReportError(f'{path.name}: entry #{number} has an invalid procedure name "{proc}"')
        if isinstance(params, dict):
            bad = [k for k in params if not PARAM_RE.match(str(k))]
            if bad:
                raise ReportError(f'{path.name}: entry #{number} has invalid parameter names {bad} (use "@name")')
        elif not isinstance(params, list):
            raise ReportError(f'{path.name}: entry #{number} "params" must be a list or an object')

        title = item.get("title")
        columns = item.get("columns")
        table = item.get("table", True)
        table_style = item.get("table_style", "TableStyleMedium2")
        hierarchy = item.get("hierarchy")
        where = f"{path.name}: entry #{number}"
        if title is not None and not isinstance(title, str):
            raise ReportError(f'{where} "title" must be text')
        if columns is not None and not (isinstance(columns, list) and all(isinstance(c, str) for c in columns)):
            raise ReportError(f'{where} "columns" must be a list of text values')
        if not isinstance(table, bool):
            raise ReportError(f'{where} "table" must be true or false')
        if not (isinstance(table_style, str) and TABLE_STYLE_RE.match(table_style)):
            raise ReportError(f'{where} "table_style" must look like TableStyleMedium2')
        if hierarchy is not None and not isinstance(hierarchy, bool):
            raise ReportError(f'{where} "hierarchy" must be true or false')
        period = item.get("period")
        if period is not None and not (
            isinstance(period, dict) and period and set(period) <= {"from", "to"}
            and all(isinstance(v, str) for v in period.values())
        ):
            raise ReportError(f'{where} "period" must be an object with "from" and/or "to" (text dates)')
        schedule = item.get("schedule")
        if schedule is not None:
            schedule = normalize_schedule(schedule, where)
        period_rule = item.get("period_rule")
        if period_rule is not None and period_rule not in PERIOD_RULES:
            raise ReportError(f'{where} "period_rule" must be one of: {", ".join(PERIOD_RULES)}')
        group = item.get("group")
        if group is not None and not (isinstance(group, str) and group.strip()):
            raise ReportError(f'{where} "group" must be a non-empty text')
        mail_to = item.get("mail_to")
        if mail_to is not None and not (
            isinstance(mail_to, list) and mail_to and all(isinstance(a, str) and "@" in a for a in mail_to)
        ):
            raise ReportError(f'{where} "mail_to" must be a list of email addresses')
        jobs.append(Job(sheet=sheet, proc=proc, params=params, title=title, columns=columns,
                        table=table, table_style=table_style, hierarchy=hierarchy, period=period, schedule=schedule,
                        period_rule=period_rule, group=group.strip() if group else None,
                        name_template=(groups_cfg.get(group.strip() if group else sheet) or {}).get("name"),
                        mail_to=[a.strip() for a in mail_to] if mail_to else None))
    unused = sorted(set(groups_cfg) - {job.group or job.sheet for job in jobs})
    if unused:
        raise ReportError(f'{path.name}: "groups" names {unused} do not match any report group (check the spelling)')
    return jobs


def default_period(today: dt.date | None = None) -> tuple[dt.date, dt.date]:
    """Default = the previous full calendar month. Change it for weekly reports."""
    today = today or dt.date.today()
    first_this_month = today.replace(day=1)
    last_prev_month = first_this_month - dt.timedelta(days=1)
    return last_prev_month.replace(day=1), last_prev_month


def _last_day(year: int, month: int) -> int:
    return ((dt.date(year, month, 28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)).day


def normalize_schedule(schedule: Any, where: str) -> dict[str, Any]:
    """Validate a schedule and fill in the defaults."""
    if not isinstance(schedule, dict) or schedule.get("frequency") not in FREQUENCIES:
        raise ReportError(f'{where} "schedule" needs a "frequency": {", ".join(FREQUENCIES)}')
    freq = schedule["frequency"]
    allowed = {"frequency"} | {"weekly": {"weekday"}, "monthly": {"day"}, "quarterly": {"day"},
                               "yearly": {"month", "day"}, "daily": set()}[freq]
    extra = set(schedule) - allowed
    if extra:
        raise ReportError(f'{where} "schedule" ({freq}) does not use {sorted(extra)}; allowed: {sorted(allowed)}')

    result: dict[str, Any] = {"frequency": freq}
    if freq == "weekly":
        raw = schedule.get("weekday", "monday")
        day = WEEKDAYS.get(str(raw).lower()) if isinstance(raw, str) else raw
        if not (isinstance(day, int) and not isinstance(day, bool) and 1 <= day <= 7):
            raise ReportError(f'{where} "weekday" must be monday..sunday (or 1..7, Monday = 1)')
        result["weekday"] = day
    if freq == "yearly":
        month = schedule.get("month", 1)
        if not (isinstance(month, int) and not isinstance(month, bool) and 1 <= month <= 12):
            raise ReportError(f'{where} "month" must be a number from 1 to 12')
        result["month"] = month
    if freq in ("monthly", "quarterly", "yearly"):
        day = schedule.get("day", 1)
        if day != "last" and not (isinstance(day, int) and not isinstance(day, bool) and 1 <= day <= 31):
            raise ReportError(f'{where} "day" must be 1..31 or "last"')
        result["day"] = day
    return result


def is_due(schedule: dict[str, Any], today: dt.date) -> bool:
    freq = schedule["frequency"]
    if freq == "daily":
        return True
    if freq == "weekly":
        return today.isoweekday() == schedule["weekday"]
    if freq == "quarterly" and today.month not in (1, 4, 7, 10):
        return False
    if freq == "yearly" and today.month != schedule["month"]:
        return False
    last = _last_day(today.year, today.month)
    day = last if schedule["day"] == "last" else min(schedule["day"], last)  # day 31 in February = last day
    return today.day == day


def rule_period(rule: str, today: dt.date) -> tuple[dt.date, dt.date]:
    """The period a scheduled run covers, relative to the day it runs."""
    if rule == "yesterday":
        day = today - dt.timedelta(days=1)
        return day, day
    if rule == "previous_week":  # Monday..Sunday of the week before
        monday = today - dt.timedelta(days=today.weekday()) - dt.timedelta(days=7)
        return monday, monday + dt.timedelta(days=6)
    if rule == "previous_month":
        end = today.replace(day=1) - dt.timedelta(days=1)
        return end.replace(day=1), end
    if rule == "previous_quarter":
        this_quarter = dt.date(today.year, 3 * ((today.month - 1) // 3) + 1, 1)
        end = this_quarter - dt.timedelta(days=1)
        return dt.date(end.year, 3 * ((end.month - 1) // 3) + 1, 1), end
    if rule == "previous_year":
        return dt.date(today.year - 1, 1, 1), dt.date(today.year - 1, 12, 31)
    if rule == "month_to_date":
        return today.replace(day=1), today
    if rule == "year_to_date":
        return dt.date(today.year, 1, 1), today
    if rule == "year_over_year":  # two points in time: end of last year versus the end of the year before
        return dt.date(today.year - 2, 12, 31), dt.date(today.year - 1, 12, 31)
    raise ReportError(f"Unknown period rule: {rule}")


def job_period_rule(job: Job) -> str:
    return job.period_rule or DEFAULT_RULE[job.schedule["frequency"]]


def _slug(text: str) -> str:
    ascii_text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", ascii_text.lower()).strip("_") or "bao_cao"


ROMAN_QUARTER = {1: "I", 2: "II", 3: "III", 4: "IV"}


def period_parts(from_date: dt.date, to_date: dt.date) -> tuple[str, str] | None:
    """(kind, text) for a calendar month, quarter or year: ("month", "tháng 01/2024"); None for any other period."""
    if from_date.day == 1 and (from_date.year, from_date.month) == (to_date.year, to_date.month) \
            and to_date.day == _last_day(to_date.year, to_date.month):
        return "month", f"tháng {from_date:%m/%Y}"
    if from_date.day == 1 and from_date.month in (1, 4, 7, 10) and to_date.year == from_date.year \
            and to_date.month == from_date.month + 2 and to_date.day == _last_day(to_date.year, to_date.month):
        return "quarter", f"quý {ROMAN_QUARTER[(from_date.month - 1) // 3 + 1]}/{from_date.year}"
    if (from_date.month, from_date.day, to_date.month, to_date.day) == (1, 1, 12, 31) and from_date.year == to_date.year:
        return "year", f"năm {from_date.year}"
    return None


def group_title(name: str, items: list[Job], ctx: dict[str, Any], today: dt.date) -> str:
    """Name of a group's file and email, from the period its reports cover.

    Uses the group's "name" from reports.json (with {period} replaced by "tháng 01/2024", "quý IV/2024",
    "năm 2024"); without one: "Báo cáo tháng 01/2024", "Báo cáo quý IV/2024", "Báo cáo tổng kết năm 2024".
    """
    parts = set()
    for job in items:
        own = job_context(job, ctx)
        parts.add(period_parts(own["from_date"], own["to_date"]))
    if len(parts) != 1 or None in parts:  # reports cover different periods, or not a whole month/quarter/year
        return f"Báo cáo {name} ({today:%d/%m/%Y})"
    kind, text = parts.pop()
    template = next((job.name_template for job in items if job.name_template), None)
    if template:
        return template.replace("{period}", text)
    return f"Báo cáo tổng kết {text}" if kind == "year" else f"Báo cáo {text}"


def safe_filename(text: str) -> str:
    """Keep the wording (including Vietnamese letters) but remove what Windows does not allow in a file name."""
    cleaned = re.sub(r'[<>:"/\\|?*]', "-", text)
    return re.sub(r"\s+", " ", cleaned).strip(" .") or "bao_cao"


def make_context(from_date: dt.date, to_date: dt.date, today: dt.date | None = None) -> dict[str, Any]:
    """All tokens available to the parameters, derived from one period."""
    today = today or dt.date.today()
    next_month = (to_date.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
    return {
        "from_date": from_date,
        "day_before_from": from_date - dt.timedelta(days=1),  # last day of the previous period
        "to_date": to_date,
        "year": to_date.year,
        "today": today,
        "year_start": dt.date(to_date.year, 1, 1),
        "year_end": dt.date(to_date.year, 12, 31),
        "prev_year_end": dt.date(to_date.year - 1, 12, 31),
        "month_start": to_date.replace(day=1),
        "month_end": next_month - dt.timedelta(days=1),
    }


def _parse_period_value(value: str, ctx: dict[str, Any], which: str) -> dt.date:
    value = value.strip()
    match = TOKEN_RE.match(value)
    if match:
        found = ctx.get(match.group(1))
        if not isinstance(found, dt.date):
            raise ReportError(f'period "{which}": "{value}" is not a date token')
        return found
    try:
        return dt.date.fromisoformat(value)
    except ValueError as exc:
        raise ReportError(
            f'period "{which}": invalid date "{value}" (use YYYY-MM-DD or a token such as {{year_start}})'
        ) from exc


def job_context(job: Job, ctx: dict[str, Any]) -> dict[str, Any]:
    """The tokens for one report: its own period if it has one, otherwise the global period."""
    ctx = job.base_ctx or ctx
    if not job.period:
        return ctx
    from_date = _parse_period_value(job.period["from"], ctx, "from") if "from" in job.period else ctx["from_date"]
    to_date = _parse_period_value(job.period["to"], ctx, "to") if "to" in job.period else ctx["to_date"]
    if from_date > to_date:
        raise ReportError(f"Sheet '{job.sheet}': period starts ({from_date}) after it ends ({to_date})")
    return make_context(from_date, to_date, ctx["today"])


def _resolve_value(value: Any, ctx: dict[str, Any]) -> Any:
    if isinstance(value, str):
        match = TOKEN_RE.match(value)
        if match:
            key = match.group(1)
            if key not in ctx:
                raise ReportError(f'Unknown parameter token "{value}" (available: {", ".join("{" + k + "}" for k in ctx)})')
            return ctx[key]
        if ISO_DATE_RE.match(value):  # a literal date such as "2023-06-30" is sent as a real date
            try:
                return dt.date.fromisoformat(value)
            except ValueError:
                return value
    return value


def build_call(job: Job, ctx: dict[str, Any]) -> tuple[str, list[Any], str]:
    """Return (EXEC statement with ? placeholders, values, readable description)."""
    ctx = job_context(job, ctx)
    if isinstance(job.params, dict):
        pairs = [(str(k), _resolve_value(v, ctx)) for k, v in job.params.items()]
        placeholders = ", ".join(f"{k} = ?" for k, _ in pairs)
        values = [v for _, v in pairs]
        text = ", ".join(f"{k}={_fmt(v)}" for k, v in pairs)
    else:
        values = [_resolve_value(v, ctx) for v in job.params]
        placeholders = ", ".join("?" for _ in values)
        text = ", ".join(_fmt(v) for v in values)
    sql = f"EXEC {job.proc}" + (f" {placeholders}" if placeholders else "")
    return sql, values, text


def _fmt(value: Any) -> str:
    return value.isoformat() if isinstance(value, (dt.date, dt.datetime)) else str(value)


# --------------------------------------------------------------------------
# 1) SQL Server -> DataFrames
# --------------------------------------------------------------------------
def _odbc_value(value: str) -> str:
    return "{" + value.replace("}", "}}") + "}"


def fetch_results(cfg: Config, jobs: list[Job], ctx: dict[str, Any]) -> list[tuple[Job, pd.DataFrame, str]]:
    import pyodbc  # imported here so the other functions can be used/tested without it

    parts = [
        f"DRIVER={_odbc_value(cfg.sql_driver)}",
        f"SERVER={cfg.sql_server}",
        f"DATABASE={cfg.sql_database}",
        f"Encrypt={cfg.sql_encrypt}",
        f"TrustServerCertificate={cfg.sql_trust_cert}",
        "Connection Timeout=30",
    ]
    if cfg.sql_trusted:
        parts.append("Trusted_Connection=yes")
    else:
        parts.append(f"UID={_odbc_value(cfg.sql_user)}")
        parts.append(f"PWD={_odbc_value(cfg.sql_password)}")

    results: list[tuple[Job, pd.DataFrame, str]] = []
    conn = pyodbc.connect(";".join(parts))
    try:
        conn.timeout = cfg.sql_query_timeout  # query timeout in seconds
        for job in jobs:
            sql, values, text = build_call(job, ctx)
            if job.period or job.base_ctx:
                own = job_context(job, ctx)
                job.period_label = f"{own['from_date']:%d/%m/%Y} - {own['to_date']:%d/%m/%Y}"
            log.info("Running %s (%s)", sql, text)
            try:
                cur = conn.cursor()
                cur.execute(sql, values) if values else cur.execute(sql)
                # Skip "rows affected" messages: go to the first result set that has columns.
                while cur.description is None:
                    if not cur.nextset():
                        raise ReportError("the procedure returned no result set")
                columns = [c[0] for c in cur.description]
                rows = [tuple(r) for r in cur.fetchall()]
            except ReportError as exc:
                raise ReportError(f"Sheet '{job.sheet}' ({job.proc}): {exc}") from exc
            except Exception as exc:  # noqa: BLE001
                raise ReportError(f"Sheet '{job.sheet}' ({job.proc}) failed: {exc}") from exc

            df = pd.DataFrame.from_records(rows, columns=columns)
            log.info("  -> %d rows x %d columns", len(df), len(df.columns))
            if df.empty and not cfg.allow_empty:
                raise ReportError(f"Sheet '{job.sheet}' ({job.proc}) returned no rows; nothing was sent")
            results.append((job, df, text))
    finally:
        conn.close()
    return results


# --------------------------------------------------------------------------
# 2) DataFrames -> one Excel workbook
# --------------------------------------------------------------------------
HEADER_FILL = "1F3864"  # dark blue
LEVEL1_FILL = "DDEBF7"  # light blue for level-1 rows of hierarchical reports


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Turn Decimal columns into floats so Excel receives numbers, not text."""
    if df.shape[1] == 0:
        return df
    columns = []
    for i in range(df.shape[1]):
        series = df.iloc[:, i]
        if series.dtype == object:
            non_null = series.dropna()
            if len(non_null) and non_null.map(lambda v: isinstance(v, decimal.Decimal)).all():
                series = series.map(lambda v: float(v) if isinstance(v, decimal.Decimal) else v).astype(float)
        columns.append(series)
    cleaned = pd.concat(columns, axis=1)
    cleaned.columns = df.columns
    return cleaned


def _number_format(series: pd.Series, header: str = "") -> str | None:
    non_null = series.dropna()
    if non_null.empty:
        return None
    first = non_null.iloc[0]
    if isinstance(first, bool):
        return None
    if isinstance(first, (dt.datetime, pd.Timestamp)):
        has_time = any(
            (v.hour or v.minute or v.second) for v in non_null if isinstance(v, (dt.datetime, pd.Timestamp))
        )
        return "dd/mm/yyyy hh:mm" if has_time else "dd/mm/yyyy"
    if isinstance(first, dt.date):
        return "dd/mm/yyyy"
    if pd.api.types.is_numeric_dtype(series):
        if "%" in header:
            return '0.00"%"'  # the value is already a percentage (12.5 means 12.5%)
        return "#,##0" if (non_null % 1 == 0).all() else "#,##0.00"
    return None


def _prepare_headers(names: list, override: list[str] | None, sheet: str) -> list[str]:
    """Return unique, non-empty column names (Excel tables require them)."""
    names = ["" if n is None else str(n).strip() for n in names]
    if override is not None:
        if len(override) != len(names):
            raise ReportError(
                f"Sheet '{sheet}': \"columns\" has {len(override)} names but the procedure returned {len(names)} columns"
            )
        names = [str(n).strip() for n in override]
    result: list[str] = []
    seen: dict[str, int] = {}
    for i, name in enumerate(names, start=1):
        base = name or f"Cột {i}"
        count = seen.get(base.lower(), 0)
        seen[base.lower()] = count + 1
        result.append(base if count == 0 else f"{base} {count + 1}")
    return result


def _style_sheet(ws, df: pd.DataFrame, header_row: int, title: str, subtitle: str, table_name: str,
                 use_table: bool, table_style: str, hierarchy: bool | None) -> None:
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT_NAME, size=14, bold=True, color=HEADER_FILL)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name=FONT_NAME, size=9, italic=True, color="595959")

    n_cols = df.shape[1]
    if n_cols == 0:
        return
    first_row = header_row + 1
    last_row = header_row + len(df)
    if hierarchy is None:
        hierarchy = str(df.columns[0]).lower().startswith("tiêu chí cấp 1")

    header_font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color=HEADER_FILL, end_color=HEADER_FILL)
    for cell in ws[header_row]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 30

    body_font = Font(name=FONT_NAME, size=10)
    bold_font = Font(name=FONT_NAME, size=10, bold=True)
    level1_fill = PatternFill("solid", start_color=LEVEL1_FILL, end_color=LEVEL1_FILL)

    for i in range(n_cols):
        header = str(df.columns[i])
        fmt = _number_format(df.iloc[:, i], header)
        width = len(header)
        for row in range(first_row, last_row + 1):
            cell = ws.cell(row=row, column=i + 1)
            cell.font = body_font
            if cell.value is None or cell.value == "":
                continue
            if fmt:
                cell.number_format = fmt
            width = max(width, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(i + 1)].width = min(max(width + 2, 10), 50)

    if not hierarchy:  # light zebra stripes on flat tables
        stripe_fill = PatternFill("solid", start_color="EEF3FA", end_color="EEF3FA")
        for offset, row in enumerate(range(first_row, last_row + 1)):
            if offset % 2 == 1:
                for col in range(1, n_cols + 1):
                    ws.cell(row=row, column=col).fill = stripe_fill

    if hierarchy:
        for row in range(first_row, last_row + 1):
            if ws.cell(row=row, column=1).value not in (None, ""):  # a level-1 row: bold + light fill
                for col in range(1, n_cols + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.font = bold_font
                    cell.fill = level1_fill
            else:  # indent the detail rows
                for col in (2, 3):
                    if col <= n_cols and ws.cell(row=row, column=col).value not in (None, ""):
                        ws.cell(row=row, column=col).alignment = Alignment(indent=col - 1)

    ref = f"A{header_row}:{get_column_letter(n_cols)}{max(last_row, first_row)}"
    if use_table and len(df) > 0:
        tbl = Table(displayName=table_name, ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name=table_style, showFirstColumn=False, showLastColumn=False,
            showRowStripes=False, showColumnStripes=False,
        )
        ws.add_table(tbl)  # the table brings its own filter buttons
    else:
        ws.auto_filter.ref = ref

    ws.freeze_panes = f"A{first_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = f"{header_row}:{header_row}"


def build_workbook(results: list[tuple[Job, pd.DataFrame, str]], path: Path, label: str | None = None) -> Path:
    """label = the period shown on the sheets when a report has no period of its own (manual runs)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    created = dt.datetime.now().strftime("%d/%m/%Y %H:%M")

    def subtitle_for(period: str | None) -> str:
        return f"Kỳ báo cáo: {period}   |   Tạo lúc: {created}" if period else f"Tạo lúc: {created}"

    periods = {job.period_label or label for job, _, _ in results}
    subtitle = subtitle_for(periods.pop() if len(periods) == 1 else None)
    info = pd.DataFrame(
        [
            {"Sheet": job.sheet, "Procedure": job.proc, "Kỳ báo cáo": job.period_label or label or "",
             "Tham số": text, "Số dòng": len(df)}
            for job, df, text in results
        ]
    )
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        info.to_excel(writer, sheet_name=INFO_SHEET, index=False, startrow=3)
        _style_sheet(writer.sheets[INFO_SHEET], info, 4, "Thông tin báo cáo", subtitle, "tbl_thong_tin",
                     True, "TableStyleMedium2", False)

        for idx, (job, df, _) in enumerate(results, start=1):
            clean = _clean_dataframe(df)
            clean.columns = _prepare_headers(list(clean.columns), job.columns, job.sheet)
            clean.to_excel(writer, sheet_name=job.sheet, index=False, startrow=3)
            sheet_subtitle = subtitle_for(job.period_label or label)
            _style_sheet(writer.sheets[job.sheet], clean, 4, job.title or job.sheet, sheet_subtitle,
                         f"tbl_{idx}", job.table, job.table_style, job.hierarchy)
    return path


# --------------------------------------------------------------------------
# 3) Workbook -> email
# --------------------------------------------------------------------------
def build_message(cfg: Config, attachment: Path, subject: str, intro: str, sheet_lines: list[str],
                  recipients: list[str] | None = None) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = cfg.mail_from
    msg["To"] = ", ".join(recipients or cfg.mail_to)
    if cfg.mail_cc:
        msg["Cc"] = ", ".join(cfg.mail_cc)
    # Adjust the subject and body wording to your needs
    msg["Subject"] = subject
    lines = "\n".join(f"  - {name}" for name in sheet_lines)
    msg.set_content(
        "Kính gửi anh/chị,\n\n"
        f"{intro}\n"
        f"{lines}\n\n"
        "Đây là email tự động, nếu cần chỉnh sửa hoặc thêm người nhận, vui lòng liên hệ người phụ trách báo cáo.\n"
    )
    msg.add_attachment(
        attachment.read_bytes(),
        maintype="application",
        subtype=XLSX_SUBTYPE,
        filename=attachment.name,
    )
    return msg


def send_message(cfg: Config, msg: EmailMessage, recipients: list[str] | None = None) -> None:
    recipients = recipients or (cfg.mail_to + cfg.mail_cc)
    context = ssl.create_default_context()

    if cfg.smtp_port == 465:  # implicit TLS
        with smtplib.SMTP_SSL(cfg.smtp_host, cfg.smtp_port, context=context, timeout=60) as server:
            if cfg.smtp_user:
                server.login(cfg.smtp_user, cfg.smtp_password or "")
            server.send_message(msg, to_addrs=recipients)
        return

    with smtplib.SMTP(cfg.smtp_host, cfg.smtp_port, timeout=60) as server:
        server.ehlo()
        if cfg.smtp_starttls:
            if not server.has_extn("starttls"):
                raise ReportError(
                    "The mail server does not offer STARTTLS "
                    "(set SMTP_STARTTLS=false only for a trusted internal relay)"
                )
            server.starttls(context=context)
            server.ehlo()
        if cfg.smtp_user:
            server.login(cfg.smtp_user, cfg.smtp_password or "")
        server.send_message(msg, to_addrs=recipients)


def send_alert(cfg: Config, subject: str, body: str) -> None:
    """Best effort: tell the person in charge that the job failed."""
    if not cfg.alert_to:
        return
    try:
        msg = EmailMessage()
        msg["From"] = cfg.mail_from
        msg["To"] = ", ".join(cfg.alert_to)
        msg["Subject"] = subject
        msg.set_content(body)
        send_message(cfg, msg, recipients=cfg.alert_to)
        log.info("Alert email sent to %s", ", ".join(cfg.alert_to))
    except Exception:  # noqa: BLE001 - an alert must never crash the job
        log.exception("Could not send the alert email")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def _parse_date(value: str) -> dt.date:
    try:
        return dt.date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"Invalid date '{value}', use YYYY-MM-DD") from exc


def group_recipients(cfg: Config, jobs: list[Job]) -> list[str]:
    """Recipients of one email: the addresses set on its reports, otherwise MAIL_TO."""
    own: list[str] = []
    for job in jobs:
        for address in job.mail_to or []:
            if address not in own:
                own.append(address)
    return own or cfg.mail_to


def run_group(cfg: Config, jobs: list[Job], ctx: dict[str, Any], label: str | None, out_file: Path,
              subject: str, intro: str, dry_run: bool) -> int:
    """Run a set of reports into one workbook and email it. Returns 0 on success, 1 on failure."""
    try:
        results = fetch_results(cfg, jobs, ctx)
        build_workbook(results, out_file, label)
        size_mb = out_file.stat().st_size / (1024 * 1024)
        log.info("Excel file written: %s (%.2f MB)", out_file, size_mb)
        if size_mb > cfg.max_attach_mb:
            raise ReportError(f"File is {size_mb:.1f} MB, above the {cfg.max_attach_mb} MB limit for email")

        if dry_run:
            log.info("Dry run: email not sent")
            return 0

        recipients = group_recipients(cfg, jobs)
        sheet_lines = [f"{job.sheet} ({job.period_label})" if job.period_label else job.sheet for job, _, _ in results]
        send_message(cfg, build_message(cfg, out_file, subject, intro, sheet_lines, recipients),
                     recipients + cfg.mail_cc)
        log.info("Email sent to %s", ", ".join(recipients + cfg.mail_cc))
        return 0
    except Exception as exc:  # noqa: BLE001
        log.exception("Report job failed (%s)", subject)
        send_alert(
            cfg,
            f"[LỖI] {subject}",
            f"Job báo cáo thất bại.\n\n{subject}\nLỗi: {type(exc).__name__}: {exc}\n\n"
            f"Chi tiết:\n{traceback.format_exc()}",
        )
        return 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run stored procedures, export to Excel, email the result.")
    parser.add_argument("--from-date", type=_parse_date, help="start date, YYYY-MM-DD")
    parser.add_argument("--to-date", type=_parse_date, help="end date, YYYY-MM-DD")
    parser.add_argument("--dry-run", action="store_true", help="build the Excel file but do not send the email")
    parser.add_argument("--only", action="append", metavar="SHEET", help="run only this sheet (repeatable)")
    parser.add_argument("--reports-file", type=Path, help="path to the reports JSON (default: reports.json)")
    parser.add_argument("--due", action="store_true",
                        help="run only the reports whose own schedule is due today, each with its own period")
    parser.add_argument("--today", type=_parse_date, help="with --due: pretend today is this date (YYYY-MM-DD)")
    args = parser.parse_args(argv)

    setup_logging()

    if args.due and (args.from_date or args.to_date):
        log.error("--due chooses the period from each report's schedule: do not combine it with --from-date/--to-date")
        return 2
    if args.today and not args.due:
        log.error("--today only makes sense together with --due")
        return 2
    if bool(args.from_date) != bool(args.to_date):
        log.error("Give both --from-date and --to-date, or neither")
        return 2
    from_date, to_date = (args.from_date, args.to_date) if args.from_date else default_period()
    if from_date > to_date:
        log.error("--from-date is after --to-date")
        return 2

    today = args.today or dt.date.today()
    ctx = make_context(from_date, to_date)

    try:
        cfg = Config.from_env()
        jobs = load_jobs(args.reports_file or cfg.reports_file)
        if args.only:
            wanted = {name.strip().lower() for name in args.only}
            available = {job.sheet.lower() for job in jobs}
            unknown = sorted(wanted - available)
            if unknown:
                raise ReportError(f"--only: unknown sheet(s) {unknown}; available: {[j.sheet for j in jobs]}")
            jobs = [job for job in jobs if job.sheet.lower() in wanted]
        if args.due:
            jobs = [job for job in jobs if job.schedule and is_due(job.schedule, today)]
            for job in jobs:
                job.base_ctx = make_context(*rule_period(job_period_rule(job), today), today)
        for job in jobs:  # catch unknown parameter tokens before connecting to anything
            build_call(job, ctx)
    except ReportError as exc:
        log.error("Configuration error: %s", exc)
        return 2

    # ---- scheduled mode: one workbook and one email per group of reports that are due today ----
    if args.due:
        if not jobs:
            log.info("No report is due on %s", today)
            return 0
        groups: dict[str, list[Job]] = {}
        for job in jobs:
            groups.setdefault(job.group or job.sheet, []).append(job)
        log.info("Reports due on %s: %s", today, "; ".join(f"{name} ({len(items)} sheet(s))" for name, items in groups.items()))
        titles = {name: group_title(name, items, ctx, today) for name, items in groups.items()}
        counts: dict[str, int] = {}
        for title in titles.values():
            counts[title] = counts.get(title, 0) + 1
        status = 0
        for name, items in groups.items():
            title = titles[name]
            if counts[title] > 1:  # two groups with the same period would share a name: add the group name
                title = f"{title} - {name}"
            out_file = cfg.output_dir / f"{safe_filename(title)}.xlsx"
            intro = f"Gửi anh/chị {title[0].lower() + title[1:]} (file Excel đính kèm), gồm các sheet (kỳ báo cáo trong ngoặc):"
            status = max(status, run_group(cfg, items, ctx, None, out_file, title, intro, args.dry_run))
        return status

    # ---- manual mode: everything in one workbook for the period given on the command line ----
    label = f"{from_date:%d/%m/%Y} - {to_date:%d/%m/%Y}"
    out_file = cfg.output_dir / f"bao_cao_{from_date:%Y%m%d}_{to_date:%Y%m%d}.xlsx"
    log.info("Starting report for %s (%d sheet(s))", label, len(jobs))
    return run_group(
        cfg, jobs, ctx, label, out_file,
        subject=f"Báo cáo định kỳ: {label}",
        intro=f"Gửi anh/chị báo cáo định kỳ {label} (file Excel đính kèm), gồm các sheet:",
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    sys.exit(main())
