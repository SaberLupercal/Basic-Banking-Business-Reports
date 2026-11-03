# report_job - stored procedures -> one Excel workbook -> email

Runs several SQL Server stored procedures, writes each result to its own sheet of one
.xlsx file (plus an info sheet with period, time, parameters and row counts), and emails it.

## Setup
1. Install **ODBC Driver 18 for SQL Server** on the machine that runs the script.
2. `pip install -r requirements.txt`
3. Copy `.env.example` to `.env` and fill in the SQL Server, mail and recipient settings.
4. Copy `reports.example.json` to `reports.json`, then check every procedure name and
   parameter name against your database. One entry = one sheet.

## Run
    python report_job.py --dry-run                       # build the file only, send nothing
    python report_job.py                                 # previous full calendar month, then email
    python report_job.py --from-date 2026-07-01 --to-date 2026-12-31
    python report_job.py --only "Co cau tin dung"        # only one sheet (repeatable)

Parameter tokens in `reports.json`: `{from_date}`, `{to_date}`, `{year}` (year of to_date), `{today}`.

## Standard form of a report
Convention of this project: **every report procedure takes the same two parameters**, `@ngay_batdau` and `@ngay_ketthuc`
(start and end date, both inclusive), and returns one result set. A report in `reports.json` then needs no `params`:

    { "sheet": "Co cau tin dung", "proc": "dbo.Bao_cao_tindung", "schedule": { "frequency": "quarterly" } }

The program passes the period (`{from_date}`, `{to_date}`) to those two parameters. Anything else goes in `extra_params`
(for example `"extra_params": { "@TYLE_DTBB": 0.05 }`; better still, give that parameter a default value in the procedure).
If your parameters have other names, set them once at the top of `reports.json`:
`"standard_params": { "from": "@tu_ngay", "to": "@den_ngay" }`. A procedure that does not follow the standard can still
list its own `params` (see below).

## Per-report schedules (`--due`)
Give each report its own fixed schedule in `reports.json`, then create **one daily trigger** that runs
`python report_job.py --due`. Every day the program checks which reports are due and runs only those.

    { "sheet": "Co cau tin dung", "proc": "dbo.Bao_cao_tindung", "params": { "@nam": "{year}" },
      "schedule": { "frequency": "quarterly", "day": 1 }, "group": "Tin dung",
      "mail_to": ["quanly@congty.vn"] }

| frequency | extra keys | due on | covers (default `period_rule`) |
|---|---|---|---|
| `daily` | none | every day | `yesterday` |
| `weekly` | `weekday`: monday..sunday (default monday) | that weekday | `previous_week` (Monday to Sunday) |
| `monthly` | `day`: 1..31 or `"last"` (default 1) | that day of the month | `previous_month` |
| `quarterly` | `day` (default 1) | that day in January, April, July, October | `previous_quarter` |
| `yearly` | `month` (default 1), `day` (default 1) | that day of that month | `previous_year` |

- A `day` larger than the month (31 in February) means the last day of that month.
- Other period rules: `month_to_date`, `year_to_date`, and `year_over_year` (31 December of the year before last versus 31 December of last year: for reports that compare two points in time, such as `Baocao_duno`). Set `"period_rule"` on a report to override the default.
  An explicit `"period"` on a report still wins over everything.
- `"group"`: reports of the same group that are due on the same day share one workbook and one email.
  A report without a group gets its own email. Reports without `"schedule"` are never run by `--due`.
- `"mail_to"`: recipients of that report (default `MAIL_TO`). If a group mixes lists, all addresses receive the email.
- Names come from the **period that is reported**, not from the run date. The email subject and the file name are
  `Báo cáo tháng 01/2024`, `Báo cáo quý IV/2024` or `Báo cáo tổng kết năm 2024` (the file name uses `-` instead of `/`,
  which Windows does not allow: `Báo cáo tháng 01-2024.xlsx`). A group whose reports cover another period (for example
  year to date) is named `Báo cáo <group> (<run date>)`. If two groups would get the same name in one run, the group name
  is added (`Báo cáo tháng 01/2024 - <group>`) so that no file overwrites another.
- Choose the wording per group in the `"groups"` block at the top of `reports.json`; `{period}` becomes `tháng 01/2024`,
  `quý IV/2024` or `năm 2024`:

      "groups": {
        "Hằng năm":   { "name": "Báo cáo tổng kết {period}" },
        "Hằng quý":   { "name": "Báo cáo định kỳ {period}" },
        "Hằng tháng": { "name": "Báo cáo tiền gửi không kỳ hạn {period}" }
      }

  gives `Báo cáo tổng kết năm 2024`, `Báo cáo định kỳ quý IV/2024` and `Báo cáo tiền gửi không kỳ hạn tháng 01/2024`.
  The key must match the `"group"` of the reports (a typo is reported). Without it the default names above are used.
- The email body lists every sheet with its period. Running the same day again rewrites the same file.
- If one group fails, the others still run; `ALERT_TO` gets an email for the failed one and the exit code is 1.

Test without waiting for the real date (nothing is sent with `--dry-run`):

    python report_job.py --due --today 2024-01-01 --dry-run

Limits: if the computer is off on the day a report is due, that run is skipped (there is no catch-up),
and running `--due` twice on the same day sends the emails twice.

## Schedule used in this project (`reports.example.json`)
One daily trigger runs `python report_job.py --due`. Three groups give three Excel files, two sheets each:

| File (group) | Sheets (procedures) | Due | Covers | Procedure input |
|---|---|---|---|---|
| Hằng năm | `tiengui_tietkiem1`, `Bao_cao_tindung` | 1 January | last year | the year |
| Hằng quý | `bao_cao_tiengui_khachhang`, `Baocao_duno` | 1 Jan, 1 Apr, 1 Jul, 1 Oct | the quarter just ended | first and last day of the quarter |
| Hằng tháng | `tinhlai_tiengui_khongkyhan`, `tinh_phiduytri_taikhoan` | 1st of every month | the month just ended | first and last day of the month |

`LAISUAT_SPTK` (impact of the reserve requirement) is **not scheduled on purpose**: it is a what-if analysis driven by a rate
chosen by a person, so it is run by hand in SSMS. Try it without waiting: `python report_job.py --due --today 2024-01-01 --dry-run`.

## Reporting period: global, per report, or year only
The command line gives one global period (`--from-date` / `--to-date`). Each report decides how to use it in `params`:

| The procedure needs | Put in `params` |
|---|---|
| a year only | `"@nam": "{year}"` (year of the end date) |
| the whole period | `"@NGAYBD": "{from_date}", "@NGAYKT": "{to_date}"` |
| a baseline snapshot, e.g. last day of the previous year versus the end date | `"@ngay_batdau": "{prev_year_end}", "@ngay_ketthuc": "{to_date}"` |
| one fixed date | `"@ngay": "2023-06-30"` |

Tokens: `{from_date}`, `{to_date}`, `{year}`, `{year_start}`, `{year_end}`, `{prev_year_end}`,
`{month_start}`, `{month_end}`, `{today}`.

A report can also have **its own period**, independent of the command line. The tokens then refer to that period,
and its sheet shows that period in the subtitle:

    {
      "sheet": "Quan ly khach hang",
      "proc": "dbo.BAOCAO_KHACHHANG",
      "period": { "from": "2023-07-01", "to": "{year_end}" },
      "params": { "@NGAYBD": "{from_date}", "@NGAYKT": "{to_date}" }
    }

`from` and `to` are each a `YYYY-MM-DD` date or a token. Leave one out to keep the global value.
Everything is checked before the program connects to the database, so a typo stops it with a clear message.

## Excel layout
Each sheet gets a title, the reporting period, a header row, number formats, filter buttons and a frozen header.
Reports whose first column is "Tiêu chí cấp 1" are shown as a hierarchy (level-1 rows in bold on a light fill).
Optional keys in `reports.json`, per report:
- `"title"`: heading shown at the top of the sheet (default: the sheet name)
- `"columns"`: list of column names that replaces the names the procedure returns (same number, same order);
  useful when a procedure returns columns without a name
- `"table": false`: plain range instead of an Excel table
- `"hierarchy": true/false`: force or disable the hierarchical layout
- `"table_style"`: any Excel table style name, e.g. `TableStyleMedium2`

## Behaviour
- All-or-nothing: if any procedure fails or returns no rows, nothing is sent to the recipients.
  The person in `ALERT_TO` gets an email with the error instead.
- Logs are written to `logs/report_job.log`.
- Exit codes: 0 ok, 1 runtime failure, 2 configuration or argument error.

## Schedule
- Windows Task Scheduler: **one daily trigger**. Program = path to python.exe, Arguments = `report_job.py --due`,
  Start in = this folder. Choose "Run whether user is logged on or not".
- Linux cron, 07:00 every day:
  `0 7 * * * cd /path/to/report_job && python3 report_job.py --due`
- For a single fixed run instead, use `report_job.py --from-date ... --to-date ...` as before.

## Security
- Use a SQL login that can only execute these procedures.
- Never commit `.env` to git and never email it.
