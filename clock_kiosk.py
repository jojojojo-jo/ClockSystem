import csv
import html
import json
import os
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, HTTPServer
from threading import Timer

LOG_FILE = "raw_logs.csv"
WORKERS_FILE = "workers.csv"
REPORT_PREFIX = "Warehouse_Report_"  # Generates Warehouse_Report_YYYY-MM-DD.xlsx
COOLDOWN_SECONDS = 20
ADMIN_PIN = "1234"
PORT = 8080


def init_files():
    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, "w", newline="") as f:
            csv.writer(f).writerow(["Timestamp", "Date", "Worker_ID", "Worker_Name", "Action", "Notes"])
    if not os.path.exists(WORKERS_FILE):
        with open(WORKERS_FILE, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Worker_ID", "Worker_Name"])
            w.writerow(["1001", "Example Worker"])


def load_workers():
    workers = {}
    with open(WORKERS_FILE, "r", newline="") as f:
        for row in csv.DictReader(f):
            workers[row["Worker_ID"].strip()] = row["Worker_Name"].strip()
    return workers


def parse_ts(ts_val):
    """Parses text timestamps (YYYY-MM-DD HH:MM, DD.MM.YY HH:MM, etc.) AND Excel numeric serial dates."""
    s = str(ts_val).strip()

    # 1. Handle Excel numeric serial date/time (e.g., "46350.52986111111")
    try:
        serial = float(s)
        # Excel's day 0 is 1899-12-30 (accounting for Excel's 1900 leap-year bug)
        dt = datetime(1899, 12, 30) + timedelta(days=serial)
        # Round to nearest minute to avoid floating-point 12:42:59.999 drift
        if dt.second >= 30:
            dt += timedelta(minutes=1)
        return dt.replace(second=0, microsecond=0)
    except ValueError:
        pass

    # 2. Handle standard & European/German text date-time formats
    formats = (
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%d.%m.%y %H:%M",
        "%d.%m.%y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%y %H:%M",
    )
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue

    raise ValueError(f"Invalid timestamp format: {s}")


def get_today_scans(worker_id, today_str):
    with open(LOG_FILE, "r", newline="") as f:
        return [r for r in csv.DictReader(f) if r["Worker_ID"] == worker_id and r["Date"] == today_str]


def recalculate_worker_day(rows, worker_id, target_date):
    rows.sort(key=lambda r: parse_ts(r["Timestamp"]))
    scan_index = 0
    for row in rows:
        if row["Worker_ID"] == worker_id and row["Date"] == target_date:
            expected = "IN" if scan_index % 2 == 0 else "OUT"
            if row["Action"] != expected:
                row["Action"] = expected
                if "Auto-Corrected" not in row["Notes"]:
                    row["Notes"] = f"{row['Notes']} (Auto-Corrected)".strip()
            scan_index += 1
    return rows


# ==============================================================================
# .XLSX WRITER & READER 
# ==============================================================================
def _col_letter(n):
    res = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res = chr(65 + rem) + res
    return res


def _col_index_from_ref(cell_ref):
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch.upper()) - 64)
    return idx - 1


def _build_worksheet_xml(headers, data_rows, status_col_idx=None):
    all_rows = [headers] + data_rows

    cols_xml = ["<cols>"]
    for c_idx in range(len(headers)):
        max_len = max((len(str(r[c_idx])) for r in all_rows if c_idx < len(r)), default=10)
        width = max(max_len + 5, 13)
        cols_xml.append(f'<col min="{c_idx + 1}" max="{c_idx + 1}" width="{width}" customWidth="1"/>')
    cols_xml.append("</cols>")

    rows_xml = ["<sheetData>"]
    for r_idx, row in enumerate(all_rows, start=1):
        rows_xml.append(f'<row r="{r_idx}">')
        for c_idx, val in enumerate(row, start=1):
            cell_ref = f"{_col_letter(c_idx)}{r_idx}"
            if r_idx == 1:
                style = ' s="1"'
            elif status_col_idx and c_idx == status_col_idx:
                style = ' s="2"' if str(val) == "Complete" else ' s="3"'
            else:
                style = ' s="0"'

            if isinstance(val, (int, float)) and r_idx > 1:
                rows_xml.append(f'<c r="{cell_ref}"{style}><v>{val}</v></c>')
            else:
                escaped = html.escape(str(val))
                rows_xml.append(f'<c r="{cell_ref}" t="inlineStr"{style}><is><t>{escaped}</t></is></c>')
        rows_xml.append("</row>")
    rows_xml.append("</sheetData>")

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + "".join(cols_xml)
        + "".join(rows_xml)
        + "</worksheet>"
    )


def write_native_xlsx(filepath, sheets):
    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    workbook_rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    ]
    workbook_sheets = []

    for idx, (sheet_name, _, _, _) in enumerate(sheets, start=1):
        content_types.append(
            f'<Override PartName="/xl/worksheets/sheet{idx}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        workbook_rels.append(
            f'<Relationship Id="rId{idx}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{idx}.xml"/>'
        )
        workbook_sheets.append(f'<sheet name="{html.escape(sheet_name)}" sheetId="{idx}" r:id="rId{idx}"/>')

    content_types.append("</Types>")
    workbook_rels.append("</Relationships>")

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>" + "".join(workbook_sheets) + "</sheets></workbook>"
    )

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF166534"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF92400E"/><name val="Calibri"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFCBD5E1"/></left>
      <right style="thin"><color rgb="FFCBD5E1"/></right>
      <top style="thin"><color rgb="FFCBD5E1"/></top>
      <bottom style="thin"><color rgb="FFCBD5E1"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
</styleSheet>"""

    with zipfile.ZipFile(filepath, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "".join(content_types))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", "".join(workbook_rels))
        zf.writestr("xl/styles.xml", styles_xml)
        for idx, (_, headers, data_rows, status_col) in enumerate(sheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", _build_worksheet_xml(headers, data_rows, status_col))


def export_single_day_xlsx(target_date, all_rows):
    """Creates Warehouse_Report_YYYY-MM-DD.xlsx with HH:MM timestamps and break columns."""
    day_rows = [r for r in all_rows if r["Date"] == target_date]
    if not day_rows:
        return None

    day_rows.sort(key=lambda r: parse_ts(r["Timestamp"]))
    grouped_by_worker = defaultdict(list)
    for r in day_rows:
        grouped_by_worker[r["Worker_ID"]].append(r)

    processed_records = []
    max_breaks = 1

    for worker_id, scans in sorted(grouped_by_worker.items()):
        worker_name = scans[-1]["Worker_Name"]
        work_sec = break_sec = 0.0
        in_time = last_out = None
        break_starts = []
        break_ends = []

        for idx, s in enumerate(scans):
            ts = parse_ts(s["Timestamp"])
            time_hhmm = ts.strftime("%H:%M")

            if s["Action"] == "IN":
                if last_out is not None:
                    break_sec += (ts - last_out).total_seconds()
                    break_ends.append(time_hhmm)
                in_time = ts
            elif s["Action"] == "OUT" and in_time is not None:
                work_sec += (ts - in_time).total_seconds()
                last_out = ts
                in_time = None
                if idx < len(scans) - 1:
                    break_starts.append(time_hhmm)

        max_breaks = max(max_breaks, len(break_starts), len(break_ends))
        processed_records.append({
            "date_str": target_date,
            "worker_id": worker_id,
            "worker_name": worker_name,
            "scans": scans,
            "break_starts": break_starts,
            "break_ends": break_ends,
            "work_sec": work_sec,
            "break_sec": break_sec,
        })

    break_headers = []
    for b in range(1, max_breaks + 1):
        break_headers.extend([f"Break_{b}_Start", f"Break_{b}_End"])

    sum_headers = (
        ["Date", "Worker_ID", "Worker_Name", "First_IN"]
        + break_headers
        + ["Last_OUT", "Total_Scans", "Break_Minutes", "Work_Hours", "Status"]
    )

    sum_rows = []
    for rec in processed_records:
        scans = rec["scans"]
        is_complete = len(scans) % 2 == 0
        hrs, mins = int(rec["work_sec"] // 3600), int((rec["work_sec"] % 3600) // 60)

        first_in_str = parse_ts(scans[0]["Timestamp"]).strftime("%H:%M")
        last_out_str = parse_ts(scans[-1]["Timestamp"]).strftime("%H:%M") if is_complete else "STILL CLOCKED IN"
        break_min_val = int(round(rec["break_sec"] / 60.0))
        work_hhmm = f"{hrs:02d}:{mins:02d}"
        status_text = "Complete" if is_complete else "WARNING: Missing OUT Punch"

        break_cells = []
        for b in range(max_breaks):
            b_start = rec["break_starts"][b] if b < len(rec["break_starts"]) else "-"
            b_end = rec["break_ends"][b] if b < len(rec["break_ends"]) else "-"
            break_cells.extend([b_start, b_end])

        sum_rows.append(
            [rec["date_str"], rec["worker_id"], rec["worker_name"], first_in_str]
            + break_cells
            + [last_out_str, len(scans), break_min_val, work_hhmm, status_text]
        )

    log_headers = ["Timestamp", "Date", "Worker_ID", "Worker_Name", "Action", "Notes"]
    log_rows = [
        [
            parse_ts(r["Timestamp"]).strftime("%Y-%m-%d %H:%M"),
            r["Date"],
            r["Worker_ID"],
            r["Worker_Name"],
            r["Action"],
            r["Notes"],
        ]
        for r in day_rows
    ]

    worker_headers = ["Worker_ID", "Worker_Name"]
    worker_rows = [[wid, wname] for wid, wname in sorted(load_workers().items())]

    filepath = f"{REPORT_PREFIX}{target_date}.xlsx"
    sheets = [
        ("Daily_Summary", sum_headers, sum_rows, len(sum_headers)),
        ("Raw_Logs", log_headers, log_rows, None),
        ("Workers", worker_headers, worker_rows, None),
    ]
    write_native_xlsx(filepath, sheets)
    return filepath


def generate_excel_reports(target_date=None):
    with open(LOG_FILE, "r", newline="") as f:
        rows = list(csv.DictReader(f))

    if not rows:
        return False, "No scan logs found to export."

    if target_date:
        created = export_single_day_xlsx(target_date, rows)
        if not created:
            return False, f"No logs found for {target_date}."
        return True, f"Exported {created}."

    unique_dates = sorted({r["Date"] for r in rows if r["Date"]})
    created_files = []
    for dt_str in unique_dates:
        fp = export_single_day_xlsx(dt_str, rows)
        if fp:
            created_files.append(fp)

    return True, f"Generated {len(created_files)} daily Excel file(s): {', '.join(created_files)}"


def sync_from_native_xlsx(target_date):
    """Reads edited rows from Warehouse_Report_YYYY-MM-DD.xlsx (accepts HH:MM or HH:MM:SS), updates raw_logs.csv, and refreshes the daily Excel file."""
    filepath = f"{REPORT_PREFIX}{target_date}.xlsx"
    if not os.path.exists(filepath):
        return False, f"{filepath} not found. Export the Excel file first."

    try:
        ns_main = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        with zipfile.ZipFile(filepath, "r") as zf:
            shared_strings = []
            if "xl/sharedStrings.xml" in zf.namelist():
                ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
                for si in ss_root.findall("m:si", ns_main):
                    text_parts = [t.text or "" for t in si.findall(".//m:t", ns_main)]
                    shared_strings.append("".join(text_parts))

            wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
            rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root}

            target_sheet_path = None
            for sheet in wb_root.findall(".//m:sheet", ns_main):
                if sheet.attrib.get("name") == "Raw_Logs":
                    rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    target = rel_map.get(rid, "")
                    target_sheet_path = target if target.startswith("xl/") else f"xl/{target}"
                    break

            if not target_sheet_path or target_sheet_path not in zf.namelist():
                return False, "Could not find the 'Raw_Logs' tab inside the Excel workbook."

            ws_root = ET.fromstring(zf.read(target_sheet_path))
            parsed_rows = []
            for row_el in ws_root.findall(".//m:sheetData/m:row", ns_main):
                row_vals = {}
                for c_el in row_el.findall("m:c", ns_main):
                    c_idx = _col_index_from_ref(c_el.attrib.get("r", "A1"))
                    c_type = c_el.attrib.get("t", "")
                    val = ""
                    if c_type == "s":
                        v_el = c_el.find("m:v", ns_main)
                        if v_el is not None and v_el.text is not None:
                            val = shared_strings[int(v_el.text)]
                    elif c_type == "inlineStr":
                        t_el = c_el.find(".//m:t", ns_main)
                        val = t_el.text if (t_el is not None and t_el.text) else ""
                    else:
                        v_el = c_el.find("m:v", ns_main)
                        val = v_el.text if (v_el is not None and v_el.text) else ""
                    row_vals[c_idx] = val.strip()

                if row_vals:
                    max_c = max(row_vals.keys())
                    parsed_rows.append([row_vals.get(i, "") for i in range(max_c + 1)])

        if len(parsed_rows) < 2:
            return False, "No data rows found in the Raw_Logs tab."

        headers = parsed_rows[0]
        workers_map = load_workers()
        updated_day_rows = []
        affected_workers = set()

        for r in parsed_rows[1:]:
            row_dict = {headers[i]: (r[i] if i < len(r) else "") for i in range(len(headers))}
            ts = row_dict.get("Timestamp", "").strip()
            wid = row_dict.get("Worker_ID", "").strip()
            if not ts or not wid:
                continue

            try:
                dt_obj = parse_ts(ts)
            except ValueError:
                return False, f"Invalid Timestamp '{ts}' in Excel. Use YYYY-MM-DD HH:MM format."

            clean_ts = dt_obj.strftime("%Y-%m-%d %H:%M")
            date_str = dt_obj.strftime("%Y-%m-%d")
            wname = row_dict.get("Worker_Name", "").strip() or workers_map.get(wid, f"ID {wid}")
            notes = row_dict.get("Notes", "").strip() or "Excel Manual Edit"

            updated_day_rows.append({
                "Timestamp": clean_ts,
                "Date": date_str,
                "Worker_ID": wid,
                "Worker_Name": wname,
                "Action": row_dict.get("Action", "").strip().upper(),
                "Notes": notes,
            })
            affected_workers.add((wid, date_str))

        with open(LOG_FILE, "r", newline="") as f:
            existing_other_days = [r for r in csv.DictReader(f) if r["Date"] != target_date]

        combined_rows = existing_other_days + updated_day_rows
        for wid, dt_str in affected_workers:
            combined_rows = recalculate_worker_day(combined_rows, wid, dt_str)

        fieldnames = ["Timestamp", "Date", "Worker_ID", "Worker_Name", "Action", "Notes"]
        with open(LOG_FILE, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(combined_rows)

        export_single_day_xlsx(target_date, combined_rows)
        return True, f"Synced {len(updated_day_rows)} rows from {filepath} and refreshed Daily_Summary!"

    except PermissionError:
        return False, f"Close {filepath} in Excel first, then click Sync again."
    except Exception as e:
        return False, f"Sync error: {e}"


HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Warehouse Time Clock</title>
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, sans-serif; background: #1e293b; color: #f8fafc;
         display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; margin: 0; }
  h1 { font-size: 2.4rem; margin-bottom: 0.3rem; }
  #clock { font-size: 1.5rem; color: #94a3b8; margin-bottom: 2rem; }
  input { font-size: 2.5rem; text-align: center; width: 260px; padding: 12px; border-radius: 10px;
          border: 3px solid #38bdf8; outline: none; font-weight: bold; }
  #status { margin-top: 2rem; font-size: 1.6rem; font-weight: bold; min-height: 3.5rem; text-align: center; max-width: 700px; }
  .controls { position: fixed; bottom: 24px; display: flex; gap: 12px; flex-wrap: wrap; justify-content: center; }
  button { padding: 10px 16px; font-size: 0.92rem; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; }
  .btn-fix { background: #334155; color: #e2e8f0; }
  .btn-export { background: #0f766e; color: #fff; }
  .btn-export-all { background: #115e59; color: #cbd5e1; }
  .btn-sync { background: #1d4ed8; color: #fff; }
</style>
</head>
<body>
  <h1>WAREHOUSE CLOCK IN / OUT</h1>
  <div id="clock"></div>
  <input id="scanInput" type="text" placeholder="Scan ID" autofocus autocomplete="off" />
  <div id="status" style="color:#38bdf8;">Ready for scan...</div>

  <div class="controls">
    <button class="btn-fix" onclick="supervisorFix()">Supervisor: Fix Missed Punch</button>
    <button class="btn-export" onclick="exportTodayExcel()">Export Today's Excel (.xlsx)</button>
    <button class="btn-export-all" onclick="exportAllExcel()">Export All Days (.xlsx)</button>
    <button class="btn-sync" onclick="syncExcel()">Sync Edits from Daily Excel</button>
  </div>

<script>
  const input = document.getElementById('scanInput');
  const statusEl = document.getElementById('status');

  document.addEventListener('click', () => input.focus());
  setInterval(() => {
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    document.getElementById('clock').textContent = now.toLocaleDateString('en-CA') + '  |  ' + hh + ':' + mm;
  }, 500);

  input.addEventListener('keydown', async (e) => {
    if (e.key === 'Enter') {
      const code = input.value.trim();
      input.value = '';
      if (!code) return;
      const res = await fetch('/scan', { method: 'POST', body: JSON.stringify({ worker_id: code }) });
      const data = await res.json();
      statusEl.style.color = data.color;
      statusEl.textContent = data.message;
    }
  });

  async function supervisorFix() {
    const pin = prompt("Enter Supervisor PIN:");
    if (!pin) return;
    const worker_id = prompt("Enter Worker 4-Digit ID:");
    if (!worker_id) return;
    const target_date = prompt("Enter Date (YYYY-MM-DD):", new Date().toLocaleDateString('en-CA'));
    if (!target_date) return;
    const missed_time = prompt("Enter Missed Time (HH:MM, 24h format), e.g., 12:00:");
    if (!missed_time) return;

    const res = await fetch('/fix', {
      method: 'POST',
      body: JSON.stringify({ pin, worker_id, target_date, missed_time })
    });
    const data = await res.json();
    alert(data.message);
    input.focus();
  }

  async function exportTodayExcel() {
    const today = new Date().toLocaleDateString('en-CA');
    const res = await fetch('/export', { method: 'POST', body: JSON.stringify({ target_date: today }) });
    const data = await res.json();
    alert(data.message);
    input.focus();
  }

  async function exportAllExcel() {
    const res = await fetch('/export', { method: 'POST', body: JSON.stringify({}) });
    const data = await res.json();
    alert(data.message);
    input.focus();
  }

  async function syncExcel() {
    const pin = prompt("Enter Supervisor PIN to sync manual Excel edits:");
    if (!pin) return;
    const target_date = prompt("Enter the Date of the Excel file to sync (YYYY-MM-DD):", new Date().toLocaleDateString('en-CA'));
    if (!target_date) return;

    const res = await fetch('/sync', {
      method: 'POST',
      body: JSON.stringify({ pin, target_date })
    });
    const data = await res.json();
    alert(data.message);
    input.focus();
  }
</script>
</body>
</html>
"""


class KioskHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def _send_json(self, payload):
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(payload).encode())

    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        self.wfile.write(HTML_PAGE.encode())

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length).decode()) if length else {}
        workers = load_workers()

        if self.path == "/scan":
            worker_id = body.get("worker_id", "").strip()
            now = datetime.now()
            today_str = now.strftime("%Y-%m-%d")
            ts_str = now.strftime("%Y-%m-%d %H:%M")
            worker_name = workers.get(worker_id, f"ID {worker_id}")
            today_scans = get_today_scans(worker_id, today_str)

            if today_scans:
                last_ts = parse_ts(today_scans[-1]["Timestamp"])
                elapsed = (now - last_ts).total_seconds()
                if elapsed < COOLDOWN_SECONDS:
                    wait_left = max(1, int(COOLDOWN_SECONDS - elapsed))
                    self._send_json({
                        "color": "#f59e0b",
                        "message": f"DUPLICATE SCAN IGNORED ({worker_name}) — Wait {wait_left}s"
                    })
                    return

            action = "IN" if len(today_scans) % 2 == 0 else "OUT"
            with open(LOG_FILE, "a", newline="") as f:
                csv.writer(f).writerow([ts_str, today_str, worker_id, worker_name, action, "Auto Scan"])

            color = "#22c55e" if action == "IN" else "#ef4444"
            self._send_json({
                "color": color,
                "message": f"{worker_name} — Clocked {action} at {now.strftime('%H:%M')}"
            })

        elif self.path == "/fix":
            if body.get("pin") != ADMIN_PIN:
                self._send_json({"message": "Access Denied: Incorrect PIN."})
                return

            worker_id = body.get("worker_id", "").strip()
            target_date = body.get("target_date", "").strip()
            missed_time = body.get("missed_time", "").strip()
            full_ts = f"{target_date} {missed_time}"

            try:
                dt_obj = parse_ts(full_ts)
                clean_ts = dt_obj.strftime("%Y-%m-%d %H:%M")
            except ValueError:
                self._send_json({"message": "Invalid format. Use YYYY-MM-DD and HH:MM (e.g., 12:00)."})
                return

            worker_name = workers.get(worker_id, f"ID {worker_id}")
            with open(LOG_FILE, "r", newline="") as f:
                rows = list(csv.DictReader(f))

            rows.append({
                "Timestamp": clean_ts, "Date": target_date, "Worker_ID": worker_id,
                "Worker_Name": worker_name, "Action": "", "Notes": "Supervisor Manual Fix"
            })
            rows = recalculate_worker_day(rows, worker_id, target_date)

            with open(LOG_FILE, "w", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=["Timestamp", "Date", "Worker_ID", "Worker_Name", "Action", "Notes"])
                writer.writeheader()
                writer.writerows(rows)

            assigned = next((r["Action"] for r in rows if r["Worker_ID"] == worker_id and r["Timestamp"] == clean_ts), "UPDATED")
            self._send_json({"message": f"Inserted {assigned} at {clean_ts} for {worker_name} and recalculated day."})

        elif self.path == "/export":
            target_date = body.get("target_date", "").strip() or None
            _, msg = generate_excel_reports(target_date)
            self._send_json({"message": msg})

        elif self.path == "/sync":
            if body.get("pin") != ADMIN_PIN:
                self._send_json({"message": "Access Denied: Incorrect PIN."})
                return
            target_date = body.get("target_date", "").strip()
            _, msg = sync_from_native_xlsx(target_date)
            self._send_json({"message": msg})


class ReusableHTTPServer(HTTPServer):
    allow_reuse_address = True


if __name__ == "__main__":
    init_files()
    server = ReusableHTTPServer(("127.0.0.1", PORT), KioskHandler)
    print(f"Kiosk running at http://127.0.0.1:{PORT}")
    Timer(0.8, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}")).start()
    server.serve_forever()