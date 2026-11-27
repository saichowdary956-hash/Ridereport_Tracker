
import csv
import hashlib
import io
import json
import os
from datetime import datetime

import psycopg
from psycopg.rows import dict_row
from flask import Flask, jsonify, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

FIELDS = [
    ("", ""), ("Driver", "Driver"), ("Annotator", "Annotator"),
    ("Day", "Day"), ("Route", "Route"), ("State, City", "State, City"),
    ("Start Point", "Start Point"), ("End Point", "End Point"),
    ("Drive Cycle Start Time", "Drive Cycle Start Time"),
    ("Drive Cycle End Time", "Drive Cycle End Time"),
    ("Calendar Week", "Calendar Week"), ("Date", "Date"),
    ("Vehicle", "Vehicle"), ("Drive ID", "Drive ID"),
    ("RSU Start Date", "RSU Start Date"), ("RSU Storage %", "RSU Storage %"),
    ("RSU Number", "RSU Number"), ("Sunny", "Sunny"), ("Low Sun", "Low Sun"),
    ("Cloudy", "Cloudy"), ("Rain", "Rain"), ("Fog", "Fog"), ("Snow", "Snow"),
    ("City (intense traffic)", "City (intense traffic)"), ("Country", "Country"),
    ("Highway", "Highway"), ("Construction Site", "Construction Site"),
    ("Tunnel", "Tunnel"), ("Lighting - Day", "Day"), ("Dawn", "Dawn"),
    ("Dusk", "Dusk"), ("Lit Night", "Lit Night"), ("Dark Night", "Dark Night"),
    ("Flow", "Flow"), ("Jam", "Jam"), ("03_18 mph", "03_18 mph"),
    ("18_37 mph", "18_37 mph"), ("37_55 mph", "37_55 mph"),
    ("55_80 mph", "55_80 mph"), ("81-155mph", "81-155mph"),
    ("Drive cycle minutes", "Drive cycle minutes"), ("hh:mm", "hh:mm"),
    ("Z Frame Checker (Status)", "Z Frame Checker (Status)"),
    ("Sanity Check-Comments", "Sanity Check-Comments"), ("Comments", "Comments")
]

KEYS = [k for k, _ in FIELDS]
HEADERS = [h for _, h in FIELDS]

DURATION_FIELDS = {
    "Sunny", "Low Sun", "Cloudy", "Rain", "Fog", "Snow",
    "City (intense traffic)", "Country", "Highway",
    "Construction Site", "Tunnel", "Lighting - Day", "Dawn",
    "Dusk", "Lit Night", "Dark Night", "Flow", "Jam",
    "03_18 mph", "18_37 mph", "37_55 mph",
    "55_80 mph", "81-155mph"
}

CATEGORY_MAP = {
    "Weather": {
        "Sunny": "Sunny", "Low Sun": "Low Sun",
        "Cloudy": "Cloudy", "Rain": "Rain",
        "Fog": "Fog", "Snow": "Snow"
    },
    "Road Type": {
        "City": "City (intense traffic)",
        "Country": "Country", "Highway": "Highway",
        "Construction": "Construction Site",
        "Construction Site": "Construction Site",
        "Tunnel": "Tunnel"
    },
    "Lighting": {
        "Day": "Lighting - Day", "Dawn": "Dawn",
        "Dusk": "Dusk", "Lit Night": "Lit Night",
        "Dark Night": "Dark Night"
    },
    "Traffic": {"Flow": "Flow", "Jam": "Jam"},
    "Speed": {
        "3-18mph": "03_18 mph", "3-18 mph": "03_18 mph",
        "18-37mph": "18_37 mph", "19-37mph": "18_37 mph",
        "37-55mph": "37_55 mph", "38-55mph": "37_55 mph",
        "55-80mph": "55_80 mph", "56-80mph": "55_80 mph",
        "81-155mph": "81-155mph"
    }
}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024


def database_url():
    value = os.environ.get("DATABASE_URL", "").strip()
    if value.startswith("postgres://"):
        value = "postgresql://" + value[len("postgres://"):]
    if not value:
        raise RuntimeError("DATABASE_URL is not configured")
    return value


def db():
    return psycopg.connect(
        database_url(),
        row_factory=dict_row,
        connect_timeout=10
    )


def init_db():
    with db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS reports (
                digest TEXT PRIMARY KEY,
                drive_id TEXT,
                filename TEXT NOT NULL,
                row_data JSONB NOT NULL,
                imported_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS corrections (
                digest TEXT NOT NULL REFERENCES reports(digest)
                    ON DELETE CASCADE,
                field TEXT NOT NULL,
                value TEXT,
                updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                PRIMARY KEY(digest, field)
            )
        """)


def parse_duration(value):
    try:
        parts = [
            int(x) for x in str(value or "0:0:0").strip().split(":")
        ]
        if len(parts) == 2:
            h, m = parts
            return h * 3600 + m * 60
        if len(parts) == 3:
            h, m, s = parts
            return h * 3600 + m * 60 + s
    except (TypeError, ValueError):
        pass
    return 0


def fmt_duration(seconds):
    h, rem = divmod(int(seconds), 3600)
    m, s = divmod(rem, 60)
    return f"{h:02}:{m:02}:{s:02}"


def parse_csv(raw, filename):
    text = raw.decode("utf-8-sig", errors="replace")
    rows = list(csv.reader(io.StringIO(text)))

    metadata = {}
    for row in rows:
        if (
            len(row) >= 2
            and row[0].strip()
            and row[0].strip() != "DataLogger Report"
        ):
            metadata[row[0].strip()] = row[1].strip()

    category_values = {}

    for idx, row in enumerate(rows):
        if row and row[0].strip() == "Weather":
            headings = [x.strip() for x in row]

            for data_row in rows[idx + 1:]:
                if not any(x.strip() for x in data_row):
                    continue

                for col in range(0, len(headings) - 1, 2):
                    group = headings[col]
                    label = (
                        data_row[col].strip()
                        if col < len(data_row) else ""
                    )
                    value = (
                        data_row[col + 1].strip()
                        if col + 1 < len(data_row) else ""
                    )

                    if group in CATEGORY_MAP and label in CATEGORY_MAP[group]:
                        category_values[
                            CATEGORY_MAP[group][label]
                        ] = value or "00:00:00"
            break

    raw_date = metadata.get("Date", "")
    parsed_date = None

    for date_format in ("%m/%d/%y", "%m/%d/%Y"):
        try:
            parsed_date = datetime.strptime(raw_date, date_format)
            break
        except ValueError:
            continue

    overall = metadata.get("Overall Sessiontime", "")
    total_seconds = parse_duration(overall)

    record = {key: "" for key in KEYS if key}

    record.update({
        "Driver": metadata.get("Driver", ""),
        "Annotator": metadata.get("Annotator", ""),
        "Date": (
            parsed_date.strftime("%m/%d/%Y")
            if parsed_date else raw_date
        ),
        "Day": parsed_date.strftime("%A") if parsed_date else "",
        "Calendar Week": (
            f"CW{parsed_date.isocalendar().week:02}"
            if parsed_date else ""
        ),
        "Drive ID": metadata.get("Drive id", ""),
        "RSU Start Date": metadata.get("RSU Startdate", ""),
        "RSU Storage %": metadata.get("RSU Storage", "").replace("%", ""),
        "RSU Number": metadata.get("RSU No", ""),
        "Drive Cycle Start Time": metadata.get("Session Starttime", ""),
        "Drive Cycle End Time": metadata.get("Session Endtime", ""),
        "Drive cycle minutes": overall,
        "hh:mm": (
            f"{total_seconds // 3600:02}:"
            f"{(total_seconds % 3600) // 60:02}"
        ),
        "Comments": metadata.get("Comments", "")
    })

    record.update(category_values)

    digest = hashlib.sha256(raw).hexdigest()
    return digest, record


def get_records():
    init_db()

    with db() as conn:
        records = conn.execute("""
            SELECT digest, row_data
            FROM reports
            ORDER BY
                CASE
                    WHEN row_data->>'Date' ~ '^\\d{2}/\\d{2}/\\d{4}$'
                    THEN TO_DATE(row_data->>'Date', 'MM/DD/YYYY')
                    ELSE DATE '9999-12-31'
                END,
                row_data->>'Drive Cycle Start Time',
                row_data->>'Drive ID'
        """).fetchall()

        for item in records:
            fixes = conn.execute(
                "SELECT field, value FROM corrections WHERE digest=%s",
                (item["digest"],)
            ).fetchall()

            for fix in fixes:
                item["row_data"][fix["field"]] = fix["value"]

    return records


def make_workbook():
    records = get_records()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Daily Tracker"
    sheet.sheet_view.showGridLines = False
    sheet.append(HEADERS)
    sheet.row_dimensions[1].height = 42

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        col = cell.column

        if 18 <= col <= 23:
            color = "CFE2F3"
        elif 24 <= col <= 28:
            color = "FCE4D6"
        elif 29 <= col <= 33:
            color = "E2F0D9"
        elif 34 <= col <= 35:
            color = "D9EAD3"
        elif 36 <= col <= 41:
            color = "FCE4D6"
        else:
            color = "D9EAD3"

        cell.fill = PatternFill("solid", fgColor=color)

    for item in records:
        row = item["row_data"]
        sheet.append([
            row.get(key, "") if key else ""
            for key in KEYS
        ])

    total_row = len(records) + 2

    for col, key in enumerate(KEYS, 1):
        cell = sheet.cell(total_row, col)

        if key in DURATION_FIELDS or key == "Drive cycle minutes":
            cell.value = fmt_duration(sum(
                parse_duration(r["row_data"].get(key, ""))
                for r in records
            ))

        elif key == "hh:mm":
            seconds = sum(
                parse_duration(r["row_data"].get("Drive cycle minutes", ""))
                for r in records
            )
            cell.value = (
                f"{seconds // 3600:02}:"
                f"{(seconds % 3600) // 60:02}"
            )

        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="F4B183")

    widths = {
        1: 3, 2: 16, 3: 15, 6: 18,
        7: 28, 8: 28, 14: 18,
        43: 20, 44: 38, 45: 48
    }

    for col in range(1, len(HEADERS) + 1):
        sheet.column_dimensions[
            get_column_letter(col)
        ].width = widths.get(col, 12)

    sheet.freeze_panes = "B2"
    sheet.auto_filter.ref = (
        f"A1:{get_column_letter(len(HEADERS))}"
        f"{max(1, total_row - 1)}"
    )

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


@app.get("/health")
def health():
    return jsonify(ok=True, service="ride-report-api")


@app.get("/api/status")
def status():
    return jsonify(ok=True, records=len(get_records()))


@app.post("/api/import")
def import_csv():
    uploaded = request.files.get("file")
    payload = request.get_json(silent=True) or {}

    if uploaded:
        raw = uploaded.read()
        filename = uploaded.filename or "ride-report.csv"
    else:
        raw = (payload.get("csv") or "").encode("utf-8")
        filename = payload.get("filename") or "ride-report.csv"

    if not raw.strip():
        return jsonify(ok=False, error="CSV file is empty"), 400

    digest, record = parse_csv(raw, filename)
    init_db()

    with db() as conn:
        inserted = conn.execute("""
            INSERT INTO reports(digest, drive_id, filename, row_data)
            VALUES(%s, %s, %s, %s)
            ON CONFLICT(digest) DO NOTHING
            RETURNING digest
        """, (
            digest,
            record.get("Drive ID"),
            filename,
            json.dumps(record)
        )).fetchone()

        same_id_count = 0

        if record.get("Drive ID"):
            same_id_count = conn.execute(
                "SELECT COUNT(*) AS count FROM reports WHERE drive_id=%s",
                (record["Drive ID"],)
            ).fetchone()["count"]

    return jsonify(
        ok=True,
        created=bool(inserted),
        duplicate=not bool(inserted),
        drive_id=record.get("Drive ID"),
        date=record.get("Date"),
        drive_id_conflict=same_id_count > 1,
        message=(
            "Imported" if inserted
            else "Exact duplicate ignored"
        )
    )


@app.patch("/api/reports/<digest>")
def edit_report(digest):
    changes = request.get_json(silent=True) or {}
    allowed = set(KEYS) - {""}

    if not changes or any(field not in allowed for field in changes):
        return jsonify(ok=False, error="Invalid fields"), 400

    init_db()

    with db() as conn:
        exists = conn.execute(
            "SELECT 1 FROM reports WHERE digest=%s",
            (digest,)
        ).fetchone()

        if not exists:
            return jsonify(ok=False, error="Report not found"), 404

        for field, value in changes.items():
            conn.execute("""
                INSERT INTO corrections(digest, field, value)
                VALUES(%s, %s, %s)
                ON CONFLICT(digest, field)
                DO UPDATE SET
                    value=EXCLUDED.value,
                    updated_at=NOW()
            """, (digest, field, str(value)))

    return jsonify(ok=True, corrected=list(changes))


@app.get("/download")
def download():
    return send_file(
        make_workbook(),
        as_attachment=True,
        download_name="Daily_Ride_Report_Tracker.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


@app.after_request
def cors(response):
    allowed_origin = os.environ.get(
        "FRONTEND_ORIGIN", ""
    ).strip()

    origin = request.headers.get("Origin", "")

    if allowed_origin and origin == allowed_origin:
        response.headers["Access-Control-Allow-Origin"] = allowed_origin
        response.headers["Vary"] = "Origin"

    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, PATCH, OPTIONS"

    return response


@app.route("/api/<path:path>", methods=["OPTIONS"])
@app.route("/download", methods=["OPTIONS"])
def preflight(path=None):
    return ("", 204)


@app.errorhandler(413)
def too_large(_error):
    return jsonify(
        ok=False,
        error="CSV exceeds the 20 MB upload limit"
    ), 413


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", "10000"))
    )
