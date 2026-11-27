import csv, io, os, hashlib, json
from datetime import datetime
from flask import Flask, request, jsonify, send_file, Response
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from psycopg.rows import dict_row
import psycopg

HEADERS = ["", "Driver", "Annotator", "Day", "Route", "State, City", "Start Point", "End Point", "Drive Cycle Start Time", "Drive Cycle End Time", "Calendar Week", "Date", "Vehicle", "Drive ID", "RSU Start Date", "RSU Storage %", "RSU Number", "Sunny", "Low Sun", "Cloudy", "Rain", "Fog", "Snow", "City (intense traffic)", "Country", "Highway", "Construction Site", "Tunnel", "Day", "Dawn", "Dusk", "Lit Night", "Dark Night", "Flow", "Jam", "03_18 mph", "18_37 mph", "37_55 mph", "55_80 mph", "81-155mph", "Drive cycle minutes", "hh:mm", "Z Frame Checker (Status)", "Sanity Check-Comments", "Comments"]
CATEGORY_MAP = {"Weather":{"Sunny":"Sunny","Low Sun":"Low Sun","Cloudy":"Cloudy","Rain":"Rain","Fog":"Fog","Snow":"Snow"},"Road Type":{"City":"City (intense traffic)","Country":"Country","Highway":"Highway","Construction":"Construction Site","Construction Site":"Construction Site","Tunnel":"Tunnel"},"Lighting":{"Day":"Day","Dawn":"Dawn","Dusk":"Dusk","Lit Night":"Lit Night","Dark Night":"Dark Night"},"Traffic":{"Flow":"Flow","Jam":"Jam"},"Speed":{"3-18mph":"03_18 mph","3-18 mph":"03_18 mph","19-37mph":"18_37 mph","18-37mph":"18_37 mph","37-55mph":"37_55 mph","38-55mph":"37_55 mph","56-80mph":"55_80 mph","55-80mph":"55_80 mph","81-155mph":"81-155mph"}}
DURATION_FIELDS = {v for g in CATEGORY_MAP.values() for v in g.values()}
DATABASE_URL = os.environ.get("DATABASE_URL", "")


def db():
    if not DATABASE_URL: raise RuntimeError("DATABASE_URL is not configured")
    url = DATABASE_URL.replace("postgres://", "postgresql://", 1)
    return psycopg.connect(url, row_factory=dict_row, sslmode="require")


def init_db():
    with db() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS reports (digest TEXT PRIMARY KEY, drive_id TEXT, filename TEXT NOT NULL, row_data JSONB NOT NULL, imported_at TIMESTAMPTZ NOT NULL DEFAULT NOW())""")
        c.execute("""CREATE TABLE IF NOT EXISTS corrections (digest TEXT NOT NULL REFERENCES reports(digest) ON DELETE CASCADE, field TEXT NOT NULL, value TEXT, updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), PRIMARY KEY(digest,field))""")


def parse_duration(s):
    try:
        h,m,sec = [int(x) for x in str(s or "0:0:0").split(":")]
        return h*3600+m*60+sec
    except Exception: return 0


def fmt_duration(n):
    h,r=divmod(int(n),3600); m,s=divmod(r,60)
    return f"{h:02}:{m:02}:{s:02}"


def parse_csv(raw, filename):
    rows=list(csv.reader(io.StringIO(raw.decode("utf-8-sig",errors="replace"))))
    meta={}
    for r in rows:
        if len(r)>=2 and r[0].strip() and r[0].strip()!="DataLogger Report": meta[r[0].strip()]=r[1].strip()
    cats={}
    for i,r in enumerate(rows):
        if r and r[0].strip()=="Weather":
            hdr=[x.strip() for x in r]
            for d in rows[i+1:]:
                if not any(x.strip() for x in d): continue
                for j in range(0,len(hdr)-1,2):
                    cat=hdr[j]; label=d[j].strip() if j<len(d) else ""; val=d[j+1].strip() if j+1<len(d) else ""
                    if cat in CATEGORY_MAP and label in CATEGORY_MAP[cat]: cats[CATEGORY_MAP[cat][label]]=val or "00:00:00"
            break
    rawdate=meta.get("Date",""); dt=None
    for f in ("%m/%d/%y","%m/%d/%Y"):
        try: dt=datetime.strptime(rawdate,f); break
        except ValueError: pass
    overall=meta.get("Overall Sessiontime",""); secs=parse_duration(overall)
    row={h:"" for h in HEADERS if h}
    row.update({"Driver":meta.get("Driver",""),"Annotator":meta.get("Annotator",""),"Date":dt.strftime("%m/%d/%Y") if dt else rawdate,"Day":dt.strftime("%A") if dt else "","Calendar Week":f"CW{dt.isocalendar().week:02}" if dt else "","Drive ID":meta.get("Drive id",""),"RSU Start Date":meta.get("RSU Startdate",""),"RSU Storage %":meta.get("RSU Storage","").replace("%",""),"RSU Number":meta.get("RSU No",""),"Drive Cycle Start Time":meta.get("Session Starttime",""),"Drive Cycle End Time":meta.get("Session Endtime",""),"Drive cycle minutes":overall,"hh:mm":f"{secs//3600:02}:{(secs%3600)//60:02}","Comments":meta.get("Comments","")})
    row.update(cats)
    return hashlib.sha256(raw).hexdigest(),row


def all_rows():
    init_db()
    with db() as c:
        records=c.execute("SELECT digest,row_data FROM reports ORDER BY (row_data->>'Date'), (row_data->>'Drive Cycle Start Time'), (row_data->>'Drive ID')").fetchall()
        for rec in records:
            fixes=c.execute("SELECT field,value FROM corrections WHERE digest=%s",(rec['digest'],)).fetchall()
            for fix in fixes: rec['row_data'][fix['field']]=fix['value']
    return [(r['digest'],r['row_data']) for r in records]


def build_xlsx():
    records=all_rows(); wb=Workbook(); ws=wb.active; ws.title="Daily Tracker"; ws.sheet_view.showGridLines=False
    ws.append(HEADERS); ws.row_dimensions[1].height=42
    for cell in ws[1]:
        cell.font=Font(bold=True); cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        col=cell.column; color="CFE2F3" if 18<=col<=23 else "FCE4D6" if 24<=col<=28 else "E2F0D9" if 29<=col<=33 else "D9EAD3" if 34<=col<=35 else "FCE4D6" if 36<=col<=40 else "D9EAD3"
        cell.fill=PatternFill("solid",fgColor=color)
    for _,row in records: ws.append([row.get(h,"") if h else "" for h in HEADERS])
    total=len(records)+2
    for col,h in enumerate(HEADERS,1):
        cell=ws.cell(total,col)
        if h in DURATION_FIELDS or h=="Drive cycle minutes": cell.value=fmt_duration(sum(parse_duration(r.get(h,"")) for _,r in records))
        elif h=="hh:mm":
            n=sum(parse_duration(r.get("Drive cycle minutes","")) for _,r in records); cell.value=f"{n//3600:02}:{(n%3600)//60:02}"
        cell.font=Font(bold=True); cell.fill=PatternFill("solid",fgColor="F4B183")
    widths={1:3,2:16,3:15,6:18,7:28,8:28,14:18,43:20,44:38,45:48}
    for col in range(1,46): ws.column_dimensions[get_column_letter(col)].width=widths.get(col,12)
    ws.freeze_panes="B2"; ws.auto_filter.ref=f"A1:AS{max(1,total-1)}"
    out=io.BytesIO(); wb.save(out); out.seek(0); return out

app=Flask(__name__)
@app.get("/health")
def health(): return jsonify(ok=True,service="ride-report-api")
@app.get("/api/status")
def status(): return jsonify(ok=True,records=len(all_rows()))
@app.post("/api/import")
def import_csv():
    token=os.environ.get("IMPORT_TOKEN","")
    if not token or request.headers.get("X-Import-Token")!=token: return jsonify(ok=False,error="Unauthorized"),401
    f=request.files.get("file"); payload=request.get_json(silent=True) or {}
    if f: raw=f.read(); filename=f.filename or "ride-report.csv"
    else: raw=(payload.get("csv") or "").encode(); filename=payload.get("filename") or "ride-report.csv"
    if not raw.strip(): return jsonify(ok=False,error="CSV file is empty"),400
    digest,row=parse_csv(raw,filename)
    init_db()
    with db() as c:
        cur=c.execute("INSERT INTO reports(digest,drive_id,filename,row_data) VALUES(%s,%s,%s,%s) ON CONFLICT(digest) DO NOTHING RETURNING digest",(digest,row.get("Drive ID"),filename,json.dumps(row)))
        created=cur.fetchone() is not None
        same_id=c.execute("SELECT COUNT(*) AS n FROM reports WHERE drive_id=%s",(row.get("Drive ID"),)).fetchone()['n'] if row.get("Drive ID") else 0
    return jsonify(ok=True,created=created,duplicate=not created,drive_id=row.get("Drive ID"),date=row.get("Date"),drive_id_conflict=same_id>1,message="Imported" if created else "Exact duplicate ignored")

@app.patch("/api/reports/<digest>")
def edit_report(digest):
    token=os.environ.get("IMPORT_TOKEN","")
    if not token or request.headers.get("X-Import-Token")!=token: return jsonify(ok=False,error="Unauthorized"),401
    changes=request.get_json(silent=True) or {}; allowed=set(HEADERS)-{""}
    if not changes or any(k not in allowed for k in changes): return jsonify(ok=False,error="Invalid fields"),400
    init_db()
    with db() as c:
        if not c.execute("SELECT 1 FROM reports WHERE digest=%s",(digest,)).fetchone(): return jsonify(ok=False,error="Report not found"),404
        for k,v in changes.items(): c.execute("INSERT INTO corrections(digest,field,value) VALUES(%s,%s,%s) ON CONFLICT(digest,field) DO UPDATE SET value=EXCLUDED.value,updated_at=NOW()",(digest,k,str(v)))
    return jsonify(ok=True,corrected=list(changes))

@app.get("/download")
def download(): return send_file(build_xlsx(),as_attachment=True,download_name="Daily_Ride_Report_Tracker.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.after_request
def cors(resp):
    origin=os.environ.get("FRONTEND_ORIGIN","")
    resp.headers["Access-Control-Allow-Origin"]=origin or "https://*.vercel.app"
    resp.headers["Access-Control-Allow-Headers"]="Content-Type,X-Import-Token"
    resp.headers["Access-Control-Allow-Methods"]="GET,POST,PATCH,OPTIONS"
    return resp

if __name__=="__main__":
    app.run(host="0.0.0.0",port=int(os.environ.get("PORT","10000")))
