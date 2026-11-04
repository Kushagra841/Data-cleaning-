import re
import os
import time
import secrets
import threading
import requests
from openpyxl import Workbook
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from pathlib import Path
from contextlib import asynccontextmanager
from fastapi import Depends, FastAPI, Header, HTTPException
from fastapi.responses import PlainTextResponse
from apscheduler.schedulers.background import BackgroundScheduler

# ---------- settings ----------
TALLY_URL = "http://localhost:9000"              # this API runs ON the Tally PC
COMPANY = "Shantam International UP (2026-27)"   # exactly as shown in Tally
FY_START = "20260401"                            # start of the financial year
OUTPUT_DIR = Path(r"C:\TallyData")               # a folder, not a file
RUN_AT_HOUR, RUN_AT_MINUTE = 21, 0               # daily auto-fetch at 9:00 PM
GST_MULTIPLIER = 1.18                            # assumption: check against real tax amounts
CHUNK_DAYS = 7                                   # max days per request to Tally
PAUSE_SECONDS = 3                                # wait between chunks
TIMEOUT_SECONDS = 180                            # max wait for one Tally reply
TRIES = 3                                        # attempts per request
RETRY_WAIT = 30                                  # seconds to wait before a retry

# The API listens on this PC's LAN address only (not on every network card).
# Change it if this PC's IP changes. Use "127.0.0.1" for local-only testing.
LISTEN_HOST = os.environ.get("TALLY_LISTEN_HOST", "192.168.1.27")
LISTEN_PORT = 8000
# Secret key callers must send in the X-API-Key header.
# Set it before starting:  setx TALLY_API_KEY "your-long-random-key"
API_KEY = os.environ.get("TALLY_API_KEY", "")

# Match every voucher type under Sales / Credit Note (same as the registers)
SALES_FILTER = "$$IsSales:$VoucherTypeName"
RETURNS_FILTER = "$$IsCreditNote:$VoucherTypeName"

# Stock is OFF until the sales and returns data are confirmed correct.
# Stock requests are heavy and crashed Tally before.
ENABLE_STOCK = False
# ------------------------------

tally_lock = threading.Lock()   # only one request to Tally at a time


def num(text):
    m = re.search(r"-?[\d,]*\.?\d+", text or "")
    return float(m.group().replace(",", "")) if m else 0.0


def clean_date(value, name):
    """Accept only a real YYYYMMDD date. Also blocks XML tricks and odd file names."""
    try:
        if len(value) != 8 or not value.isdigit():
            raise ValueError
        datetime.strptime(value, "%Y%m%d")
    except ValueError:
        raise HTTPException(400, f"{name} must be a real date in YYYYMMDD format, e.g. 20260401.")
    return value


def date_range(from_date, to_date):
    """Fill in today for missing dates, validate both, and check the order."""
    today = date.today().strftime("%Y%m%d")
    f = clean_date(from_date or today, "from_date")
    t = clean_date(to_date or today, "to_date")
    if t < f:
        raise HTTPException(400, "to_date is before from_date.")
    return f, t


def require_key(x_api_key: str = Header(default="")):
    if not API_KEY:
        raise HTTPException(503, "Server has no TALLY_API_KEY set.")
    if not secrets.compare_digest(x_api_key.encode(), API_KEY.encode()):
        raise HTTPException(401, "Missing or wrong X-API-Key header.")


def post_to_tally(xml: str) -> str:
    """Send one request to Tally and return the raw reply text."""
    # The lock is held during retries too, so nothing else reaches a struggling Tally.
    with tally_lock:
        for attempt in range(1, TRIES + 1):
            try:
                r = requests.post(TALLY_URL, data=xml.encode("utf-8"), timeout=TIMEOUT_SECONDS)
                return r.content.decode("utf-8", errors="ignore")
            except requests.RequestException:
                if attempt == TRIES:
                    raise HTTPException(
                        503,
                        f"Cannot reach Tally after {TRIES} tries. "
                        "Check the PC, port 9000 and that Tally is open.",
                    )
                time.sleep(RETRY_WAIT)


def send_to_tally(xml: str) -> ET.Element:
    text = post_to_tally(xml)
    text = re.sub(r"&#(?:[0-8]|1[1-2]|1[4-9]|2[0-9]|3[0-1]);", "", text)  # bad XML chars
    if not text.strip():
        raise HTTPException(502, "Tally returned an empty reply.")
    try:
        root = ET.fromstring(text)
    except ET.ParseError as e:
        raise HTTPException(
            502, f"Tally's reply was not valid XML ({e}). Start of reply: {text[:300]}"
        )
    for err in root.iter("LINEERROR"):
        raise HTTPException(502, f"Tally error: {err.text}")
    return root


def static_vars(from_date, to_date):
    return f"""<STATICVARIABLES>
      <SVCURRENTCOMPANY>{COMPANY}</SVCURRENTCOMPANY>
      <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      <SVFROMDATE>{from_date}</SVFROMDATE>
      <SVTODATE>{to_date}</SVTODATE>
    </STATICVARIABLES>"""


# ---------- SALES and RETURNS ----------
def voucher_request(type_formula, from_date, to_date):
    return f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE><ID>TallyVouchers</ID></HEADER>
  <BODY><DESC>
    {static_vars(from_date, to_date)}
    <TDL><TDLMESSAGE>
      <COLLECTION NAME="TallyVouchers" ISMODIFY="No">
        <TYPE>Voucher</TYPE>
        <FILTERS>VchTypeFilter</FILTERS>
        <FETCH>Date, VoucherTypeName, VoucherNumber, PartyLedgerName,
          AllInventoryEntries.StockItemName, AllInventoryEntries.BilledQty,
          AllInventoryEntries.Amount</FETCH>
      </COLLECTION>
      <SYSTEM TYPE="Formulae" NAME="VchTypeFilter">{type_formula}</SYSTEM>
    </TDLMESSAGE></TDL>
  </DESC></BODY>
</ENVELOPE>"""


def date_chunks(from_date, to_date, days=CHUNK_DAYS):
    """Cut a YYYYMMDD range into pieces of at most `days` days."""
    from_date, to_date = date_range(from_date, to_date)
    start = datetime.strptime(from_date, "%Y%m%d").date()
    end = datetime.strptime(to_date, "%Y%m%d").date()
    chunks = []
    while start <= end:
        piece_end = min(start + timedelta(days=days - 1), end)
        chunks.append((start.strftime("%Y%m%d"), piece_end.strftime("%Y%m%d")))
        start = piece_end + timedelta(days=1)
    return chunks


def sales_like(type_formula, from_date, to_date):
    """Fetch one chunk at a time, pausing between chunks so Tally is not overloaded."""
    chunks = date_chunks(from_date, to_date)
    rows = []
    for i, (f, t) in enumerate(chunks):
        if i:
            time.sleep(PAUSE_SECONDS)
        rows.extend(sales_chunk(type_formula, f, t))
    return rows


def sales_chunk(type_formula, from_date, to_date):
    root = send_to_tally(voucher_request(type_formula, from_date, to_date))
    rows = []
    for v in root.iter("VOUCHER"):
        for item in v:
            if item.tag.startswith(("ALLINVENTORYENTRIES", "INVENTORYENTRIES")):
                rows.append({
                    "Date": v.findtext("DATE"),
                    "Platform": v.findtext("PARTYLEDGERNAME"),
                    "SKU": item.findtext("STOCKITEMNAME"),
                    "Quantity": num(item.findtext("BILLEDQTY")),
                    "Value": num(item.findtext("AMOUNT")),
                })
    for r in rows:
        r["Date"] = datetime.strptime(r["Date"], "%Y%m%d").date()
        r["net sale"] = r["Value"] * GST_MULTIPLIER
    return rows


# ---------- STOCK (unverified: kept off by default) ----------
def stock_request(to_date):
    return f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Data</TYPE><ID>Stock Summary</ID></HEADER>
  <BODY><DESC>
    <STATICVARIABLES>
      <SVCURRENTCOMPANY>{COMPANY}</SVCURRENTCOMPANY>
      <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      <SVFROMDATE>{FY_START}</SVFROMDATE>
      <SVTODATE>{to_date}</SVTODATE>
      <EXPLODEFLAG>Yes</EXPLODEFLAG>
    </STATICVARIABLES>
  </DESC></BODY>
</ENVELOPE>"""


def stock_df(from_date, to_date):
    if not ENABLE_STOCK:
        raise HTTPException(
            403, "Stock is switched off (ENABLE_STOCK = False). Enable it only after sales/returns are confirmed."
        )
    root = send_to_tally(stock_request(to_date))
    rows, name = [], None
    for el in root.iter():
        if el.tag == "DSPDISPNAME":
            name = (el.text or "").strip()
        elif el.tag == "DSPCLQTY" and name:
            rows.append({"SKU": name, "Closing Qty": num(el.text), "Closing Value": 0.0})
        elif el.tag == "DSPCLAMTA" and rows:
            rows[-1]["Closing Value"] = num(el.text)
    return rows


COLUMNS = {
    "sales":   ["Date", "Platform", "SKU", "Quantity", "Value", "net sale"],
    "returns": ["Date", "Platform", "SKU", "Quantity", "Value", "net sale"],
    "stock":   ["SKU", "Closing Qty", "Closing Value"],
}


def save_excel(rows, columns, path):
    wb = Workbook()
    ws = wb.active
    ws.append(columns)
    for r in rows:
        ws.append([r.get(c) for c in columns])
    wb.save(path)


FETCHERS = {
    "sales":   lambda f, t: sales_like(SALES_FILTER, f, t),
    "returns": lambda f, t: sales_like(RETURNS_FILTER, f, t),
    "stock":   stock_df,
}


# ---------- saving + scheduling ----------
def export_all(from_date=None, to_date=None):
    from_date, to_date = date_range(from_date, to_date)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    saved, failed = [], {}
    for name, fetch in FETCHERS.items():
        if name == "stock" and not ENABLE_STOCK:
            continue
        try:
            rows = fetch(from_date, to_date)
            path = OUTPUT_DIR / f"{name}_{from_date}_{to_date}.xlsx"
            save_excel(rows, COLUMNS[name], path)
            saved.append(str(path))
        except Exception as e:   # one failure should not stop the others
            failed[name] = str(getattr(e, "detail", e))
    return {"saved": saved, "failed": failed}


@asynccontextmanager
async def lifespan(app):
    scheduler = BackgroundScheduler()
    scheduler.add_job(
        export_all, "cron", hour=RUN_AT_HOUR, minute=RUN_AT_MINUTE,
        max_instances=1, coalesce=True,
    )
    scheduler.start()
    yield
    scheduler.shutdown()


app = FastAPI(title="Tally API", lifespan=lifespan)


@app.get("/health")
def health():
    try:
        requests.get(TALLY_URL, timeout=5)
        return {"tally": "running"}
    except requests.RequestException:
        return {"tally": "not reachable"}


@app.get("/data/{kind}", dependencies=[Depends(require_key)])
def get_data(kind: str, from_date: str = None, to_date: str = None):
    """kind = sales | returns | stock. Dates as YYYYMMDD, default today."""
    if kind not in FETCHERS:
        raise HTTPException(404, "Use sales, returns or stock")
    f, t = date_range(from_date, to_date)
    return FETCHERS[kind](f, t)   # dates are sent as "YYYY-MM-DD"


@app.post("/export", dependencies=[Depends(require_key)])
def export_now(from_date: str = None, to_date: str = None):
    """Fetch the enabled reports and save them as Excel files right now."""
    return export_all(from_date, to_date)


@app.get("/debug/raw/{kind}", response_class=PlainTextResponse, dependencies=[Depends(require_key)])
def debug_raw(kind: str, from_date: str = None, to_date: str = None):
    """Shows the raw reply from Tally, to check tag names. Keep it to one day."""
    f, t = date_range(from_date, to_date)
    if kind == "sales":
        xml = voucher_request(SALES_FILTER, f, t)
    elif kind == "returns":
        xml = voucher_request(RETURNS_FILTER, f, t)
    elif kind == "stock":
        if not ENABLE_STOCK:
            raise HTTPException(403, "Stock is switched off (ENABLE_STOCK = False).")
        xml = stock_request(t)
    else:
        raise HTTPException(404, "Use sales, returns or stock")
    return post_to_tally(xml)[:20000]


if __name__ == "__main__":
    import uvicorn
    if not API_KEY:
        raise SystemExit('Set the key first:  setx TALLY_API_KEY "your-long-random-key"  (then open a new terminal)')
    uvicorn.run(app, host=LISTEN_HOST, port=LISTEN_PORT)