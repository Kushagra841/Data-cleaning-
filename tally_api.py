import re
import json
import time
import threading
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from pathlib import Path
from contextlib import asynccontextmanager
from fastapi import FastAPI, HTTPException
from fastapi.responses import PlainTextResponse
from apscheduler.schedulers.background import BackgroundScheduler

# ---------- settings ----------
# When this file runs ON the Tally PC, use localhost.
# (If you run it on the laptop instead, use "http://192.168.1.27:9000")
TALLY_URL = "http://192.168.1.27:9000"
COMPANY = "Shantam International UP (2026-27)"   # exactly as shown in Tally
FY_START = "20260401"                            # start of the financial year
OUTPUT_DIR = Path(r"C:\TallyData")               # a folder on the Tally PC
RUN_AT_HOUR, RUN_AT_MINUTE = 21, 0               # daily auto-fetch at 9:00 PM
GST_MULTIPLIER = 1.18                            # assumption: check against real tax amounts

# NEW: chunk settings. Long date ranges are cut into small pieces.
CHUNK_DAYS = 7          # ask Tally for 7 days at a time
PAUSE_BETWEEN = 3       # seconds to wait between pieces
RETRY_WAIT = 30         # seconds to wait if Tally does not answer
TALLY_TIMEOUT = 180     # seconds to wait for one reply

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


# ---------- NEW: date chunks ----------
def date_chunks(from_date, to_date, days=CHUNK_DAYS):
    """Cut a long date range into small pieces, e.g. 7 days each."""
    start = datetime.strptime(from_date, "%Y%m%d")
    end = datetime.strptime(to_date, "%Y%m%d")
    while start <= end:
        stop = min(start + timedelta(days=days - 1), end)
        yield start.strftime("%Y%m%d"), stop.strftime("%Y%m%d")
        start = stop + timedelta(days=1)


# ---------- CHANGED: patient request to Tally ----------
def post_to_tally(xml: str) -> str:
    """Send one request to Tally and return the raw reply text.
    If Tally does not answer, wait before trying again so we do not pile
    new requests on a busy Tally."""
    with tally_lock:
        for attempt in range(3):
            try:
                r = requests.post(
                    TALLY_URL, data=xml.encode("utf-8"), timeout=TALLY_TIMEOUT
                )
                return r.content.decode("utf-8", errors="ignore")
            except requests.RequestException:
                time.sleep(RETRY_WAIT)
        raise HTTPException(
            503, "Cannot reach Tally. Check the PC, port 9000 and that Tally is open."
        )


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


def _sales_chunk(type_formula, from_date, to_date):
    """Fetch ONE small piece (e.g. 7 days) from Tally."""
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
    df = pd.DataFrame(rows, columns=["Date", "Platform", "SKU", "Quantity", "Value"])
    if df.empty:
        return df
    df["Date"] = pd.to_datetime(df["Date"], format="%Y%m%d")
    df["net sale"] = df["Value"] * GST_MULTIPLIER
    return df


def sales_like(type_formula, from_date, to_date):
    """NEW: loop over small pieces, then join them into one table."""
    frames = []
    for f, t in date_chunks(from_date, to_date):
        df = _sales_chunk(type_formula, f, t)
        if not df.empty:
            frames.append(df)
        time.sleep(PAUSE_BETWEEN)          # give Tally a breather
    cols = ["Date", "Platform", "SKU", "Quantity", "Value", "net sale"]
    if not frames:
        return pd.DataFrame(columns=cols)
    return pd.concat(frames, ignore_index=True)


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
    return pd.DataFrame(rows, columns=["SKU", "Closing Qty", "Closing Value"])


FETCHERS = {
    "sales":   lambda f, t: sales_like(SALES_FILTER, f, t),
    "returns": lambda f, t: sales_like(RETURNS_FILTER, f, t),
    "stock":   stock_df,
}


# ---------- saving + scheduling ----------
def export_all(from_date=None, to_date=None):
    today = date.today().strftime("%Y%m%d")
    from_date, to_date = from_date or today, to_date or today
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    saved, failed = [], {}
    for name, fetch in FETCHERS.items():
        if name == "stock" and not ENABLE_STOCK:
            continue
        try:
            df = fetch(from_date, to_date)
            path = OUTPUT_DIR / f"{name}_{from_date}_{to_date}.xlsx"
            df.to_excel(path, index=False)
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
    # Never send a second request while one is running: if the lock is taken,
    # Tally is busy with an export, so just say so instead of adding load.
    if not tally_lock.acquire(blocking=False):
        return {"tally": "busy (export in progress)"}
    try:
        requests.get(TALLY_URL, timeout=5)
        return {"tally": "running"}
    except requests.RequestException:
        return {"tally": "not reachable"}
    finally:
        tally_lock.release()


@app.get("/data/{kind}")
def get_data(kind: str, from_date: str = None, to_date: str = None):
    """kind = sales | returns | stock. Dates as YYYYMMDD, default today."""
    if kind not in FETCHERS:
        raise HTTPException(404, "Use sales, returns or stock")
    today = date.today().strftime("%Y%m%d")
    df = FETCHERS[kind](from_date or today, to_date or today)
    return json.loads(df.to_json(orient="records", date_format="iso"))


@app.post("/export")
def export_now(from_date: str = None, to_date: str = None):
    """Fetch the enabled reports and save them as Excel files right now."""
    return export_all(from_date, to_date)


@app.get("/debug/raw/{kind}", response_class=PlainTextResponse)
def debug_raw(kind: str, from_date: str = None, to_date: str = None):
    """Shows the raw reply from Tally, to check tag names. Use ONE day only."""
    today = date.today().strftime("%Y%m%d")
    f, t = from_date or today, to_date or today
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
    uvicorn.run(app, host="0.0.0.0", port=8000)