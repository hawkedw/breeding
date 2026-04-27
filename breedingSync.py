# breedingSync.py
# import_registry -> attaches to open workbook via GetActiveObject (like submit_registry)
# submit_registry -> reads dirty rows directly from workbook via win32com, sends to ArcGIS

import sys, os, json, datetime
import requests

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(BASE_DIR, "breedingSync.log")


def log(msg: str):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    try:
        print(line)
    except Exception:
        pass
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


# ---------- CONSTANTS ----------

SHEET_REGISTRY = "\u0420\u0435\u0435\u0441\u0442\u0440 \u043e\u043f\u044b\u0442\u043e\u0432"
DIRTY_ALIAS    = "Dirty"

URL_PARENT = "https://maps.ekoniva-apk.org/arcgis/rest/services/breeding/breeding/FeatureServer/0"
URL_CHILD  = "https://maps.ekoniva-apk.org/arcgis/rest/services/breeding/breeding/FeatureServer/1"
SORT_FIELD = "created_date"

FIELDS_PARENT = [
    {"n": "country",          "alias": "\u0421\u0442\u0440\u0430\u043d\u0430",                       "type": "TEXT",   "col": 1},
    {"n": "region",           "alias": "\u0420\u0435\u0433\u0438\u043e\u043d",                       "type": "TEXT",   "col": 2},
    {"n": "site",             "alias": "\u041e\u043f\u044b\u0442\u043d\u0430\u044f \u043f\u043b\u043e\u0449\u0430\u0434\u043a\u0430",             "type": "TEXT",   "col": 3},
    {"n": "devision",         "alias": "\u041e\u0442\u0434\u0435\u043b\u0435\u043d\u0438\u0435 \u0426\u0421\u0438\u041f\u0421",              "type": "TEXT",   "col": 4},
    {"n": "crop",             "alias": "\u041a\u0443\u043b\u044c\u0442\u0443\u0440\u0430",                     "type": "TEXT",   "col": 5},
    {"n": "farm",             "alias": "\u0425\u043e\u0437\u044f\u0439\u0441\u0442\u0432\u043e (\u043f\u043e\u0434\u0440\u0430\u0437\u0434\u0435\u043b\u0435\u043d\u0438\u0435 \u041f\u0425)", "type": "TEXT",   "col": 6},
    {"n": "responsable",      "alias": "\u041e\u0442\u0432. \u041b\u0438\u0446\u043e \u0432 \u041f\u0425",               "type": "TEXT",   "col": 7},
    {"n": "fieldNumber",      "alias": "\u041d\u043e\u043c\u0435\u0440 \u043f\u043e\u043b\u044f",                   "type": "TEXT",   "col": 8},
    {"n": "areaHa",           "alias": "\u041f\u043b\u043e\u0449\u0430\u0434\u044c \u043e\u043f\u044b\u0442\u0430, \u0433\u0430",            "type": "NUMBER", "col": 9},
    {"n": "scheme",           "alias": "\u0421\u0445\u0435\u043c\u0430 \u043e\u043f\u044b\u0442\u0430",                  "type": "TEXT",   "col": 10},
    {"n": "experimentName",   "alias": "\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435 \u043e\u043f\u044b\u0442\u0430",               "type": "TEXT",   "col": 11},
    {"n": "type",             "alias": "\u0422\u0438\u043f \u043e\u043f\u044b\u0442\u0430",                    "type": "TEXT",   "col": 12},
    {"n": "productPurpose",   "alias": "\u041d\u0430\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435 \u043f\u0440\u043e\u0434\u0443\u043a\u0446\u0438\u0438 \u043e\u043f\u044b\u0442\u0430",   "type": "TEXT",   "col": 13},
    {"n": "trialPurpose",     "alias": "\u0426\u0435\u043b\u044c, \u0437\u0430\u0434\u0430\u0447\u0430 \u043e\u043f\u044b\u0442\u0430",           "type": "TEXT",   "col": 14},
    {"n": "status",           "alias": "\u0421\u0442\u0430\u0442\u0443\u0441 \u043e\u043f\u044b\u0442\u0430",                 "type": "TEXT",   "col": 15},
    {"n": "plantingDate",     "alias": "\u0414\u0430\u0442\u0430 \u043f\u043e\u0441\u0435\u0432\u0430",                  "type": "DATE",   "col": 16},
    {"n": "haverstDate",      "alias": "\u0414\u0430\u0442\u0430 \u0443\u0431\u043e\u0440\u043a\u0438",                    "type": "DATE",   "col": 17},
    {"n": "report",           "alias": "\u041e\u0442\u0447\u0451\u0442 (\u0412\u044b\u0432\u043e\u0434\u044b, \u0440\u0435\u043a\u043e\u043c\u0435\u043d\u0434\u0430\u0446\u0438\u0438)", "type": "TEXT",   "col": 18},
    {"n": "created_date",     "alias": "created_date",                 "type": "DATE",   "col": 27},
    {"n": "last_edited_date", "alias": "last_edited_date",             "type": "DATE",   "col": 28},
]

CUSTOMER_COLS  = [20, 21, 22, 23]
CUSTOMER_FIELD = "customer"
CUSTOMER_ALIAS = "\u0417\u0430\u043a\u0430\u0437\u0447\u0438\u043a \u043e\u043f\u044b\u0442\u0430"

DIRTY_COL      = 29
PARENT_GID_COL = 30
CHILD_GID_COL  = 31
TOTAL_COLS     = CHILD_GID_COL

EDITABLE_COLS    = set(range(3, 19))
SYS_SKIP         = {"created_user", "created_date", "last_edited_user", "last_edited_date"}
DATE_ONLY_FIELDS = {"plantingDate", "haverstDate"}

ALIAS_TO_NAME = {f["alias"]: f["n"] for f in FIELDS_PARENT}
ALIAS_TO_NAME[CUSTOMER_ALIAS] = CUSTOMER_FIELD

# ---------- AUTH ----------

TOKEN_URL        = "https://maps.ekoniva-apk.org/portal/sharing/rest/generateToken"
ARC_USERNAME_ENV = "ARCGIS_BREEDING_USER"
ARC_PASSWORD_ENV = "ARCGIS_BREEDING_PASS"


def get_token() -> str:
    username = os.environ.get(ARC_USERNAME_ENV)
    password = os.environ.get(ARC_PASSWORD_ENV)
    if not username or not password:
        raise RuntimeError(f"\u041d\u0435 \u0437\u0430\u0434\u0430\u043d\u044b {ARC_USERNAME_ENV}/{ARC_PASSWORD_ENV}")
    payload = {
        "username": username, "password": password,
        "client": "referer", "referer": "https://maps.ekoniva-apk.org",
        "expiration": 60, "f": "json",
    }
    resp = requests.post(TOKEN_URL, data=payload, timeout=30)
    js = resp.json()
    tok = js.get("token")
    if not tok:
        raise RuntimeError(f"Token error: {js}")
    return tok


# ---------- DATE HELPERS ----------

EPOCH       = datetime.datetime(1970, 1, 1)
OFFSET      = datetime.timedelta(hours=3)   # MSK = UTC+3
EXCEL_EPOCH = datetime.datetime(1899, 12, 30)

_FMT_DATE     = "dd.mm.yyyy"
_FMT_DATETIME = "dd.mm.yyyy hh:mm"


def esri_ms_to_dt(ms):
    return EPOCH + datetime.timedelta(milliseconds=int(ms)) + OFFSET


def dt_to_excel_serial(dt: datetime.datetime) -> float:
    if dt.tzinfo is not None:
        dt = dt.astimezone(datetime.timezone.utc).replace(tzinfo=None)
    delta = dt - EXCEL_EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86400.0


def arc_ms_to_excel_serial(ms, date_only: bool = False) -> float:
    if date_only:
        d = (EPOCH + datetime.timedelta(milliseconds=int(ms))).date()
        dt = datetime.datetime(d.year, d.month, d.day)
    else:
        dt = esri_ms_to_dt(int(ms))
    return dt_to_excel_serial(dt)


def dt_to_esri(dt: datetime.datetime) -> int:
    if dt.tzinfo is not None:
        dt = dt.astimezone(datetime.timezone.utc).replace(tzinfo=None)
    else:
        dt = dt - OFFSET
    return int((dt - EPOCH).total_seconds() * 1000)


def date_to_esri(d: datetime.date) -> int:
    dt = datetime.datetime(d.year, d.month, d.day)
    return int((dt - EPOCH).total_seconds() * 1000)


def excel_serial_to_dt(x: float) -> datetime.datetime:
    return EXCEL_EPOCH + datetime.timedelta(days=float(x))


def _set_number_format(rng, fmt: str):
    try:
        rng.NumberFormat = fmt
    except Exception as e:
        log(f"NumberFormat failed ({e})")


# ---------- QUERY LAYER ----------

def query_layer(url, where="1=1", order_by=""):
    token = get_token()
    session = requests.Session()
    feats, offset, page_size = [], 0, 2000
    while True:
        params = {"where": where, "outFields": "*", "f": "json",
                  "token": token, "resultOffset": offset, "resultRecordCount": page_size}
        if order_by:
            params["orderByFields"] = f"{order_by} DESC"
        r = session.get(url + "/query", params=params, timeout=60)
        js = r.json()
        if "error" in js:
            raise RuntimeError(f"ArcGIS error: {js['error']}")
        page = js.get("features", [])
        feats.extend(page)
        log(f"query_layer offset={offset}: got {len(page)}, total {len(feats)}")
        if not js.get("exceededTransferLimit") or not page:
            break
        offset += len(page)
    return feats


# ---------- ATTACH WORKBOOK ----------

def _attach_workbook(wb_path: str):
    import win32com.client as win32
    import time

    last_err = None
    for attempt in range(10):
        try:
            xl = win32.GetActiveObject("Excel.Application")
            break
        except Exception as e:
            last_err = e
            log(f"GetActiveObject attempt {attempt+1} failed: {e}, retrying...")
            time.sleep(1)
    else:
        raise RuntimeError(f"Cannot attach to Excel: {last_err}")

    target_full = wb_path.lower()
    target_name = os.path.basename(wb_path).lower()

    wb = None
    by_name = None
    for book in xl.Workbooks:
        try:
            full = book.FullName.lower()
            log(f"  checking open workbook: {book.FullName}")
            if full == target_full:
                wb = book
                break
            if os.path.basename(full) == target_name and by_name is None:
                by_name = book
        except Exception:
            continue

    if wb is None and by_name is not None:
        log(f"Exact path not matched, using filename match: {by_name.FullName}")
        wb = by_name

    if wb is None:
        names = [b.FullName for b in xl.Workbooks]
        raise RuntimeError(f"Workbook not found.\nExpected: {wb_path}\nOpen: {names}")

    log(f"Found open workbook: {wb.FullName}")
    return xl, wb


# ---------- IMPORT ----------

def _to_2d(rows):
    return tuple(tuple(r) for r in rows)


def import_registry(wb_path: str):
    log("=== import_registry START ===")

    try:
        import win32com.client as win32
    except ImportError:
        log("ERROR: pywin32 not installed")
        return 1

    parent_feats = query_layer(URL_PARENT, "1=1", SORT_FIELD)
    log(f"Parent: {len(parent_feats)} features")
    parent_feats.sort(
        key=lambda f: f.get("attributes", {}).get(SORT_FIELD) or 0,
        reverse=True
    )

    child_feats = query_layer(URL_CHILD, "1=1", "")
    log(f"Child: {len(child_feats)} records")

    child_index = {}
    for cf in child_feats:
        attrs = cf.get("attributes", {})
        pgid = attrs.get("parentglobalid")
        if pgid:
            child_index.setdefault(pgid, []).append(attrs.get(CUSTOMER_FIELD))

    headers = [""] * TOTAL_COLS
    for f in FIELDS_PARENT:
        col = f.get("col")
        if col:
            headers[col - 1] = f["alias"]
    for c in CUSTOMER_COLS:
        headers[c - 1] = CUSTOMER_ALIAS
    headers[DIRTY_COL - 1]      = DIRTY_ALIAS
    headers[PARENT_GID_COL - 1] = "GlobalID"
    headers[CHILD_GID_COL - 1]  = "ChildGlobalID"

    data = []
    for ft in parent_feats:
        attrs = ft.get("attributes", {})
        row = [None] * TOTAL_COLS

        for f in FIELDS_PARENT:
            col = f.get("col")
            if not col:
                continue
            v = attrs.get(f["n"])
            if v is None:
                row[col - 1] = ""
                continue
            if f["type"] == "DATE":
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    row[col - 1] = float(arc_ms_to_excel_serial(
                        v, date_only=(f["n"] in DATE_ONLY_FIELDS)
                    ))
                else:
                    row[col - 1] = ""
            else:
                row[col - 1] = v if v is not None else ""

        parent_gid = attrs.get("GlobalID", "")
        customers  = child_index.get(parent_gid, [])
        child_gid  = ""
        if customers:
            for cf in child_feats:
                if cf.get("attributes", {}).get("parentglobalid") == parent_gid:
                    child_gid = cf.get("attributes", {}).get("GlobalID", "")
                    break

        for i, c in enumerate(CUSTOMER_COLS):
            row[c - 1] = customers[i] if i < len(customers) and customers[i] is not None else ""

        row[DIRTY_COL - 1]      = False
        row[PARENT_GID_COL - 1] = parent_gid
        row[CHILD_GID_COL - 1]  = child_gid
        data.append(row)

    log(f"Data ready: {len(data)} rows. Attaching to Excel...")

    xl, wb = _attach_workbook(wb_path)

    xl.ScreenUpdating = False
    xl.Calculation    = -4135  # xlCalculationManual
    xl.EnableEvents   = False

    try:
        try:
            sh = wb.Worksheets(SHEET_REGISTRY)
            sh.Cells.Clear()
        except Exception:
            sh = wb.Worksheets.Add()
            sh.Name = SHEET_REGISTRY

        # header
        sh.Range(sh.Cells(1, 1), sh.Cells(1, TOTAL_COLS)).Value = _to_2d([headers])

        if data:
            n = len(data)

            # 1) NumberFormat BEFORE writing values — prevents Excel from overriding with General
            for f in FIELDS_PARENT:
                col = f.get("col")
                if not col or f["type"] != "DATE":
                    continue
                fmt = _FMT_DATE if f["n"] in DATE_ONLY_FIELDS else _FMT_DATETIME
                rng = sh.Range(sh.Cells(2, col), sh.Cells(1 + n, col))
                _set_number_format(rng, fmt)
                log(f"  col {col} '{f['n']}' -> NumberFormat='{fmt}'")

            # 2) write data
            sh.Range(sh.Cells(2, 1), sh.Cells(1 + n, TOTAL_COLS)).Value = _to_2d(data)

        wb.Save()
        log(f"import_registry complete: {len(data)} rows written to {wb.FullName}")

    finally:
        xl.Calculation    = -4105  # xlCalculationAutomatic
        xl.ScreenUpdating = True
        xl.EnableEvents   = True

    return 0


# ---------- SUBMIT ----------

def submit_registry(wb_path: str):
    log("=== submit_registry START ===")

    try:
        import win32com.client as win32
    except ImportError:
        log("ERROR: pywin32 not installed")
        return 1

    xl, wb = _attach_workbook(wb_path)

    try:
        sh = wb.Worksheets(SHEET_REGISTRY)
    except Exception:
        log(f"ERROR: sheet '{SHEET_REGISTRY}' not found")
        return 1

    last_col = sh.Cells(1, sh.Columns.Count).End(-4159).Column
    last_row = sh.Cells(sh.Rows.Count, 1).End(-4162).Row

    if last_row < 2:
        log("No data rows")
        return 0

    hdr_vals = list(sh.Range(sh.Cells(1, 1), sh.Cells(1, last_col)).Value[0])

    def col_idx(name):
        for i, h in enumerate(hdr_vals):
            if h == name:
                return i + 1
        return 0

    dirty_col = col_idx(DIRTY_ALIAS)
    gid_col   = col_idx("GlobalID")

    if not dirty_col:
        log("ERROR: 'Dirty' column not found")
        return 1
    if not gid_col:
        log("ERROR: 'GlobalID' column not found")
        return 1

    last_row = max(
        last_row,
        sh.Cells(sh.Rows.Count, dirty_col).End(-4162).Row,
        sh.Cells(sh.Rows.Count, gid_col).End(-4162).Row,
    )

    data = sh.Range(sh.Cells(2, 1), sh.Cells(last_row, last_col)).Value
    if not data:
        log("No data")
        return 0

    name_to_type = {f["n"]: f["type"] for f in FIELDS_PARENT}
    token   = get_token()
    edits   = []
    row_map = []

    for r_idx, row in enumerate(data, start=2):
        row = list(row)
        if not row[dirty_col - 1]:
            continue

        parent_gid = row[gid_col - 1]
        if not parent_gid:
            log(f"Row {r_idx}: no GlobalID, skip")
            continue

        attrs = {"GlobalID": str(parent_gid).strip()}

        for c_idx, alias in enumerate(hdr_vals, start=1):
            if not alias or alias == DIRTY_ALIAS:
                continue
            if alias in ("GlobalID", "ChildGlobalID"):
                continue
            field_name = ALIAS_TO_NAME.get(alias, alias)
            if field_name.lower() in SYS_SKIP:
                continue
            f_type = name_to_type.get(field_name)
            if f_type is None:
                continue

            v = row[c_idx - 1]

            if v in ("", None):
                attrs[field_name] = None
            elif f_type == "DATE":
                date_only = field_name in DATE_ONLY_FIELDS
                if isinstance(v, datetime.date) and not isinstance(v, datetime.datetime):
                    attrs[field_name] = date_to_esri(v)
                elif isinstance(v, datetime.datetime):
                    attrs[field_name] = date_to_esri(v.date()) if date_only else dt_to_esri(v)
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    dt = excel_serial_to_dt(float(v))
                    attrs[field_name] = date_to_esri(dt.date()) if date_only else dt_to_esri(dt)
                else:
                    attrs[field_name] = None
            elif f_type == "NUMBER":
                try:
                    attrs[field_name] = float(v)
                except Exception:
                    attrs[field_name] = None
            else:
                attrs[field_name] = str(v) if v is not None else None

        edits.append({"attributes": attrs})
        row_map.append((r_idx, dirty_col))

    log(f"Dirty rows found: {len(edits)}")
    if not edits:
        log("No dirty rows")
        return 0

    feats_json = json.dumps([{"attributes": e["attributes"]} for e in edits])
    resp = requests.post(URL_PARENT + "/applyEdits", data={
        "f": "json", "token": token,
        "rollbackOnFailure": "True", "useGlobalIds": "True",
        "updates": feats_json,
    }, timeout=60)
    js = resp.json()
    log(f"applyEdits response: {json.dumps(js, ensure_ascii=False)[:500]}")

    if "error" in js:
        log(f"applyEdits error: {js['error']}")
        return 1

    for (excel_row, d_col), r in zip(row_map, js.get("updateResults", [])):
        if r.get("success"):
            sh.Cells(excel_row, d_col).Value = False
            log(f"Row {excel_row}: OK")
        else:
            log(f"Row {excel_row}: FAILED - {r.get('error', {}).get('description', '')}")

    wb.Save()
    log("submit_registry complete")
    return 0


# ---------- MAIN ----------

def normalize_action(a):
    a = (a or "").strip()
    if a.lower().startswith("action="):
        a = a.split("=", 1)[1].strip()
    return a.lower()


def main(argv=None):
    if argv is None:
        argv = sys.argv
    if len(argv) < 3:
        log("Usage: breedingSync.py <action> <workbook_path>")
        return 1

    action  = normalize_action(argv[1])
    wb_path = argv[2]

    log("=== breedingSync START ===")
    log(f"action={action!r}  workbook={wb_path}")
    log(f"python={sys.executable}  cwd={os.getcwd()}")

    if action == "import_registry":
        return import_registry(wb_path) or 0

    if action == "submit_registry":
        return submit_registry(wb_path) or 0

    log(f"Unknown action: {action!r}. Available: import_registry, submit_registry")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
