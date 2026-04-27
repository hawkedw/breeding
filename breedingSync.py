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

SHEET_REGISTRY = "Реестр опытов"
DIRTY_ALIAS    = "Dirty"

URL_PARENT = "https://maps.ekoniva-apk.org/arcgis/rest/services/breeding/breeding/FeatureServer/0"
URL_CHILD  = "https://maps.ekoniva-apk.org/arcgis/rest/services/breeding/breeding/FeatureServer/1"
SORT_FIELD = "created_date"

FIELDS_PARENT = [
    {"n": "country",          "alias": "Страна",                               "type": "TEXT",   "col": 1},
    {"n": "region",           "alias": "Регион",                               "type": "TEXT",   "col": 2},
    {"n": "site",             "alias": "Опытная площадка",                     "type": "TEXT",   "col": 3},
    {"n": "devision",         "alias": "Отделение ЦСиПС",                      "type": "TEXT",   "col": 4},
    {"n": "crop",             "alias": "Культура",                             "type": "TEXT",   "col": 5},
    {"n": "farm",             "alias": "Хозяйство (подразделение ПХ)",         "type": "TEXT",   "col": 6},
    {"n": "responsable",      "alias": "Отв. Лицо в ПХ",                       "type": "TEXT",   "col": 7},
    {"n": "fieldNumber",      "alias": "Номер поля",                           "type": "TEXT",   "col": 8},
    {"n": "areaHa",           "alias": "Площадь опыта, га",                    "type": "NUMBER", "col": 9},
    {"n": "scheme",           "alias": "Схема опыта",                          "type": "TEXT",   "col": 10},
    {"n": "experimentName",   "alias": "Название опыта",                       "type": "TEXT",   "col": 11},
    {"n": "type",             "alias": "Тип опыта",                            "type": "TEXT",   "col": 12},
    {"n": "productPurpose",   "alias": "Назначение продукции опыта",           "type": "TEXT",   "col": 13},
    {"n": "trialPurpose",     "alias": "Цель, задача опыта",                   "type": "TEXT",   "col": 14},
    {"n": "status",           "alias": "Статус опыта",                         "type": "TEXT",   "col": 15},
    {"n": "plantingDate",     "alias": "Дата посева",                          "type": "DATE",   "col": 16},
    {"n": "haverstDate",      "alias": "Дата уборки",                          "type": "DATE",   "col": 17},
    {"n": "report",           "alias": "Отчёт (Выводы, рекомендации)",         "type": "TEXT",   "col": 18},
    {"n": "created_date",     "alias": "created_date",                         "type": "DATE",   "col": 27},
    {"n": "last_edited_date", "alias": "last_edited_date",                     "type": "DATE",   "col": 28},
]

CUSTOMER_COLS  = [20, 21, 22, 23]
CUSTOMER_FIELD = "customer"
CUSTOMER_ALIAS = "Заказчик опыта"

DIRTY_COL          = 29
PARENT_GID_COL     = 30
CHILD_GID_COL      = 31
PARENT_OID_COL     = 32   # новая техколонка
CHILD_OID_COL      = 33   # новая техколонка
TOTAL_COLS         = CHILD_OID_COL

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
        raise RuntimeError(f"Не заданы {ARC_USERNAME_ENV}/{ARC_PASSWORD_ENV}")
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


# ---------- QUERY LAYER ----------

def query_layer(url, where="1=1", order_by=""):
    token = get_token()
    session = requests.Session()
    feats, offset, page_size = [], 0, 2000
    while True:
        params = {"where": where, "outFields": "*", "f": "json",
                  "resultOffset": offset, "resultRecordCount": page_size,
                  "token": token}
        if order_by:
            params["orderByFields"] = order_by
        r = session.get(url + "/query", params=params, timeout=60)
        r.raise_for_status()
        js = r.json()
        if "error" in js:
            raise RuntimeError(js["error"])
        chunk = js.get("features", [])
        feats.extend(chunk)
        log(f"query_layer offset={offset}: got {len(chunk)}, total {len(feats)}")
        if len(chunk) < page_size:
            break
        offset += page_size
    return feats


# ---------- HELPERS ----------

def _to_2d(rows):
    return [list(r) for r in rows]


def _attach_workbook(wb_path: str):
    import win32com.client as win32
    xl = win32.Dispatch("Excel.Application")
    abs_path = os.path.abspath(wb_path)
    for wb in xl.Workbooks:
        log(f"  checking open workbook: {wb.FullName}")
        if os.path.abspath(wb.FullName) == abs_path:
            log(f"Found open workbook: {wb.FullName}")
            return wb, xl
    raise RuntimeError(f"Книга не открыта в Excel: {wb_path}")


def _get_oid_field(feats: list) -> str:
    """Detect objectid field name from first feature (objectid / OBJECTID / FID)."""
    if not feats:
        return "objectid"
    a = feats[0].get("attributes", {})
    for name in ("objectid", "OBJECTID", "FID", "fid"):
        if name in a:
            return name
    # fallback: first int-looking key that's not globalid
    for k, v in a.items():
        if isinstance(v, int) and "global" not in k.lower():
            return k
    return "objectid"


# ---------- IMPORT ----------

def import_registry(wb_path: str):
    log("=== import_registry START ===")

    parent_feats = query_layer(URL_PARENT, order_by=SORT_FIELD)
    log(f"Parent: {len(parent_feats)} features")

    child_feats = query_layer(URL_CHILD)
    log(f"Child: {len(child_feats)} records")

    parent_oid_field = _get_oid_field(parent_feats)
    child_oid_field  = _get_oid_field(child_feats)
    log(f"OID fields: parent='{parent_oid_field}' child='{child_oid_field}'")

    # child lookup: parent_globalid -> list of (customer, globalid, objectid)
    child_map: dict[str, list] = {}
    for ft in child_feats:
        a = ft.get("attributes", {})
        pgid = a.get("parentglobalid") or a.get("ParentGlobalID") or a.get("parent_globalid")
        cval = a.get(CUSTOMER_FIELD, "") or ""
        cgid = a.get("globalid") or a.get("GlobalID") or ""
        coid = a.get(child_oid_field)
        if pgid:
            child_map.setdefault(pgid, []).append((cval, cgid, coid))

    # build headers
    col_map = {f["col"]: f["alias"] for f in FIELDS_PARENT}
    headers = []
    for c in range(1, TOTAL_COLS + 1):
        if c in col_map:
            headers.append(col_map[c])
        elif c in CUSTOMER_COLS:
            headers.append(CUSTOMER_ALIAS)
        elif c == DIRTY_COL:
            headers.append(DIRTY_ALIAS)
        elif c == PARENT_GID_COL:
            headers.append("parent_globalid")
        elif c == CHILD_GID_COL:
            headers.append("child_globalid")
        elif c == PARENT_OID_COL:
            headers.append("parent_objectid")
        elif c == CHILD_OID_COL:
            headers.append("child_objectid")
        else:
            headers.append("")

    # build data rows
    data = []
    date_log_done = set()
    for ft in parent_feats:
        a = ft.get("attributes", {})
        row = [""] * TOTAL_COLS

        for f in FIELDS_PARENT:
            col = f.get("col")
            if not col:
                continue
            v = a.get(f["n"])
            if v is None:
                continue
            if f["type"] == "DATE" and isinstance(v, (int, float)):
                date_only = f["n"] in DATE_ONLY_FIELDS
                serial = arc_ms_to_excel_serial(v, date_only=date_only)
                if f["n"] not in date_log_done:
                    log(f"  DATE '{f['n']}': ms={v} -> serial={serial:.4f} type={type(serial).__name__}")
                    date_log_done.add(f["n"])
                row[col - 1] = serial
            else:
                row[col - 1] = v

        pgid = a.get("globalid") or a.get("GlobalID") or ""
        poid = a.get(parent_oid_field)

        # customers (first child only for now — each customer col = separate child)
        children = child_map.get(pgid, [])
        for i, cc in enumerate(CUSTOMER_COLS):
            row[cc - 1] = children[i][0] if i < len(children) else ""

        # first child gid/oid
        row[CHILD_GID_COL - 1] = children[0][1] if children else ""
        row[CHILD_OID_COL - 1] = children[0][2] if children else ""

        row[DIRTY_COL - 1]      = False
        row[PARENT_GID_COL - 1] = pgid
        row[PARENT_OID_COL - 1] = poid
        data.append(row)

    log(f"Data ready: {len(data)} rows. Attaching to Excel...")
    wb, xl = _attach_workbook(wb_path)

    xl.ScreenUpdating = False
    xl.EnableEvents   = False
    xl.Calculation    = -4135  # xlCalculationManual

    try:
        try:
            sh = wb.Worksheets(SHEET_REGISTRY)
            sh.Cells.Clear()
        except Exception:
            sh = wb.Worksheets.Add()
            sh.Name = SHEET_REGISTRY

        sh.Range(sh.Cells(1, 1), sh.Cells(1, TOTAL_COLS)).Value = _to_2d([headers])

        if data:
            n = len(data)
            sh.Range(sh.Cells(2, 1), sh.Cells(1 + n, TOTAL_COLS)).Value = _to_2d(data)

            for f in FIELDS_PARENT:
                col = f.get("col")
                if not col or f["type"] != "DATE":
                    continue
                date_only = f["n"] in DATE_ONLY_FIELDS
                fmt       = _FMT_DATE     if date_only else _FMT_DATETIME
                fmt_local = "ДД.ММ.ГГГГ" if date_only else "ДД.ММ.ГГГГ чч:мм"
                rng = sh.Range(sh.Cells(2, col), sh.Cells(1 + n, col))
                try:
                    rng.NumberFormat = fmt
                except Exception as e:
                    log(f"  col {col} NumberFormat failed: {e}")
                try:
                    rng.NumberFormatLocal = fmt_local
                except Exception:
                    pass
                applied    = rng.NumberFormat
                first_text = sh.Cells(2, col).Text
                log(f"  col {col} '{f['n']}' -> NumberFormat='{applied}' cell='{first_text}'")

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
    wb, xl = _attach_workbook(wb_path)

    try:
        sh = wb.Worksheets(SHEET_REGISTRY)
    except Exception:
        raise RuntimeError(f"Лист '{SHEET_REGISTRY}' не найден")

    last_col = TOTAL_COLS
    last_row = sh.Cells(sh.Rows.Count, 1).End(-4162).Row
    if last_row < 2:
        log("Нет данных для отправки")
        return 0

    hdr_vals = list(sh.Range(sh.Cells(1, 1), sh.Cells(1, last_col)).Value[0])
    col_idx: dict[str, int] = {}
    for i, h in enumerate(hdr_vals):
        if h and str(h).strip():
            col_idx[str(h).strip()] = i

    dirty_i      = col_idx.get(DIRTY_ALIAS)
    pgid_i       = col_idx.get("parent_globalid")
    poid_i       = col_idx.get("parent_objectid")
    cgid_i       = col_idx.get("child_globalid")
    coid_i       = col_idx.get("child_objectid")

    if dirty_i is None:
        raise RuntimeError(f"Колонка '{DIRTY_ALIAS}' не найдена")

    data = sh.Range(sh.Cells(2, 1), sh.Cells(last_row, last_col)).Value
    if not data:
        log("Нет строк данных")
        return 0

    token = get_token()
    updates_parent, adds_parent = [], []
    updates_child,  adds_child  = [], []

    for row_data in data:
        row = list(row_data)
        if not row[dirty_i]:
            continue

        pgid = row[pgid_i] if pgid_i is not None else None
        poid = row[poid_i] if poid_i is not None else None
        cgid = row[cgid_i] if cgid_i is not None else None
        coid = row[coid_i] if coid_i is not None else None

        p_attrs: dict = {}
        c_attrs: dict = {}

        for alias, fi in col_idx.items():
            field_name = ALIAS_TO_NAME.get(alias)
            if not field_name:
                continue
            if field_name in SYS_SKIP:
                continue

            excel_col = fi + 1
            if excel_col not in EDITABLE_COLS:
                continue

            raw    = row[fi]
            f_type = next((f["type"] for f in FIELDS_PARENT if f["n"] == field_name), "TEXT")

            if f_type == "DATE":
                date_only = field_name in DATE_ONLY_FIELDS
                if raw is None or raw == "":
                    val = None
                elif isinstance(raw, (int, float)):
                    dt = excel_serial_to_dt(float(raw))
                    val = date_to_esri(dt.date()) if date_only else dt_to_esri(dt)
                elif isinstance(raw, datetime.datetime):
                    val = date_to_esri(raw.date()) if date_only else dt_to_esri(raw)
                elif isinstance(raw, datetime.date):
                    val = date_to_esri(raw)
                else:
                    val = None
            elif f_type == "NUMBER":
                val = float(raw) if raw not in (None, "") else None
            elif f_type in ("INT", "OID"):
                val = int(raw) if raw not in (None, "") else None
            else:
                val = str(raw) if raw not in (None, "") else None

            if field_name == CUSTOMER_FIELD:
                c_attrs[field_name] = val
            else:
                p_attrs[field_name] = val

        # parent: update requires objectid; add uses negative placeholder
        if pgid and poid is not None:
            p_attrs["objectid"] = int(poid)
            p_attrs["globalid"] = pgid
            updates_parent.append({"attributes": p_attrs})
        else:
            adds_parent.append({"attributes": p_attrs})

        if c_attrs:
            if cgid and coid is not None:
                c_attrs["objectid"] = int(coid)
                c_attrs["globalid"] = cgid
                updates_child.append({"attributes": c_attrs})
            else:
                if pgid:
                    c_attrs["parentglobalid"] = pgid
                adds_child.append({"attributes": c_attrs})

    log(f"Dirty rows: parent updates={len(updates_parent)} adds={len(adds_parent)}, "
        f"child updates={len(updates_child)} adds={len(adds_child)}")

    session = requests.Session()

    def _apply(url, updates, adds, label):
        payload = {"f": "json", "token": token}
        if updates:
            payload["updates"] = json.dumps(updates)
        if adds:
            payload["adds"] = json.dumps(adds)
        if not updates and not adds:
            return
        r = session.post(url + "/applyEdits", data=payload, timeout=60)
        r.raise_for_status()
        js = r.json()
        log(f"{label} applyEdits: {js}")

    _apply(URL_PARENT, updates_parent, adds_parent, "PARENT")
    _apply(URL_CHILD,  updates_child,  adds_child,  "CHILD")

    # clear Dirty flags
    xl.ScreenUpdating = False
    for row_i, row_data in enumerate(data, start=2):
        row = list(row_data)
        if row[dirty_i]:
            sh.Cells(row_i, dirty_i + 1).Value = False
    wb.Save()
    xl.ScreenUpdating = True

    log("submit_registry complete")
    return 0


# ---------- MAIN ----------

def main():
    log("=== breedingSync START ===")
    if len(sys.argv) < 3:
        log("Usage: breedingSync.py <action> <workbook_path>")
        sys.exit(1)

    action  = sys.argv[1]
    wb_path = sys.argv[2]
    log(f"action='{action}'  workbook={wb_path}")
    log(f"python={sys.executable}  cwd={os.getcwd()}")

    try:
        if action == "import_registry":
            sys.exit(import_registry(wb_path))
        elif action == "submit_registry":
            sys.exit(submit_registry(wb_path))
        else:
            log(f"Unknown action: {action}")
            sys.exit(1)
    except Exception as e:
        import traceback
        log(f"FATAL: {e}")
        log(traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()
