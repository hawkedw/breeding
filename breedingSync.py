# breedingSync.py
# Импорт и выгрузка листа "Реестр опытов" ArcGIS -> Excel и обратно

import sys, os, json, datetime
import requests
import win32com.client as win32

# Форсируем UTF-8 для stdout/stderr (cp1251 не умеет ✓ ✗ и др.)
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

LOG_PATH = os.path.join(os.path.dirname(__file__), "breedingSync.log")


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

DIRTY_ALIAS    = "Dirty"
SHEET_REGISTRY = "Реестр опытов"

# Родительский слой
URL_PARENT = "https://maps.ekoniva-apk.org/arcgis/rest/services/breeding/breeding/FeatureServer/0"
# Дочерняя таблица (customer)
URL_CHILD  = "https://maps.ekoniva-apk.org/arcgis/rest/services/breeding/breeding/FeatureServer/1"

SORT_FIELD = "created_date"

# ---------- FIELD MAP (родительский слой -> столбцы) ----------
# col=None — поля без фиксированного столбца (дочерние / служебные)

FIELDS_PARENT = [
    {"n": "country",          "alias": "country",          "type": "TEXT",   "col": 1},   # A
    {"n": "region",           "alias": "region",           "type": "TEXT",   "col": 2},   # B
    {"n": "site",             "alias": "site",             "type": "TEXT",   "col": 3},   # C
    {"n": "devision",         "alias": "devision",         "type": "TEXT",   "col": 4},   # D
    {"n": "crop",             "alias": "crop",             "type": "TEXT",   "col": 5},   # E
    {"n": "farm",             "alias": "farm",             "type": "TEXT",   "col": 6},   # F
    {"n": "responsable",      "alias": "responsable",      "type": "TEXT",   "col": 7},   # G
    {"n": "fieldNumber",      "alias": "fieldNumber",      "type": "TEXT",   "col": 8},   # H
    {"n": "areaHa",           "alias": "areaHa",           "type": "NUMBER", "col": 9},   # I
    {"n": "scheme",           "alias": "scheme",           "type": "TEXT",   "col": 10},  # J
    {"n": "experimentName",   "alias": "experimentName",   "type": "TEXT",   "col": 11},  # K
    {"n": "type",             "alias": "type",             "type": "TEXT",   "col": 12},  # L
    {"n": "productPurpose",   "alias": "productPurpose",   "type": "TEXT",   "col": 13},  # M
    {"n": "trialPurpose",     "alias": "trialPurpose",     "type": "TEXT",   "col": 14},  # N
    {"n": "status",           "alias": "status",           "type": "TEXT",   "col": 15},  # O
    {"n": "plantingDate",     "alias": "plantingDate",     "type": "DATE",   "col": 16},  # P  -> dd.mm.yyyy
    {"n": "haverstDate",      "alias": "haverstDate",      "type": "DATE",   "col": 17},  # Q  -> dd.mm.yyyy
    {"n": "report",           "alias": "report",           "type": "TEXT",   "col": 18},  # R
    # S — пустая
    # T:W — customer из дочерней таблицы
    # X:Z — пустые
    {"n": "created_date",     "alias": "created_date",     "type": "DATE",   "col": 27},  # AA -> dd.mm.yyyy hh:mm
    {"n": "last_edited_date", "alias": "last_edited_date", "type": "DATE",   "col": 28},  # AB -> dd.mm.yyyy hh:mm
    # AC — Dirty
    # AD — GlobalID родителя
    # AE — GlobalID первой дочерней записи
]

# Столбцы customer (дочерняя таблица), транспонируем до 4 строк в T:W
CUSTOMER_COLS  = [20, 21, 22, 23]   # T=20, U=21, V=22, W=23
CUSTOMER_FIELD = "customer"
CUSTOMER_ALIAS = "customer"

DIRTY_COL      = 29   # AC
PARENT_GID_COL = 30   # AD
CHILD_GID_COL  = 31   # AE

TOTAL_COLS = CHILD_GID_COL   # 31

# Редактируемые столбцы C:R (site..report) — только их пишем обратно
EDITABLE_COLS = set(range(3, 19))

# Системные поля — никогда не пишем обратно
SYS_SKIP = {"created_user", "created_date", "last_edited_user", "last_edited_date"}

# Поля дат только без времени
DATE_ONLY_FIELDS = {"plantingDate", "haverstDate"}

# ---------- AUTH ----------

TOKEN_URL        = "https://maps.ekoniva-apk.org/portal/sharing/rest/generateToken"
ARC_USERNAME_ENV = "ARCGIS_BREEDING_USER"
ARC_PASSWORD_ENV = "ARCGIS_BREEDING_PASS"


def get_token() -> str:
    username = os.environ.get(ARC_USERNAME_ENV)
    password = os.environ.get(ARC_PASSWORD_ENV)
    if not username or not password:
        raise RuntimeError(
            f"Не заданы переменные окружения "
            f"{ARC_USERNAME_ENV}/{ARC_PASSWORD_ENV} с учётными данными."
        )
    payload = {
        "username": username,
        "password": password,
        "client": "referer",
        "referer": "https://maps.ekoniva-apk.org",
        "expiration": 60,
        "f": "json",
    }
    resp = requests.post(TOKEN_URL, data=payload, timeout=30)
    js = resp.json()
    tok = js.get("token")
    if not tok:
        log(f"Token error: {js}")
        raise RuntimeError(f"Token error: {js}")
    return tok


# ---------- DATE HELPERS ----------

EPOCH      = datetime.datetime(1970, 1, 1)
OFFSET     = datetime.timedelta(hours=3)   # MSK = UTC+3
EXCEL_EPOCH = datetime.datetime(1899, 12, 30)


def esri_ms_to_dt(ms: int) -> datetime.datetime:
    return EPOCH + datetime.timedelta(milliseconds=int(ms)) + OFFSET


def dt_to_esri(dt: datetime.datetime) -> int:
    if dt.tzinfo is not None:
        dt_utc = dt.astimezone(datetime.timezone.utc).replace(tzinfo=None)
    else:
        dt_utc = dt - OFFSET
    return int((dt_utc - EPOCH).total_seconds() * 1000)


def dt_to_excel_serial(dt: datetime.datetime) -> float:
    if dt.tzinfo is not None:
        dt = dt.astimezone(datetime.timezone.utc).replace(tzinfo=None)
    delta = dt - EXCEL_EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86400.0


def excel_serial_to_dt(x: float) -> datetime.datetime:
    return EXCEL_EPOCH + datetime.timedelta(days=float(x))


def arc_value_to_excel_serial(v) -> float:
    v = float(v)
    if v > 100_000_000_000:
        dt = esri_ms_to_dt(int(v))
        return dt_to_excel_serial(dt)
    return v


# ---------- QUERY LAYER (PAGED) ----------

def query_layer(url: str, where: str = "1=1", order_by: str = ""):
    token = get_token()
    session = requests.Session()
    feats = []
    offset = 0
    page_size = 2000

    while True:
        params = {
            "where": where,
            "outFields": "*",
            "f": "json",
            "token": token,
            "resultOffset": offset,
            "resultRecordCount": page_size,
        }
        if order_by:
            params["orderByFields"] = f"{order_by} DESC"

        r = session.get(url + "/query", params=params, timeout=60)
        js = r.json()
        if "error" in js:
            log(f"ArcGIS error: {js['error']}")
            raise RuntimeError(f"ArcGIS error: {js['error']}")

        page = js.get("features", [])
        feats.extend(page)
        log(
            f"query_layer offset={offset}: got {len(page)}, total {len(feats)}, "
            f"exceeded={js.get('exceededTransferLimit')}"
        )
        if not js.get("exceededTransferLimit") or not page:
            break
        offset += len(page)

    return feats


# ---------- EXCEL HELPERS ----------

def _to_2d(rows):
    return tuple(tuple(r) for r in rows)


def attach_workbook(path: str):
    xl = win32.Dispatch("Excel.Application")
    abs_path = os.path.abspath(path)
    for wb in xl.Workbooks:
        if os.path.abspath(wb.FullName) == abs_path:
            return wb, False, xl, xl.Workbooks.Count
    prev_count = xl.Workbooks.Count
    wb = xl.Workbooks.Open(abs_path)
    return wb, True, xl, prev_count


def _apply_freeze(excel, sh, rows, cols):
    sh.Parent.Activate()
    sh.Activate()
    win = excel.ActiveWindow
    try:
        win.FreezePanes = False
    except Exception:
        pass
    win.SplitRow    = rows
    win.SplitColumn = cols
    win.FreezePanes = True


_LAYER_INFO_CACHE = {}


def _get_layer_oid_field(layer_url: str, token: str):
    if layer_url in _LAYER_INFO_CACHE:
        return _LAYER_INFO_CACHE[layer_url]
    try:
        r  = requests.get(layer_url, params={"f": "json", "token": token}, timeout=60)
        js = r.json()
        oid = js.get("objectIdField") or js.get("objectIdFieldName")
    except Exception:
        oid = None
    _LAYER_INFO_CACHE[layer_url] = oid
    return oid


# ---------- IMPORT ----------

def import_registry(wb):
    log("=== import_registry START ===")

    log("Querying parent layer...")
    parent_feats = query_layer(URL_PARENT, "1=1", SORT_FIELD)
    log(f"Parent: {len(parent_feats)} features")

    parent_feats.sort(
        key=lambda f: f.get("attributes", {}).get(SORT_FIELD) or 0,
        reverse=True,
    )

    log("Querying child table...")
    child_feats = query_layer(URL_CHILD, "1=1", "")
    log(f"Child: {len(child_feats)} records")

    # parentglobalid -> [customer, ...]
    child_index: dict[str, list] = {}
    for cf in child_feats:
        attrs = cf.get("attributes", {})
        pgid  = attrs.get("parentglobalid")
        cval  = attrs.get(CUSTOMER_FIELD)
        if pgid:
            child_index.setdefault(pgid, []).append(cval)

    log(f"Child index: {len(child_index)} parent GlobalIDs")

    excel = wb.Application
    xlCalcAutomatic = -4105
    xlCalcManual    = -4135

    prev_screen = excel.ScreenUpdating
    prev_calc   = excel.Calculation
    prev_events = excel.EnableEvents
    excel.ScreenUpdating = False
    excel.Calculation    = xlCalcManual
    excel.EnableEvents   = False

    try:
        try:
            sh = wb.Worksheets(SHEET_REGISTRY)
            try:
                sh.AutoFilterMode = False
                if sh.FilterMode:
                    sh.ShowAllData()
                sh.Cells.EntireRow.Hidden    = False
                sh.Cells.EntireColumn.Hidden = False
            except Exception:
                pass
            sh.Cells.Clear()
        except Exception:
            sh = wb.Worksheets.Add()
            sh.Name = SHEET_REGISTRY

        # Заголовки
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

        sh.Range(sh.Cells(1, 1), sh.Cells(1, TOTAL_COLS)).Value = _to_2d([headers])

        # Строки данных
        data = []
        date_cols_log = {}

        for ft_idx, ft in enumerate(parent_feats):
            if ft_idx % 5000 == 0 and ft_idx > 0:
                log(f"  -> prepared {ft_idx}/{len(parent_feats)} rows")

            attrs = ft.get("attributes", {})
            row   = [""] * TOTAL_COLS

            for f in FIELDS_PARENT:
                col = f.get("col")
                if not col:
                    continue
                v = attrs.get(f["n"])
                if v is None:
                    continue
                if f["type"] == "DATE" and isinstance(v, (int, float)):
                    xl_val = arc_value_to_excel_serial(v)
                    row[col - 1] = xl_val
                    if ft_idx == 0:
                        date_cols_log[col] = (f["n"], v, xl_val)
                else:
                    row[col - 1] = v

            parent_gid = attrs.get("GlobalID", "")
            customers  = child_index.get(parent_gid, [])
            child_gid  = ""
            if customers:
                for cf in child_feats:
                    if cf.get("attributes", {}).get("parentglobalid") == parent_gid:
                        child_gid = cf.get("attributes", {}).get("GlobalID", "")
                        break

            for i, c in enumerate(CUSTOMER_COLS):
                if i < len(customers):
                    row[c - 1] = customers[i] if customers[i] is not None else ""

            row[DIRTY_COL - 1]      = False
            row[PARENT_GID_COL - 1] = parent_gid
            row[CHILD_GID_COL - 1]  = child_gid
            data.append(row)

        for col, (name, arc_val, xl_val) in date_cols_log.items():
            log(f"DATE col {col} '{name}': ArcGIS={arc_val}, Excel={xl_val}")

        data_rows = len(data)
        log(f"Writing {data_rows} rows...")

        if data_rows:
            sh.Range(
                sh.Cells(2, 1),
                sh.Cells(1 + data_rows, TOTAL_COLS),
            ).Value = _to_2d(data)

        # Формат дат
        if data_rows:
            for f in FIELDS_PARENT:
                col = f.get("col")
                if not col or f["type"] != "DATE":
                    continue
                date_only = f["n"] in DATE_ONLY_FIELDS
                fmt = "dd.mm.yyyy" if date_only else "dd.mm.yyyy hh:mm"
                try:
                    rng = sh.Range(sh.Cells(2, col), sh.Cells(1 + data_rows, col))
                    rng.NumberFormat = fmt
                    try:
                        rng.NumberFormatLocal = "ДД.ММ.ГГГГ" if date_only else "ДД.ММ.ГГГГ чч:мм"
                    except Exception:
                        pass
                    log(f"[OK] col {col} '{f['n']}' fmt='{fmt}'")
                except Exception as e:
                    log(f"[ERR] col {col} '{f['n']}' fmt crash: {e}")

        # AutoFilter
        last_row = max(2, sh.Cells(sh.Rows.Count, 1).End(-4162).Row)
        sh.Range(sh.Cells(1, 1), sh.Cells(last_row, TOTAL_COLS)).AutoFilter()

        # Freeze строка 1, столбец 3 (C)
        _apply_freeze(excel, sh, 1, 3)

        log(f"import_registry complete: {data_rows} rows")

    finally:
        excel.Calculation    = prev_calc if prev_calc in (xlCalcAutomatic, xlCalcManual) else xlCalcAutomatic
        excel.ScreenUpdating = prev_screen
        excel.EnableEvents   = prev_events


# ---------- SUBMIT ----------

def submit_registry(wb):
    log("=== submit_registry START ===")

    excel = wb.Application
    xlCalcAutomatic = -4105
    xlCalcManual    = -4135

    prev_screen = excel.ScreenUpdating
    prev_calc   = excel.Calculation
    prev_events = excel.EnableEvents
    excel.ScreenUpdating = False
    excel.Calculation    = xlCalcManual
    excel.EnableEvents   = False

    try:
        try:
            sh = wb.Worksheets(SHEET_REGISTRY)
        except Exception:
            log(f"Sheet '{SHEET_REGISTRY}' not found")
            return

        last_col = sh.Cells(1, sh.Columns.Count).End(-4159).Column
        last_row = max(
            sh.Cells(sh.Rows.Count, DIRTY_COL).End(-4162).Row,
            sh.Cells(sh.Rows.Count, PARENT_GID_COL).End(-4162).Row,
            sh.Cells(sh.Rows.Count, 1).End(-4162).Row,
        )

        if last_row <= 1:
            log("No data to submit")
            return

        hdr_range = sh.Range(sh.Cells(1, 1), sh.Cells(1, last_col)).Value
        headers   = list(hdr_range[0])
        log(f"last_row={last_row} last_col={last_col}")

        alias_to_name = {}
        name_to_type  = {}
        for f in FIELDS_PARENT:
            n, al = f.get("n"), f.get("alias")
            if n and al:
                alias_to_name[al] = n
                name_to_type[n]   = f.get("type")

        data_range = sh.Range(
            sh.Cells(2, 1),
            sh.Cells(last_row, last_col),
        ).Value

        token     = get_token()
        oid_field = _get_layer_oid_field(URL_PARENT, token)
        if not oid_field:
            log(f"Cannot read objectIdField from {URL_PARENT}")
            return

        edits = []
        for r_idx, row in enumerate(data_range, start=2):
            row = list(row)
            if not row[DIRTY_COL - 1]:
                continue
            parent_gid = row[PARENT_GID_COL - 1]
            if parent_gid in (None, ""):
                log(f"Row {r_idx}: no GlobalID, skip")
                continue

            attrs = {"GlobalID": str(parent_gid).strip()}

            for col_idx, alias in enumerate(headers, start=1):
                if col_idx not in EDITABLE_COLS:
                    continue
                if not alias or alias == DIRTY_ALIAS:
                    continue
                name = alias_to_name.get(alias)
                if not name or name.lower() in SYS_SKIP:
                    continue

                v      = row[col_idx - 1]
                f_type = name_to_type.get(name)

                if v in ("", None):
                    attrs[name] = None
                    continue

                if f_type == "DATE":
                    if isinstance(v, datetime.datetime):
                        attrs[name] = dt_to_esri(v)
                    elif isinstance(v, (int, float)):
                        attrs[name] = dt_to_esri(excel_serial_to_dt(float(v)))
                    elif isinstance(v, str):
                        dt = None
                        for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y"):
                            try:
                                dt = datetime.datetime.strptime(v.strip(), fmt)
                                break
                            except ValueError:
                                pass
                        attrs[name] = dt_to_esri(dt) if dt else None
                    else:
                        attrs[name] = None
                elif isinstance(v, datetime.datetime):
                    attrs[name] = dt_to_esri(v)
                else:
                    attrs[name] = v

            edits.append({"attributes": attrs, "row": r_idx})

        if not edits:
            log("No dirty rows")
            return

        log(f"Sending {len(edits)} updates (useGlobalIds=True)...")

        feats_json = json.dumps([{"attributes": e["attributes"]} for e in edits])
        payload = {
            "f": "json",
            "token": token,
            "rollbackOnFailure": "True",
            "useGlobalIds": "True",
            "updates": feats_json,
        }
        res = requests.post(URL_PARENT + "/applyEdits", data=payload, timeout=60)
        js  = res.json()

        if "error" in js:
            log(f"applyEdits error: {js['error']}")
            return

        results = js.get("updateResults") or []
        for e, r in zip(edits, results):
            row_idx = e["row"]
            if r.get("success"):
                sh.Cells(row_idx, DIRTY_COL).Value = False
                log(f"Row {row_idx}: OK")
            else:
                err = r.get("error", {}).get("description", "?")
                log(f"Row {row_idx}: FAILED - {err}")
                try:
                    sh.Cells(row_idx, DIRTY_COL).AddComment(err)
                except Exception:
                    pass

        log("submit_registry complete")

    finally:
        excel.Calculation    = prev_calc if prev_calc in (xlCalcAutomatic, xlCalcManual) else xlCalcAutomatic
        excel.ScreenUpdating = prev_screen
        excel.EnableEvents   = prev_events


# ---------- MAIN ----------

def normalize_action(a: str) -> str:
    a = (a or "").strip()
    if a.lower().startswith("action="):
        a = a.split("=", 1)[1].strip()
    a = a.lower()
    a = a.replace("sumbit", "submit").replace("registy", "registry").replace("registery", "registry")
    return a


def main(argv=None) -> int:
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

    wb, opened_here, excel, prev_count = attach_workbook(wb_path)

    try:
        action_map = {
            "import_registry": import_registry,
            "submit_registry": submit_registry,
        }
        fn = action_map.get(action)
        if fn is None:
            log(f"Unknown action: {action!r}. Available: {list(action_map)}")
            return 1
        fn(wb)
        wb.Save()
        return 0
    finally:
        if opened_here:
            try:
                wb.Close(SaveChanges=True)
            except Exception:
                pass
            try:
                if prev_count == 0 and excel.Workbooks.Count == 0:
                    excel.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
