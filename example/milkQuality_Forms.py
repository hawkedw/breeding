# milkQuality_Forms.py
# Импорт и выгрузка форм 1 / 2 / 5 ArcGIS в Excel

import sys, os, json, datetime
import requests
import win32com.client as win32

LOG_PATH = os.path.join(os.path.dirname(__file__), "milkQuality_Forms.log")

def log(msg: str):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


DIRTY_ALIAS = "Dirty"
SHEET_F1 = "Форма 1"
SHEET_F2 = "Форма 2"
SHEET_F5 = "Форма 5"

URL_F1 = "https://maps.ekoniva-apk.org/arcgis/rest/services/quality_surveys/milkQuality/FeatureServer/2"
URL_F2 = "https://maps.ekoniva-apk.org/arcgis/rest/services/quality_surveys/milkQuality/FeatureServer/3"
URL_F5 = "https://maps.ekoniva-apk.org/arcgis/rest/services/quality_surveys/milkQuality/FeatureServer/9"

SORT_F1_FIELD = "timeStart"
SORT_F2_FIELD = "milktruk_departure_time"

# ---------- FIELD MAPS ----------
FIELDS_F1 = [
***
]

FIELDS_F2 = [
***
]

# --- Форма 5 (колонки) ---
# ВАЖНО: для виртуальных колонок итогов добавляем id, чтобы строить шапку по ключам.

FIELDS_F5 = [
***
]

# --- Форма 5: схема групп/подгрупп для строк 1-2 (ключи = f["n"]) ---

FORM5_GROUPS = [
***
]

FORM5_SUBGROUPS = [
    # block_1
***

# ---------- ARC / AUTH ----------
TOKEN_URL = "https://maps.ekoniva-apk.org/portal/sharing/rest/generateToken"


ARC_USERNAME_ENV = "ARCGIS_QUALITY_USER"
ARC_PASSWORD_ENV = "ARCGIS_QUALITY_PASS"


def get_token() -> str:
    """Получить токен ArcGIS Portal.

    Логика:
    - имя пользователя и пароль берутся из переменных окружения
      ARCGIS_QUALITY_USER и ARCGIS_QUALITY_PASS;
    - никакие учётные данные не хардкодятся в репозитории/файлах;
    - client=referer, referer указываем на портал maps.ekoniva-apk.org.
    """
    username = os.environ.get(ARC_USERNAME_ENV)
    password = os.environ.get(ARC_PASSWORD_ENV)
    if not username or not password:
        raise RuntimeError(
            "Не заданы переменные окружения "
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
        # Логируем ответ для отладки, если что-то пойдет не так
        log(f"Token error: {js}")
        raise RuntimeError(f"Token error: {js}")
    return tok


# ---------- HELPERS ----------

EPOCH = datetime.datetime(1970, 1, 1)

# Локальный оффсет (MSK). Используется ТОЛЬКО для отображения/ввода локального времени.
OFFSET = datetime.timedelta(hours=3)

# Excel / OLE Automation epoch (важно: именно 1899-12-30)
EXCEL_EPOCH = datetime.datetime(1899, 12, 30)

# ----------------- DATE FIXES (F1/F2) -----------------

# Проблемный алиас в Ф1 может отличаться пробелом перед скобкой, поэтому сравниваем "нормализовано".
_FORCE_DATE_F1_ALIASES_NORM = {
    "времяокончаниядоения(остаток)",
}

# Системные поля Esri (часто не приходят как type=DATE в fields, но по смыслу это дата)
_FORCE_DATE_NAMES = {
    "created_date",
    "last_edited_date",
}

def _norm_alias(s: str) -> str:
    if not s:
        return ""
    return "".join(str(s).split()).lower()  # убрать все пробелы


def is_date_col(sheet_name: str, f: dict) -> bool:
    """True если колонку надо трактовать как DATE/DATETIME (даже если type не DATE)."""
    if f.get("type") == "DATE":
        return True

    n = f.get("n")
    if n in _FORCE_DATE_NAMES:
        return True

    if sheet_name == SHEET_F1:
        a = f.get("alias")
        if _norm_alias(a) in _FORCE_DATE_F1_ALIASES_NORM:
            return True

    return False


def esri_ms_to_dt(ms: int) -> datetime.datetime:
    """ArcGIS: ms с 1970-01-01 UTC -> локальное время (UTC+3) для Excel."""
    return EPOCH + datetime.timedelta(milliseconds=int(ms)) + OFFSET


def dt_to_esri(dt: datetime.datetime) -> int:
    """Excel/пользовательское локальное время (UTC+3) -> ms UTC для ArcGIS."""
    if dt.tzinfo is not None:
        dt_utc = dt.astimezone(datetime.timezone.utc).replace(tzinfo=None)
    else:
        dt_utc = dt - OFFSET
    return int((dt_utc - EPOCH).total_seconds() * 1000)


def dt_to_excel_serial(dt: datetime.datetime) -> float:
    """datetime -> Excel serial (float), чтобы COM не делал tz-конвертацию."""
    if dt.tzinfo is not None:
        dt = dt.astimezone(datetime.timezone.utc).replace(tzinfo=None)
    delta = dt - EXCEL_EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86400.0


def excel_serial_to_dt(x: float) -> datetime.datetime:
    """Excel serial -> naive datetime."""
    return EXCEL_EPOCH + datetime.timedelta(days=float(x))


def _arc_value_to_excel_serial(v: float) -> float:
    """
    Если вдруг прилетело ms epoch (~1.7e12) -> конвертируем в Excel serial.
    Если прилетело уже Excel serial (~46000) -> оставляем как есть.
    """
    # ArcGIS epoch ms сейчас ~1.6e12..1.9e12, Excel serial ~40k..60k
    if v > 100_000_000_000:  # 1e11 — надежная граница
        dt = esri_ms_to_dt(int(v))
        return dt_to_excel_serial(dt)
    else:
        return float(v)


def attach_workbook(path: str):
    xl = win32.Dispatch("Excel.Application")
    abs_path = os.path.abspath(path)
    for wb in xl.Workbooks:
        if os.path.abspath(wb.FullName) == abs_path:
            return wb, False, xl, xl.Workbooks.Count
    prev_count = xl.Workbooks.Count
    wb = xl.Workbooks.Open(abs_path)
    return wb, True, xl, prev_count


# ---------- QUERY LAYER (PAGED) ----------

def query_layer(url: str, where: str, order_by: str):
    """ArcGIS запрос с пагинацией. Сортирует по order_by DESC (новые записи первыми)."""
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
            params["orderByFields"] = f"{order_by} DESC"  # ← ВСЕГДА DESC

        r = session.get(url + "/query", params=params, timeout=60)
        js = r.json()
        if "error" in js:
            log(f"ArcGIS error in query_layer: {js['error']}")
            raise RuntimeError(f"ArcGIS error: {js['error']}")

        page = js.get("features", [])
        feats.extend(page)

        log(
            f"query_layer page: got {len(page)} features, total {len(feats)}, "
            f"exceeded={js.get('exceededTransferLimit')}"
        )

        if not js.get("exceededTransferLimit") or not page:
            break

        offset += len(page)

    return feats

# ---------- WRITE TO SHEET (FAST) ----------

def _to_2d(rows):
    return tuple(tuple(r) for r in rows)


def _rgb(r, g, b):
    return r + g * 256 + b * 65536


def _apply_freeze(excel, sh, rows, cols):
    wb = sh.Parent
    wb.Activate()
    sh.Activate()
    win = excel.ActiveWindow
    try:
        win.FreezePanes = False
    except Exception:
        pass
    win.SplitRow = rows
    win.SplitColumn = cols
    win.FreezePanes = True


def build_form5_group_rows(sh, fields_f5, header_row=3):
    # +1 за Dirty
    cols = len(fields_f5) + 1

    sh.Cells.UnMerge()
    sh.Range(sh.Cells(1, 1), sh.Cells(2, cols)).ClearContents()

    # карта "ключ -> номер колонки" (ключ = f["n"] или f["id"])
    col_by_key = {}
    for i, f in enumerate(fields_f5, start=1):
        key = f.get("n") or f.get("id")
        if key:
            col_by_key[key] = i

    xlCenter = -4108

    def merge_write(row, c1, c2, text):
        if not c1 or not c2 or c2 < c1:
            return
        rg = sh.Range(sh.Cells(row, c1), sh.Cells(row, c2))
        rg.Merge()
        rg.Value = text
        rg.Font.Bold = True
        rg.HorizontalAlignment = xlCenter
        rg.WrapText = True

    # row 1 (groups)
    # row 1 (groups)
    for g in FORM5_GROUPS:
        c1 = col_by_key.get(g["start"])
        c2 = g.get("end_col") or col_by_key.get(g["end"])
        merge_write(1, c1, c2, g["text"])

    # row 2 (subgroups)
    for sg in FORM5_SUBGROUPS:
        merge_write(
            2,
            col_by_key.get(sg["start"]),
            col_by_key.get(sg["end"]),
            sg["text"]
        )

        merge_write(2, col_by_key.get(sg["start"]), col_by_key.get(sg["end"]), sg["text"])

    sh.Parent.Application.Calculate()


def fill_form5_totals(sh, header_row, data_rows):
    if data_rows == 0:
        return

    first = header_row + 1
    last = header_row + data_rows

    for r in range(first, last + 1):
        sh.Cells(r, 31).Formula = f"=SUM(F{r},H{r},J{r},L{r},N{r},P{r},R{r},T{r},V{r},X{r},Z{r},AB{r},AD{r})"
        sh.Cells(r, 56).Formula = f"=SUM(AG{r},AI{r},AK{r},AM{r},AO{r},AQ{r},AS{r},AU{r},AW{r},AY{r},BA{r},BC{r})"
        sh.Cells(r, 71).Formula = f"=SUM(BF{r},BH{r},BJ{r},BL{r},BN{r},BP{r},BR{r})"
        sh.Cells(r, 102).Formula = f"=SUM(BU{r},BW{r},BY{r},CA{r},CC{r},CE{r},CG{r},CI{r},CK{r},CM{r},CO{r},CQ{r},CS{r},CU{r},CW{r})"
        sh.Cells(r, 111).Formula = f"=SUM(DD{r},DF{r})"
        sh.Cells(r, 112).Formula = f"=SUM(AE{r},BD{r},BS{r},CX{r},DG{r})"


def color_form5_columns(sh, cols):
    last_row = sh.Rows.Count
    last_col = cols

    def set_block(c1, c2, color_hex):
        if c1 > last_col:
            return
        c2 = min(c2, last_col)

        rng = sh.Range(sh.Cells(1, c1), sh.Cells(last_row, c2))
        color_hex = color_hex.lstrip("#")
        r_int = int(color_hex[0:2], 16)
        g_int = int(color_hex[2:4], 16)
        b_int = int(color_hex[4:6], 16)
        rng.Interior.Color = _rgb(r_int, g_int, b_int)

    set_block(1, 4, "A9D08E")
    set_block(5, 31, "E2EFDA")
    set_block(32, 56, "FFF2CC")
    set_block(57, 71, "FCE4D6")
    set_block(72, 102, "D9E1F2")   # BT..CX
    set_block(103, 111, "D0CECE")  # CY..DG
    set_block(112, 112, "EC7524")  # DH


    used = sh.Range(sh.Cells(1, 1), sh.Cells(last_row, last_col))
    used.Borders.LineStyle = 1


def write_sheet(wb, sheet_name: str, fields, features, sort_field: str):
    excel = wb.Application
    xlCalcAutomatic = -4105
    xlCalcManual = -4135
    xlLeft = -4131
    xlCenter = -4108

    prev_screen = excel.ScreenUpdating
    prev_calc = excel.Calculation
    prev_events = excel.EnableEvents

    excel.ScreenUpdating = False
    excel.Calculation = xlCalcManual
    excel.EnableEvents = False

    try:
        log(f"write_sheet START: sheet={sheet_name} features={len(features)} fields={len(fields)}")

        try:
            sh = wb.Worksheets(sheet_name)

            # сброс фильтра/скрытия перед очисткой
            try:
                sh.AutoFilterMode = False
                if sh.FilterMode:
                    sh.ShowAllData()
                sh.Cells.EntireRow.Hidden = False
                sh.Cells.EntireColumn.Hidden = False
            except Exception:
                pass

            sh.Cells.Clear()

        except Exception:
            sh = wb.Worksheets.Add()
            sh.Name = sheet_name

        header_row = 3 if sheet_name == SHEET_F5 else 1
        log(f"{sheet_name}: ProtectContents={sh.ProtectContents}, ProtectDrawingObjects={sh.ProtectDrawingObjects}")

        headers = [f["alias"] for f in fields]
        headers.append(DIRTY_ALIAS)
        cols = len(headers)

        # строка 3 (основные заголовки)
        hdr_rng = sh.Range(sh.Cells(header_row, 1), sh.Cells(header_row, cols))
        hdr_rng.Value = _to_2d([headers])

        # строки 1-2 (группы/блоки) только для Формы 5
        if sheet_name == SHEET_F5:
             build_form5_group_rows(sh, FIELDS_F5, header_row=header_row)


        # Подготовка данных
        data = []
        date_cols_sample = {}

        log(f"Preparing data: {len(features)} rows x {len(fields)} cols...")
        for ft_idx, ft in enumerate(features):
            if ft_idx > 0 and ft_idx % 5000 == 0:
                log(f"  -> processed {ft_idx}/{len(features)} rows...")

            attrs = ft.get("attributes", {})
            row_vals = []

            for col_idx, f in enumerate(fields, start=1):
                name = f.get("n")
                if not name:
                    row_vals.append("")
                    continue

                v = attrs.get(name)
                if v is None:
                    row_vals.append("")
                    continue

                # DATE: ArcGIS ms -> Excel serial
                if is_date_col(sheet_name, f) and isinstance(v, (int, float)):
                    xl_val = _arc_value_to_excel_serial(float(v))
                    row_vals.append(xl_val)
                    if ft_idx == 0 and col_idx not in date_cols_sample:
                        date_cols_sample[col_idx] = (name, v, xl_val)
                else:
                    row_vals.append(v)

            row_vals.append(False)  # Dirty
            data.append(row_vals)

        data_rows = len(data)

        # Лог: что записалось в DATE-колонки
        for col_idx, (name, arc_val, xl_val) in date_cols_sample.items():
            log(
                f"DATE col {col_idx} '{name}': ArcGIS={arc_val} ({type(arc_val).__name__}), "
                f"Excel={xl_val} ({type(xl_val).__name__})"
            )

        # Запись данных
        if data_rows:
            data_rng = sh.Range(
                sh.Cells(header_row + 1, 1),
                sh.Cells(header_row + data_rows, cols),
            )
            data_rng.Value = _to_2d(data)

        # форматы (только если есть данные)
        if data_rows:
            for col_idx, f in enumerate(fields, start=1):
                fname = f.get("n", "?")
                t = f["type"]
                log(f"⚙ formatting col {col_idx} '{fname}' type={t}")

                try:
                    rng = sh.Range(
                        sh.Cells(header_row + 1, col_idx),
                        sh.Cells(header_row + data_rows, col_idx),
                    )

                    if is_date_col(sheet_name, f):
                        date_only = (sheet_name == SHEET_F5 and f.get("n") == "dateForm5")

                        rng.NumberFormat = "dd.mm.yyyy" if date_only else "dd.mm.yyyy hh:mm"
                        try:
                            rng.NumberFormatLocal = "ДД.ММ.ГГГГ" if date_only else "ДД.ММ.ГГГГ чч:мм"
                        except Exception:
                            pass

                        applied_fmt = rng.NumberFormat
                        log(f"✓ col {col_idx} '{fname}' fmt='{applied_fmt}'")

                        first_val = sh.Cells(header_row + 1, col_idx).Text
                        log(f"  col {col_idx} row {header_row+1} Text='{first_val}'")

                    elif t in ("INT", "OID"):
                        rng.NumberFormat = "0"
                    elif t == "NUMBER":
                        rng.NumberFormat = "0.00"

                except Exception as e:
                    log(f"✗✗✗ col {col_idx} '{fname}' CRASH: {e}")

        # AutoFilter
        last_row = max(header_row + 1, sh.Cells(sh.Rows.Count, 1).End(-4162).Row)
        sh.Range(sh.Cells(header_row, 1), sh.Cells(last_row, cols)).AutoFilter()

        # ColumnWidth / выравнивание для Формы 2
        if sheet_name == SHEET_F2:
            w_date, w_text, w_num = 18, 20, 12
            for i, f in enumerate(fields, start=1):
                t = f.get("type")
                if t == "DATE" or is_date_col(sheet_name, f):
                    w = w_date
                elif t in ("NUMBER", "INT"):
                    w = w_num
                else:
                    w = w_text
                sh.Columns(i).ColumnWidth = w
            sh.Columns(cols).ColumnWidth = 8

            if data_rows:
                rng = sh.Range(
                    sh.Cells(header_row + 1, 4),
                    sh.Cells(header_row + data_rows, 4),
                )
                rng.HorizontalAlignment = xlCenter

        # Итоги/окраска/Freeze
        if sheet_name == SHEET_F1:
            _apply_freeze(excel, sh, header_row, 4)
        elif sheet_name == SHEET_F2:
            _apply_freeze(excel, sh, header_row, 5)
        elif sheet_name == SHEET_F5:
            if data_rows:
                fill_form5_totals(sh, header_row, data_rows)
                color_form5_columns(sh, cols)
            _apply_freeze(excel, sh, 3, 3)

    finally:
        excel.Calculation = prev_calc if prev_calc in (xlCalcAutomatic, xlCalcManual) else xlCalcAutomatic
        excel.ScreenUpdating = prev_screen
        excel.EnableEvents = prev_events


# ---------- PUBLIC ACTIONS ----------


def import_f1(wb):
    log(f"Starting query for Форма 1...")
    feats = query_layer(URL_F1, "1=1", SORT_F1_FIELD)
    log(f"Query complete: {len(feats)} features loaded")
    
    # Ограничение ДО сортировки (если вдруг будет > 15k)
    if len(feats) > 15000:
        log(f"!!! Pre-limiting to last 15000 records (was {len(feats)})")
        feats = feats[-15000:]
    
    log(f"Sorting {len(feats)} features by {SORT_F1_FIELD} DESC...")
    feats.sort(key=lambda f: f.get("attributes", {}).get(SORT_F1_FIELD) or 0, reverse=True)
    
    if feats:
        first_date = feats[0].get("attributes", {}).get(SORT_F1_FIELD, "???")
        last_date = feats[-1].get("attributes", {}).get(SORT_F1_FIELD, "???")
        log(f"Date range: NEWEST={first_date}, OLDEST={last_date}")
    
    log(f"Starting write_sheet for Форма 1...")
    write_sheet(wb, SHEET_F1, FIELDS_F1, feats, SORT_F1_FIELD)
    log(f"Форма 1 import complete")


def import_f2(wb):
    log("Starting query for Форма 2...")
    feats = query_layer(URL_F2, "1=1", SORT_F2_FIELD)
    log(f"Query complete: {len(feats)} features loaded")
    log(f"Sorting {len(feats)} features by {SORT_F2_FIELD} DESC...")
    feats.sort(key=lambda f: f.get("attributes", {}).get(SORT_F2_FIELD) or 0, reverse=True)

    if feats:
        first_date = feats[0].get("attributes", {}).get(SORT_F2_FIELD, "???")
        last_date  = feats[-1].get("attributes", {}).get(SORT_F2_FIELD, "???")
        log(f"Date range: NEWEST={first_date}, OLDEST={last_date}")

    log("Starting write_sheet for Форма 2...")
    write_sheet(wb, SHEET_F2, FIELDS_F2, feats, SORT_F2_FIELD)
    log("Форма 2 import complete")

def import_f5(wb):
    log(f"Starting query for Форма 5...")
    feats = query_layer(URL_F5, "1=1", "dateForm5")
    log(f"Query complete: {len(feats)} features loaded")
    
    # СОРТИРОВКА В PYTHON: DESC (новые сверху)
    log(f"Sorting by dateForm5 DESC...")
    feats.sort(key=lambda f: f.get("attributes", {}).get("dateForm5", 0), reverse=True)
    
    # Логируем диапазон дат
    if feats:
        first_date = feats[0].get("attributes", {}).get("dateForm5", "???")
        last_date = feats[-1].get("attributes", {}).get("dateForm5", "???")
        log(f"Date range: NEWEST={first_date}, OLDEST={last_date}")
    
    log(f"Starting write_sheet for Форма 5...")
    write_sheet(wb, SHEET_F5, FIELDS_F5, feats, "dateForm5")
    log(f"Форма 5 import complete")


# ---------- PUSH (updateFeatures) ----------

SYS_SKIP = {"created_user", "created_date", "last_edited_user", "last_edited_date"}

_LAYER_INFO_CACHE = {}


def _get_layer_oid_field(layer_url: str, token: str) -> str | None:
    """Возвращает имя OID поля слоя/таблицы (objectIdField)."""
    if layer_url in _LAYER_INFO_CACHE:
        return _LAYER_INFO_CACHE[layer_url]

    try:
        r = requests.get(layer_url, params={"f": "json", "token": token}, timeout=60)
        js = r.json()
        oid = js.get("objectIdField") or js.get("objectIdFieldName")
    except Exception:
        oid = None

    _LAYER_INFO_CACHE[layer_url] = oid
    return oid


def _to_int_oid(v):
    if v in (None, ""):
        return None
    try:
        # Excel часто отдаёт числа как float
        if isinstance(v, float):
            return int(v)
        if isinstance(v, (int,)):
            return v
        s = str(v).strip()
        if s == "":
            return None
        return int(float(s))
    except Exception:
        return None


def push_sheet(wb, sheet_name: str, fields, url: str):
    excel = wb.Application
    xlCalcAutomatic = -4105
    xlCalcManual = -4135

    prev_screen = excel.ScreenUpdating
    prev_calc = excel.Calculation
    prev_events = excel.EnableEvents

    excel.ScreenUpdating = False
    excel.Calculation = xlCalcManual
    excel.EnableEvents = False

    try:
        try:
            sh = wb.Worksheets(sheet_name)
        except Exception:
            log(f"Sheet '{sheet_name}' not found")
            return

        header_row = 3 if sheet_name == SHEET_F5 else 1

        last_col = sh.Cells(header_row, sh.Columns.Count).End(-4159).Column
        last_row = sh.Cells(sh.Rows.Count, 1).End(-4162).Row

        if last_row <= header_row:
            log("No data to push")
            return

        hdr_range = sh.Range(sh.Cells(header_row, 1), sh.Cells(header_row, last_col)).Value
        headers = list(hdr_range[0])

        def idx(name: str) -> int:
            # Берём ПОСЛЕДНЕЕ вхождение заголовка (часто есть дубликаты)
            for i in range(len(headers) - 1, -1, -1):
                if headers[i] == name:
                    return i + 1
            return 0

        dirty_col = idx(DIRTY_ALIAS)
        oid_col = idx("OBJECTID")
        gid_col = idx("GlobalID")

        if not dirty_col:
            log("Dirty column not found")
            return

        if not oid_col and not gid_col:
            log("Neither OBJECTID nor GlobalID column found")
            return
        # пересчитываем last_row по "надежным" колонкам (A часто пустая и режет диапазон)
        last_row_dirty = sh.Cells(sh.Rows.Count, dirty_col).End(-4162).Row
        last_row_oid = sh.Cells(sh.Rows.Count, oid_col).End(-4162).Row if oid_col else 0
        last_row_gid = sh.Cells(sh.Rows.Count, gid_col).End(-4162).Row if gid_col else 0
        last_row = max(last_row, last_row_dirty, last_row_oid, last_row_gid)

        if last_row <= header_row:
            log("No data to push")
            return

        data_range = sh.Range(
            sh.Cells(header_row + 1, 1),
            sh.Cells(last_row, last_col),
        ).Value

        # Диагностика: что именно считаем колонкой Dirty
        log(f"header_row={header_row} last_row={last_row} last_col={last_col} dirty_col={dirty_col}")
        try:
            log(f"dirty_header_addr={sh.Cells(header_row, dirty_col).Address} dirty_header_val={sh.Cells(header_row, dirty_col).Value!r}")
            # 13 поменяй на номер строки, где у тебя точно Dirty=TRUE
            log(f"dirty_r13_addr={sh.Cells(13, dirty_col).Address} dirty_r13_val={sh.Cells(13, dirty_col).Value!r} type={type(sh.Cells(13, dirty_col).Value)}")
        except Exception as ex:
            log(f"dirty debug read failed: {ex}")

        data_range = sh.Range(
            sh.Cells(header_row + 1, 1),
            sh.Cells(last_row, last_col),
        ).Value


        data_range = sh.Range(
            sh.Cells(header_row + 1, 1),
            sh.Cells(last_row, last_col),
        ).Value

        alias_to_name = {}
        name_to_type = {}
        for f in fields:
            n = f.get("n")
            al = f.get("alias")
            if n:
                name_to_type[n] = f.get("type")
            if n and al:
                alias_to_name[al] = n

        token = get_token()
        oid_field = _get_layer_oid_field(url, token) if oid_col else None

        if oid_col and not oid_field:
            log(f"Can't read objectIdField from layer: {url}")
            return

        edits = []
        for r_idx, row in enumerate(data_range, start=header_row + 1):
            row = list(row)
            dirty_val = row[dirty_col - 1]
            if not dirty_val:
                continue

            oid_val_raw = row[oid_col - 1] if oid_col else None
            gid_val = row[gid_col - 1] if gid_col else None

            oid_val = _to_int_oid(oid_val_raw)

            # Для F2 обновляем по GlobalID (useGlobalIds=True)
            if sheet_name == SHEET_F2 and gid_val in (None, ""):
                continue

            # Для остальных нужен OID (обновление)
            if sheet_name != SHEET_F2 and oid_val in (None, ""):
                continue

            attrs = {}

            if sheet_name == SHEET_F2:
                if gid_val not in (None, ""):
                    attrs["GlobalID"] = str(gid_val).strip()
            else:
                # ВАЖНО: кладём OID в реальное имя OID-поля (может быть FID, а не OBJECTID)
                attrs[oid_field] = oid_val

            for c, alias in enumerate(headers, start=1):
                if c == dirty_col:
                    continue
                if alias == DIRTY_ALIAS or not alias:
                    continue

                # Не трогаем системные идентификаторы из Excel
                if alias in ("OBJECTID", "GlobalID") or alias == oid_field:
                    continue

                name = alias_to_name.get(alias)
                if not name or name.lower() in SYS_SKIP:
                    continue

                v = row[c - 1]
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
                        s = v.strip()
                        dt = None
                        for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y"):
                            try:
                                dt = datetime.datetime.strptime(s, fmt)
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
        # ... после цикла, где ты наполняешь edits (после for r_idx, row in enumerate(data_range...):)

        cnt_dirty = 0
        cnt_dirty_with_oid = 0
        for r_idx, row in enumerate(data_range, start=header_row + 1):
            row = list(row)
            if row[dirty_col - 1]:
                cnt_dirty += 1
                if oid_col and _to_int_oid(row[oid_col - 1]) not in (None, ""):
                    cnt_dirty_with_oid += 1
        log(f"dirty_total={cnt_dirty} dirty_with_oid={cnt_dirty_with_oid}")

        if not edits:
            log("No dirty rows")
            return


        if not edits:
            log("No dirty rows")
            return

        feats_json = json.dumps([{"attributes": e["attributes"]} for e in edits])

        if sheet_name == SHEET_F2:
            payload = {
                "f": "json",
                "token": token,
                "rollbackOnFailure": "True",
                "useGlobalIds": "True",
                "updates": feats_json,
            }
            res = requests.post(url + "/applyEdits", data=payload, timeout=60)
            op_name = "applyEdits"
        else:
            payload = {
                "f": "json",
                "token": token,
                "rollbackOnFailure": "True",
                "features": feats_json,
            }
            res = requests.post(url + "/updateFeatures", data=payload, timeout=60)
            op_name = "updateFeatures"

        js = res.json()
        if "error" in js:
            log(f"{op_name} error for {sheet_name}: {js['error']}")
            return

        results = js.get("updateResults") or []
        if not results:
            log(f"No updateResults for {sheet_name}: {js}")
            return

        for e, r in zip(edits, results):
            row_idx = e["row"]
            if r.get("success"):
                sh.Cells(row_idx, dirty_col).Value = False
            else:
                log(f"Row {row_idx} update failed in {sheet_name}: {r}")
                err = r.get("error", {}).get("description", "?")
                sh.Cells(row_idx, dirty_col).AddComment(err)

    finally:
        excel.Calculation = prev_calc if prev_calc in (xlCalcAutomatic, xlCalcManual) else xlCalcAutomatic
        excel.ScreenUpdating = prev_screen
        excel.EnableEvents = prev_events

# ---------- MAIN ----------

def normalize_action(a: str) -> str:
    import re

    a = (a or "").strip()

    # action=submit_f5
    if a.lower().startswith("action="):
        a = a.split("=", 1)[1].strip()

    # actionimportf2 / actionsubmitf2
    if a.lower().startswith("action") and not a.lower().startswith("action_"):
        a = a[6:].strip()

    a = a.lower().replace("sumbit", "submit")

    # importf2 / submitf2 / import_f2 / submit_f2  -> import_f2 / submit_f2
    m = re.match(r"^(import|submit)_?f(\d+)$", a)
    if m:
        return f"{m.group(1)}_f{m.group(2)}"

    return a


def submit_f1(wb):
    push_sheet(wb, SHEET_F1, FIELDS_F1, URL_F1)


def submit_f2(wb):
    push_sheet(wb, SHEET_F2, FIELDS_F2, URL_F2)


def submit_f5(wb):
    push_sheet(wb, SHEET_F5, FIELDS_F5, URL_F5)


def main(argv=None) -> int:
    if argv is None:
        argv = sys.argv

    if len(argv) < 3:
        log("Usage: milkQuality_Forms.py <action> <workbook_path>")
        return 1

    action_raw = argv[1]
    action = normalize_action(action_raw)
    wb_path = argv[2]

    log("=== milkQuality START ===")
    log(f"action_raw={action_raw} action={action} workbook={wb_path}")

    try:
        log(f"py_file={os.path.abspath(__file__)}")
        log(f"python={sys.executable}")
        log(f"cwd={os.getcwd()}")
    except Exception:
        pass

    wb, opened_here, excel, prev_count = attach_workbook(wb_path)

    try:
        action_map = {
            "import_f1": import_f1,
            "submit_f1": submit_f1,
            "import_f2": import_f2,
            "submit_f2": submit_f2,
            "import_f5": import_f5,
            "submit_f5": submit_f5,
        }

        fn = action_map.get(action)
        if fn is None:
            log(f"Unknown action: {action}")
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
