"""LAC → BSC/RNC registry and Python helpers for paging PSR aggregation.

Each KPI row is mapped by looping decoded LAC against this registry (BSC or RNC layer).
LAC decode: GCI/SAI last-4 hex digits, else HEX2DEC(RIGHT(Object Name, 4)).
Registry LAC always wins over LABEL= text so 2G LACs (5030–5033) stay on BSC only and
never duplicate onto RNC. Labels are used only when LAC cannot be decoded.
"""

from __future__ import annotations

import re
import urllib.parse
from datetime import date

import pandas as pd
from sqlalchemy import create_engine

# (LAC decimal, controller name) — update here when the network LAC plan changes.
BSC_LAC_MAP: list[tuple[int, str]] = [
    (8026, "MBEVB02"),
    (8027, "MBEVB02"),
    (1169, "TMEVB03"),
    (1243, "SAEVB02"),
    (5222, "MBEVB01"),
    (5047, "DOEVB01"),
    (1236, "TMEVB04"),
    (5032, "DOEVB02"),
    (1204, "TMEVB01"),
    (5046, "DOEVB01"),
    (1121, "SAEVB01"),
    (1132, "SAEVB01"),
    (1209, "TMEVB01"),
    (1235, "TMEVB04"),
    (5031, "DOEVB02"),
    (5030, "DOEVB02"),
    (1205, "TMEVB01"),
    (1240, "TMEVB04"),
    (1123, "SAEVB01"),
    (1170, "TMEVB03"),
    (8053, "MWEVB01"),
    (5043, "DOEVB01"),
    (5226, "MBEVB01"),
    (1207, "TMEVB01"),
    (3011, "AREVB01"),
    (1208, "TMEVB01"),
    (5224, "MBEVB01"),
    (8042, "MWEVB02"),
    (1244, "SAEVB02"),
    (5033, "DOEVB02"),
    (1206, "TMEVB01"),
    (8021, "MBEVB02"),
    (8047, "MBEVB02"),
    (1166, "TMEVB03"),
    (8055, "MWEVB01"),
    (8045, "MWEVB02"),
    (8025, "MBEVB02"),
    (3021, "AREVB01"),
    (8024, "MBEVB02"),
    (3023, "AREVB01"),
    (1231, "TMEVB04"),
    (5240, "MBEVB01"),
    (3022, "AREVB01"),
    (8043, "MWEVB02"),
    (5221, "MBEVB01"),
    (5223, "MBEVB01"),
    (8054, "MWEVB01"),
    (3024, "AREVB01"),
    (5225, "MBEVB01"),
    (8044, "MWEVB02"),
    (1168, "TMEVB03"),
    (12021, "PMEVB01"),
    (8056, "MWEVB01"),
    (8023, "MBEVB02"),
    (1239, "TMEVB04"),
    (1237, "TMEVB04"),
    (12022, "PMEVB01"),
    (1202, "TMEVB01"),
    (1233, "TMEVB04"),
    (12025, "ZNEVB01"),
    (12020, "ZNEVB01"),
    (12024, "ZNEVB01"),
    (8051, "MWEVB01"),
    (1165, "TMEVB03"),
    (8022, "MBEVB02"),
    (1232, "TMEVB04"),
    (5042, "DOEVB01"),
    (12016, "ZNEVB01"),
    (1194, "TMEVB03"),
    (1122, "SAEVB01"),
    (12019, "ZNEVB01"),
    (1164, "TMEVB03"),
    (2102, "SAEVB02"),
    (1172, "SAEVB01"),
    (1214, "TMEVB03"),
    (1143, "SAEVB02"),
    (1124, "SAEVB01"),
    (1183, "SAEVB02"),
]

RNC_LAC_MAP: list[tuple[int, str]] = [
    (6074, "DARNC06"),
    (7053, "MWEVR01"),
    (6044, "DARNC04"),
    (6032, "DARNC06"),
    (7062, "AREVR01"),
    (6045, "DARNC04"),
    (6011, "AREVR01"),
    (7041, "MWEVR02"),
    (6079, "DARNC06"),
    (6031, "DARNC06"),
    (7023, "MBEVR02"),
    (6064, "DARNC05"),
    (7043, "MWEVR02"),
    (7022, "MBEVR01"),
    (6092, "DOEVR01"),
    (6332, "PBEVR01"),
    (7031, "DOEVR01"),
    # LACs 5030-5033 are 2G only — mapped on BSC (DOEVB02), not on RNC pivot.
    (6334, "ZNEVR01"),
    (7051, "MWEVR01"),
    (6042, "DARNC04"),
    (6062, "DARNC05"),
    (6061, "DARNC05"),
    (6041, "DARNC04"),
    (6333, "ZNEVR01"),
    (6081, "TMEVR04"),
    (6082, "TMEVR04"),
]

# Registry LAC allow-lists (78 BSC + 26 RNC LAC rows → 104 unique IDs for PSR per LAC).
BSC_LAC_IDS: frozenset[int] = frozenset(lac for lac, _ in BSC_LAC_MAP)
RNC_LAC_IDS: frozenset[int] = frozenset(lac for lac, _ in RNC_LAC_MAP)
REGISTERED_LAC_IDS: frozenset[int] = BSC_LAC_IDS | RNC_LAC_IDS

_NEW_LAC_LABEL_RE = re.compile(r"^NEW_LAC[_-](\d+)$", re.IGNORECASE)

BSC_CONTROLLERS: list[str] = [
    "AREVB01",
    "DOEVB01",
    "DOEVB02",
    "MBEVB01",
    "MBEVB02",
    "MWEVB01",
    "MWEVB02",
    "PMEVB01",
    "SAEVB01",
    "SAEVB02",
    "TMEVB01",
    "TMEVB03",
    "TMEVB04",
    "ZNEVB01",
]

RNC_CONTROLLERS: list[str] = [
    "AREVR01",
    "DARNC04",
    "DARNC05",
    "DARNC06",
    "DOEVR01",
    "MBEVR01",
    "MBEVR02",
    "MWEVR01",
    "MWEVR02",
    "PBEVR01",
    "TMEVR04",
    "ZNEVR01",
]

# Map legacy label names that still appear in KPI Object Name.
BSC_LABEL_ALIASES: dict[str, str] = {
    "PEMBA_TEST": "PMEVB01",
    "TMEVB05": "TMEVB03",
    "NEW_LAC_DOEVB02": "DOEVB02",
    "NEW_LAC_DOEVB01": "DOEVB01",
}

RNC_LABEL_ALIASES: dict[str, str] = {
    "PEMBA_TEST": "PBEVR01",
}

# Reference thresholds from dashboard (Ref Thr row).
BSC_REF_THRESHOLDS: dict[str, float] = {
    "AREVB01": 92.800,
    "DOEVB01": 92.300,
    "DOEVB02": 92.300,
    "MBEVB01": 92.300,
    "MBEVB02": 92.300,
    "MWEVB01": 93.612,
    "MWEVB02": 92.300,
    "PMEVB01": 94.400,
    "SAEVB01": 92.700,
    "SAEVB02": 95.567,
    "TMEVB01": 92.300,
    "TMEVB03": 94.500,
    "TMEVB04": 93.000,
    "ZNEVB01": 96.000,
}

# Excel paste layout (template had an extra column N after ZNEVR01 for removed DOEVR02).
RNC_PIVOT_COLUMNS: list[str] = ["Row Labels", *RNC_CONTROLLERS]
BSC_PIVOT_COLUMNS: list[str] = ["Row Labels", *BSC_CONTROLLERS]
RNC_PASTE_CLEAR_THROUGH_COL = 14
BSC_PASTE_CLEAR_THROUGH_COL = 16

RNC_REF_THRESHOLDS: dict[str, float] = {
    "AREVR01": 95.842,
    "DARNC04": 96.983,
    "DARNC05": 97.056,
    "DARNC06": 95.719,
    "DOEVR01": 96.777,
    "MBEVR01": 96.477,
    "MBEVR02": 96.148,
    "MWEVR01": 96.539,
    "MWEVR02": 96.404,
    "PBEVR01": 97.322,
    "TMEVR04": 97.929,
    "ZNEVR01": 97.605,
}

_KPI_TABLE = "[CS_KPIS].[dbo].[host03_pmresult_1912976119]"

# Rolling window for Overall PSR + PSR per Service charts ("30 Days" on dashboard).
PSR_ROLLING_DAYS = 30

# Yearly / long trend on customer PSR&AVAIL sheet (chart range up to row 960).
PSR_YEARLY_DAYS = 365


def reporting_today() -> date:
    """Local calendar date used to exclude partial 'today' rows from charts."""
    return pd.Timestamp.now().normalize().date()


def reporting_yesterday() -> date:
    return (pd.Timestamp.now().normalize() - pd.Timedelta(days=1)).date()


def sql_date_before_today() -> str:
    """Strict upper bound: never include the current calendar day."""
    return "CAST([Result Time] AS DATE) < CAST(GETDATE() AS DATE)"


def sql_date_window_rolling(days: int = PSR_ROLLING_DAYS) -> str:
    """From (today - days) through yesterday inclusive."""
    return f"""CAST([Result Time] AS DATE) >= CAST(DATEADD(DAY, -{days}, GETDATE()) AS DATE)
      AND {sql_date_before_today()}"""


def sql_date_yesterday_only() -> str:
    return f"""CAST([Result Time] AS DATE) = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE)
      AND {sql_date_before_today()}"""


def exclude_today_rows(df: pd.DataFrame, date_col: str) -> pd.DataFrame:
    """Drop any row dated today or later (guards against SQL/client timezone drift)."""
    if df.empty or date_col not in df.columns:
        return df
    today = reporting_today()
    days = pd.to_datetime(df[date_col]).dt.date
    out = df.loc[days < today].copy()
    dropped = len(df) - len(out)
    if dropped:
        print(f"Excluded {dropped} row(s) on or after today ({today}) from {date_col}.")
    return out.reset_index(drop=True)


def _read_sql(sql_conn: str, sql: str) -> pd.DataFrame:
    params = urllib.parse.quote_plus(sql_conn)
    engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}", fast_executemany=True)
    with engine.connect() as conn:
        return pd.read_sql(sql, conn)


def _lac_map_dict(lac_map: list[tuple[int, str]]) -> dict[int, str]:
    return dict(lac_map)


def _extract_hex(object_name: str, key: str) -> str | None:
    if not object_name:
        return None
    marker = f"{key}="
    idx = object_name.find(marker)
    if idx < 0:
        return None
    start = idx + len(marker)
    end = object_name.find(",", start)
    token = object_name[start:end].strip() if end >= 0 else object_name[start:].strip()
    return token or None


def _hex_suffix_to_lac(token: str) -> int | None:
    if not token or len(token) < 4:
        return None
    suffix = token[-4:]
    if not all(ch in "0123456789ABCDEFabcdef" for ch in suffix):
        return None
    try:
        return int(suffix, 16)
    except ValueError:
        return None


def _lac_from_last_hex_token(object_name: str) -> int | None:
    """HEX2DEC(RIGHT(hex,4)) on the last GCI/SAI-like token — not the full Object Name path."""
    matches = re.findall(r"[0-9A-Fa-f]{6,}", object_name)
    if not matches:
        return None
    return _hex_suffix_to_lac(matches[-1])


def decode_lac_from_object_name(object_name: str) -> int | None:
    """GCI/SAI last-4 hex first, then last hex token in the string (not path/label suffix)."""
    if not object_name:
        return None
    for key in ("GCI", "SAI"):
        token = _extract_hex(object_name, key)
        lac = _hex_suffix_to_lac(token) if token else None
        if lac is not None:
            return lac
    return _lac_from_last_hex_token(object_name)


def trim_pivot_columns(pivot: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Keep exactly one column per controller; drop legacy/extra pivot columns."""
    out = pivot.loc[:, ~pivot.columns.duplicated()].copy()
    for col in columns:
        if col not in out.columns:
            out[col] = None
    return out[columns]


def extract_label_from_object_name(object_name: str) -> str | None:
    if not object_name:
        return None
    marker = "LABEL="
    idx = object_name.find(marker)
    if idx < 0:
        return None
    start = idx + len(marker)
    end = object_name.find(",", start)
    label = object_name[start:end].strip() if end >= 0 else object_name[start:].strip()
    return label or None


def extract_registered_lac(object_name: str) -> int | None:
    """Return LAC only if it appears in BSC_LAC_MAP or RNC_LAC_MAP."""
    return match_layer_lac(object_name, REGISTERED_LAC_IDS)


def match_layer_lac(object_name: str, layer_lac_ids: frozenset[int]) -> int | None:
    """Decode LAC (GCI/SAI hex, then name suffix) and keep only registry IDs for this layer."""
    lac = decode_lac_from_object_name(object_name)
    if lac is not None:
        return lac if lac in layer_lac_ids else None
    label = extract_label_from_object_name(object_name)
    if label:
        match = _NEW_LAC_LABEL_RE.match(label)
        if match:
            lac = int(match.group(1))
            return lac if lac in layer_lac_ids else None
    return None


def _label_only_controller(
    extracted_label: str,
    controllers: list[str],
    label_aliases: dict[str, str],
    pemba_test_label: str | None,
) -> str | None:
    label = parse_label_controller(extracted_label, pemba_test_label)
    if label in controllers:
        resolved = label
    else:
        resolved = label_contains_controller(extracted_label, controllers)
    if not resolved:
        return None
    return label_aliases.get(resolved, resolved)


def parse_label_controller(extracted_label: str, pemba_test_label: str | None) -> str | None:
    if not extracted_label:
        return None
    if extracted_label == "PEMBA_TEST" and pemba_test_label:
        return pemba_test_label
    if extracted_label.startswith("NEW_LAC_") or extracted_label.startswith("NEW_LAC-"):
        return extracted_label[8:]
    us = extracted_label.rfind("_")
    ds = extracted_label.rfind("-")
    sep = max(us, ds)
    if sep < 0:
        return extracted_label
    if us >= 0 and ds >= 0:
        return extracted_label[sep + 1 :]
    return extracted_label[sep + 1 :] if sep >= 0 else extracted_label


def label_contains_controller(extracted_label: str, controllers: list[str]) -> str | None:
    for ctrl in sorted(controllers, key=len, reverse=True):
        if ctrl in extracted_label:
            return ctrl
    return None


def resolve_controller(
    extracted_label: str,
    object_name: str,
    lac_map: dict[int, str],
    layer_lac_ids: frozenset[int],
    controllers: list[str],
    label_aliases: dict[str, str],
    pemba_test_label: str | None,
) -> str | None:
    """Map one KPI row to a controller — LAC registry wins; labels only when LAC is absent.

    Prevents BSC/RNC double-count: a row whose decoded LAC belongs to the other layer
    (e.g. 5030 on BSC only) is excluded from this pivot, even if the label mentions
    DOEVB02 / NEW_LAC_DOEVB02.
    """
    layer_lac = match_layer_lac(object_name, layer_lac_ids)
    if layer_lac is not None:
        ctrl = lac_map.get(layer_lac)
        if ctrl and ctrl in controllers:
            return label_aliases.get(ctrl, ctrl)
        return None

    decoded_lac = decode_lac_from_object_name(object_name)
    if decoded_lac is not None and decoded_lac in REGISTERED_LAC_IDS:
        return None

    label_lac = None
    if extracted_label:
        match = _NEW_LAC_LABEL_RE.match(extracted_label)
        if match:
            candidate = int(match.group(1))
            if candidate in REGISTERED_LAC_IDS:
                label_lac = candidate
    if label_lac is not None and label_lac not in layer_lac_ids:
        return None

    return _label_only_controller(
        extracted_label, controllers, label_aliases, pemba_test_label
    )


def log_registry_resolution_stats(
    raw: pd.DataFrame,
    work: pd.DataFrame,
    *,
    layer_name: str,
) -> None:
    """Summarize how rows were mapped (registry LAC vs label fallback vs dropped)."""
    total = len(raw)
    mapped = len(work)
    dropped = total - mapped
    if total == 0:
        return
    by_lac = int(work["MappedLAC"].notna().sum()) if "MappedLAC" in work.columns else 0
    by_label = mapped - by_lac
    print(
        f"{layer_name} mapping: {mapped}/{total} rows kept "
        f"({by_lac} by LAC registry, {by_label} by label fallback, {dropped} excluded)"
    )


def log_psr_threshold_gaps(
    pivot: pd.DataFrame,
    thresholds: dict[str, float],
    report_name: str,
    *,
    max_lines: int = 12,
) -> None:
    """Print site-days below Ref Thr (matches dashboard breach review)."""
    if pivot.empty:
        return
    breaches: list[tuple[object, str, float, float, float]] = []
    for _, row in pivot.iterrows():
        day = row.get("Row Labels")
        for site, thr in thresholds.items():
            if site not in pivot.columns:
                continue
            val = row[site]
            if pd.isna(val):
                continue
            psr = float(val)
            if psr < thr:
                breaches.append((day, site, psr, thr, psr - thr))
    if not breaches:
        print(f"{report_name}: all exported values meet Ref Thr.")
        return
    breaches.sort(key=lambda item: item[4])
    print(f"{report_name}: {len(breaches)} site-day(s) below Ref Thr (worst first):")
    for day, site, psr, thr, gap in breaches[:max_lines]:
        print(f"  {day} {site}: {psr:.3f}% < {thr:.3f}% (gap {gap:.3f}%)")
    if len(breaches) > max_lines:
        print(f"  ... and {len(breaches) - max_lines} more")


def overall_psr_network_sql(rolling_days: int = PSR_YEARLY_DAYS) -> str:
    """Network Overall PSR for PSR&AVAIL column B (yearly trend chart)."""
    return f"""
SELECT
    CAST([Result Time] AS DATE) AS [Row Labels],
    CAST(100.0 * (
        SUM(ISNULL(TRY_CAST([Number of First Paging Responses from A Interface] AS BIGINT),0)) +
        SUM(ISNULL(TRY_CAST([Number of Repeated Paging Responses from A Interface] AS BIGINT),0)) +
        SUM(ISNULL(TRY_CAST([Number of First Paging Responses from Iu Interface] AS BIGINT),0)) +
        SUM(ISNULL(TRY_CAST([Number of Repeated Paging Responses from Iu Interface] AS BIGINT),0)) +
        SUM(ISNULL(TRY_CAST([Number of First Paging Responses from SGs Interface] AS BIGINT),0)) +
        SUM(ISNULL(TRY_CAST([Number of Second Paging Responses from SGs Interface] AS BIGINT),0)) +
        SUM(ISNULL(TRY_CAST([Number of Third Paging Responses from SGs Interface] AS BIGINT),0))
    ) / NULLIF((
        SUM(ISNULL(TRY_CAST([Number of First Pagings to A Interface] AS BIGINT),0)) +
        SUM(ISNULL(TRY_CAST([Number of First Pagings to Iu Interface] AS BIGINT),0)) +
        SUM(ISNULL(TRY_CAST([Number of First Pagings to SGs Interface] AS BIGINT),0))
    ), 0) AS DECIMAL(10,3)) AS Overall_PSR
FROM {_KPI_TABLE}
WHERE {sql_date_window_rolling(rolling_days)}
GROUP BY CAST([Result Time] AS DATE)
ORDER BY [Row Labels] ASC;
"""


def fetch_overall_psr_network(sql_conn: str, rolling_days: int = PSR_YEARLY_DAYS) -> pd.DataFrame:
    df = _read_sql(sql_conn, overall_psr_network_sql(rolling_days))
    return exclude_today_rows(df, "Row Labels")


def _build_raw_psr_sql(*, include_gs_pag_succ: bool, rolling_days: int = PSR_ROLLING_DAYS) -> str:
    gs_att = (
        " + ISNULL(TRY_CAST([Number of First Pagings to Gs Interface] AS BIGINT), 0)"
        if include_gs_pag_succ
        else ""
    )
    gs_succ = (
        " + ISNULL(TRY_CAST([Number of First Paging Responses from Gs Interface] AS BIGINT), 0)"
        if include_gs_pag_succ
        else ""
    )
    return f"""
SELECT
    CAST([Result Time] AS DATE) AS [Date],
    SUBSTRING(
        [Object Name],
        CHARINDEX('LABEL=', [Object Name]) + 6,
        CHARINDEX(',', [Object Name], CHARINDEX('LABEL=', [Object Name]))
            - (CHARINDEX('LABEL=', [Object Name]) + 6)
    ) AS ExtractedLabel,
    [Object Name] AS ObjectName,
    ISNULL(TRY_CAST([Number of First Pagings to A Interface] AS BIGINT), 0) +
    ISNULL(TRY_CAST([Number of First Pagings to Iu Interface] AS BIGINT), 0) +
    ISNULL(TRY_CAST([Number of First Pagings to SGs Interface] AS BIGINT), 0){gs_att} AS PAG_ATT,
    ISNULL(TRY_CAST([Number of First Paging Responses from A Interface] AS BIGINT), 0) +
    ISNULL(TRY_CAST([Number of Repeated Paging Responses from A Interface] AS BIGINT), 0) +
    ISNULL(TRY_CAST([Number of First Paging Responses from Iu Interface] AS BIGINT), 0) +
    ISNULL(TRY_CAST([Number of Repeated Paging Responses from Iu Interface] AS BIGINT), 0) +
    ISNULL(TRY_CAST([Number of First Paging Responses from SGs Interface] AS BIGINT), 0) +
    ISNULL(TRY_CAST([Number of Second Paging Responses from SGs Interface] AS BIGINT), 0) +
    ISNULL(TRY_CAST([Number of Third Paging Responses from SGs Interface] AS BIGINT), 0){gs_succ} AS PAG_SUCC
FROM {_KPI_TABLE}
WHERE {sql_date_window_rolling(rolling_days)}
  AND CHARINDEX('LABEL=', [Object Name]) > 0
"""


def pivot_daily_psr(
    raw: pd.DataFrame,
    lac_map: list[tuple[int, str]],
    layer_lac_ids: frozenset[int],
    controllers: list[str],
    label_aliases: dict[str, str],
    *,
    layer_name: str = "Pivot",
    pemba_test_label: str | None = None,
) -> pd.DataFrame:
    """Aggregate raw KPI rows to daily PSR pivot (one row per date)."""
    lac_dict = _lac_map_dict(lac_map)
    work = raw.copy()
    obj_col = work["ObjectName"].astype(str)
    work["MappedLAC"] = obj_col.map(lambda name: match_layer_lac(name, layer_lac_ids))
    work["CleanLabel"] = work.apply(
        lambda r: resolve_controller(
            str(r["ExtractedLabel"]) if pd.notna(r["ExtractedLabel"]) else "",
            str(r["ObjectName"]) if pd.notna(r["ObjectName"]) else "",
            lac_dict,
            layer_lac_ids,
            controllers,
            label_aliases,
            pemba_test_label,
        ),
        axis=1,
    )
    log_registry_resolution_stats(raw, work[work["CleanLabel"].isin(controllers)], layer_name=layer_name)
    work = work[work["CleanLabel"].isin(controllers)]
    if work.empty:
        return pd.DataFrame(columns=["Row Labels", *controllers])

    agg = (
        work.groupby(["Date", "CleanLabel"], as_index=False)
        .agg(PAG_ATT=("PAG_ATT", "sum"), PAG_SUCC=("PAG_SUCC", "sum"))
    )
    agg["PSR"] = agg.apply(
        lambda r: round(100.0 * r["PAG_SUCC"] / r["PAG_ATT"], 3) if r["PAG_ATT"] > 0 else 0.0,
        axis=1,
    )
    pivot = agg.pivot(index="Date", columns="CleanLabel", values="PSR").reset_index()
    pivot = pivot.rename(columns={"Date": "Row Labels"})
    for col in controllers:
        if col not in pivot.columns:
            pivot[col] = None
    expected = ["Row Labels", *controllers]
    pivot = trim_pivot_columns(pivot, expected)
    return pivot.sort_values("Row Labels", ascending=False).reset_index(drop=True)


def fetch_daily_psr_pivot(
    sql_conn: str,
    lac_map: list[tuple[int, str]],
    layer_lac_ids: frozenset[int],
    controllers: list[str],
    label_aliases: dict[str, str],
    *,
    layer_name: str = "Pivot",
    pemba_test_label: str | None = None,
    include_gs_pag_succ: bool = False,
    rolling_days: int = PSR_ROLLING_DAYS,
) -> pd.DataFrame:
    """Fetch raw rows with simple SQL; map LAC/labels and pivot in Python."""
    raw = _read_sql(
        sql_conn,
        _build_raw_psr_sql(include_gs_pag_succ=include_gs_pag_succ, rolling_days=rolling_days),
    )
    raw = exclude_today_rows(raw, "Date")
    pivot = pivot_daily_psr(
        raw,
        lac_map,
        layer_lac_ids,
        controllers,
        label_aliases,
        layer_name=layer_name,
        pemba_test_label=pemba_test_label,
    )
    return exclude_today_rows(pivot, "Row Labels")


def fetch_daily_psr_rnc(sql_conn: str, rolling_days: int = PSR_ROLLING_DAYS) -> pd.DataFrame:
    pivot = fetch_daily_psr_pivot(
        sql_conn,
        RNC_LAC_MAP,
        RNC_LAC_IDS,
        RNC_CONTROLLERS,
        RNC_LABEL_ALIASES,
        layer_name="RNC",
        pemba_test_label="PBEVR01",
        include_gs_pag_succ=False,
        rolling_days=rolling_days,
    )
    pivot = trim_pivot_columns(pivot, RNC_PIVOT_COLUMNS)
    log_psr_threshold_gaps(pivot, RNC_REF_THRESHOLDS, "RNC PSR")
    return pivot


def fetch_daily_psr_bsc(sql_conn: str, rolling_days: int = PSR_ROLLING_DAYS) -> pd.DataFrame:
    pivot = fetch_daily_psr_pivot(
        sql_conn,
        BSC_LAC_MAP,
        BSC_LAC_IDS,
        BSC_CONTROLLERS,
        BSC_LABEL_ALIASES,
        layer_name="BSC",
        pemba_test_label="PMEVB01",
        include_gs_pag_succ=True,
        rolling_days=rolling_days,
    )
    pivot = trim_pivot_columns(pivot, BSC_PIVOT_COLUMNS)
    log_psr_threshold_gaps(pivot, BSC_REF_THRESHOLDS, "BSC PSR")
    return pivot


def all_registered_lac_ids() -> list[int]:
    return sorted(REGISTERED_LAC_IDS)


_PSR_PER_LAC_SQL = f"""
SELECT
    [Result Time] AS [Time],
    [Object Name] AS [LAI],
    (ISNULL(TRY_CAST([Number of First Pagings to A Interface] AS BIGINT), 0) +
     ISNULL(TRY_CAST([Number of First Pagings to Iu Interface] AS BIGINT), 0) +
     ISNULL(TRY_CAST([Number of First Pagings to SGs Interface] AS BIGINT), 0)) AS PAG_ATT,
    (ISNULL(TRY_CAST([Number of First Paging Responses from A Interface] AS BIGINT), 0) +
     ISNULL(TRY_CAST([Number of Repeated Paging Responses from A Interface] AS BIGINT), 0) +
     ISNULL(TRY_CAST([Number of First Paging Responses from Iu Interface] AS BIGINT), 0) +
     ISNULL(TRY_CAST([Number of Repeated Paging Responses from Iu Interface] AS BIGINT), 0) +
     ISNULL(TRY_CAST([Number of First Paging Responses from SGs Interface] AS BIGINT), 0) +
     ISNULL(TRY_CAST([Number of Second Paging Responses from SGs Interface] AS BIGINT), 0) +
     ISNULL(TRY_CAST([Number of Third Paging Responses from SGs Interface] AS BIGINT), 0)) AS PAG_SUC
FROM {_KPI_TABLE}
WHERE {sql_date_yesterday_only()}
  AND CHARINDEX('LABEL=', [Object Name]) > 0
"""


def _psr_per_service_sql(rolling_days: int = PSR_ROLLING_DAYS) -> str:
    return f"""
SELECT
    CAST([Result Time] AS DATE) AS [Row Labels],
    CAST(100.0 * (
        SUM(ISNULL(TRY_CAST([Number of Responses to First Pagings in Calls] AS BIGINT), 0)) +
        SUM(ISNULL(TRY_CAST([Number of Responses to Repeated Pagings in Calls] AS BIGINT), 0))
    ) / NULLIF(SUM(ISNULL(TRY_CAST([Number of First Pagings in Calls] AS BIGINT), 0)), 0) AS DECIMAL(10, 3)) AS PSR_Calls,
    CAST(100.0 * (
        SUM(ISNULL(TRY_CAST([Number of Responses to First Pagings in Short Message Service] AS BIGINT), 0)) +
        SUM(ISNULL(TRY_CAST([Number of Responses to Repeated Pagings in Short Message Service] AS BIGINT), 0))
    ) / NULLIF(SUM(ISNULL(TRY_CAST([Number of First Pagings in Short Message Service] AS BIGINT), 0)), 0) AS DECIMAL(10, 3)) AS PSR_sms,
    CAST(100.0 * (
        SUM(ISNULL(TRY_CAST([Number of Responses to First Pagings in Call Independent Supplementary Services] AS BIGINT), 0)) +
        SUM(ISNULL(TRY_CAST([Number of Responses to Repeated Pagings in Call Independent Supplementary Services] AS BIGINT), 0))
    ) / NULLIF(SUM(ISNULL(TRY_CAST([Number of First Pagings in Call Independent Supplementary Services] AS BIGINT), 0)), 0) AS DECIMAL(10, 3)) AS PSR_ussd,
    CAST(100.0 * (
        SUM(ISNULL(TRY_CAST([Number of Responses to First Pagings in PSI Service] AS BIGINT), 0)) +
        SUM(ISNULL(TRY_CAST([Number of Responses to Repeated Pagings in PSI Service] AS BIGINT), 0))
    ) / NULLIF(SUM(ISNULL(TRY_CAST([Number of First Pagings in PSI Service] AS BIGINT), 0)), 0) AS DECIMAL(10, 3)) AS PSR_psi
FROM {_KPI_TABLE}
WHERE {sql_date_window_rolling(rolling_days)}
GROUP BY CAST([Result Time] AS DATE)
ORDER BY [Row Labels] ASC;
"""


def fetch_psr_per_service(sql_conn: str) -> pd.DataFrame:
    """Daily PSR by service (Calls/SMS/USSD/PSI) for the rolling chart at AA2."""
    df = _read_sql(sql_conn, _psr_per_service_sql())
    if df.empty:
        print("PSR per Service: no rows returned from KPI table for date window.")
        return df

    df = exclude_today_rows(df, "Row Labels")
    df = df.sort_values("Row Labels", ascending=True).reset_index(drop=True)
    min_day = pd.to_datetime(df["Row Labels"].iloc[0]).date()
    max_day = pd.to_datetime(df["Row Labels"].iloc[-1]).date()
    print(
        f"PSR per Service: {len(df)} day(s) from {min_day} to {max_day} "
        f"(expected ~{PSR_ROLLING_DAYS} for dashboard chart)"
    )
    if len(df) < 7:
        print(
            "WARNING: PSR per Service has very few days - chart may look flat or empty. "
            "Check KPI collection for service counters."
        )
    return df


def fetch_psr_per_lac(sql_conn: str) -> pd.DataFrame:
    """One row per registry LAC for yesterday only (max 104 unique IDs from BSC+RNC maps)."""
    empty = pd.DataFrame(columns=["Time", "LAI", "LAC (Dec)", "PSR (%)"])
    raw = _read_sql(sql_conn, _PSR_PER_LAC_SQL)
    if raw.empty:
        return empty

    yesterday = reporting_yesterday()
    today = reporting_today()
    raw["_day"] = pd.to_datetime(raw["Time"]).dt.date
    raw = raw[(raw["_day"] == yesterday) & (raw["_day"] < today)]
    if raw.empty:
        return empty

    raw["LAC (Dec)"] = raw["LAI"].astype(str).map(
        lambda name: match_layer_lac(name, REGISTERED_LAC_IDS)
    )
    raw = raw[raw["LAC (Dec)"].notna()]
    raw["LAC (Dec)"] = raw["LAC (Dec)"].astype(int)
    raw = raw[raw["LAC (Dec)"].isin(REGISTERED_LAC_IDS)]
    if raw.empty:
        return empty

    # Same LAC can appear under many Object Names — sum counters, one row per LAC.
    agg = (
        raw.groupby("LAC (Dec)", as_index=False)
        .agg(
            PAG_ATT=("PAG_ATT", "sum"),
            PAG_SUC=("PAG_SUC", "sum"),
            Time=("Time", "max"),
            LAI=("LAI", "first"),
        )
    )
    agg = agg[agg["PAG_ATT"] > 0]
    if agg.empty:
        return empty

    agg["PSR (%)"] = (100.0 * agg["PAG_SUC"] / agg["PAG_ATT"]).round(2)
    out = agg[["Time", "LAI", "LAC (Dec)", "PSR (%)"]].drop_duplicates(
        subset=["LAC (Dec)"], keep="first"
    )
    out = out[out["LAC (Dec)"].isin(REGISTERED_LAC_IDS)]
    print(
        f"PSR per LAC: {len(out)} row(s) for {yesterday} "
        f"(registry {len(REGISTERED_LAC_IDS)} = {len(BSC_LAC_IDS)} BSC + {len(RNC_LAC_IDS)} RNC, "
        f"{len(BSC_LAC_IDS) + len(RNC_LAC_IDS) - len(REGISTERED_LAC_IDS)} overlap)"
    )
    return out.sort_values("PSR (%)", ascending=True).reset_index(drop=True)
