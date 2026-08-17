from __future__ import annotations

from io import BytesIO
from pathlib import Path
import tempfile
import unicodedata

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill


st.set_page_config(page_title="Concentrados de cartera", page_icon="📊", layout="wide")

SUM_COLUMNS = [
    "Clientes Totales", "Clientes al corriente", "Faltas", "Distribuidoras al corriente",
    "FP", "FP Al Corriente", "FP Atraso", "PP", "PP Al Corriente", "PP Atraso",
    "Cartera Total", "Cartera sin atrasos",
    "CO Renov $", "CO Nuevos $", "CO Reconquista $",
    "CO Renov #", "CO Nuevos #", "CO Reconquista #", "Nunca Abonada",
]
MAX_COLUMNS = [
    "Coord Principal", "Coord Nueva", "Cambio de Coordinadora", "dias_de_atraso",
]

CANONICAL_COLUMNS = [
    "Territorio", "Subdireccion", "Zona", "Sucursal", "Unidad de negocio",
    "Unidad de Negocio", "Base", "Pais", "País", "Corte", "Marca", "ruta",
    "tipo_desembolso", "MONTO_DESEMBOLSO", "fecha_desembolso", "dias_de_atraso",
    "colocado_ci", "colocado_pp", "colocado_con_interes_fp", "CAMBIO_COORDINAODRA",
    "coordinadora_prinsipal", "fecha_cambio_coordinadora", "pagosRealizados", "id_y_localidad",
    "Clientes Totales", "Clientes al corriente", "Faltas", "FP", "FP Al Corriente",
    "FP Atraso", "PP", "PP Al Corriente", "PP Atraso", "Cartera Total",
    "Cartera sin atrasos", "Coord Principal", "CO Renov $", "CO Nuevos $",
    "CO Reconquista $", "CO Renov #", "CO Nuevos #", "CO Reconquista #",
    "Coord Nueva", "Cambio de Coordinadora", "Nunca Abonada",
    "Fecha de Corte", "Distribuidora", "Coordinacion", "Número",
    "Clientes con Compras Pendientes", "Clientes al Corriente", "Clientes en atraso",
    "Mora Máxima", "Colocado Neto VA",
]


def clean_name(value: object) -> str:
    return " ".join(str(value).strip().split())


def name_key(value: object) -> str:
    text = unicodedata.normalize("NFKD", clean_name(value)).encode("ascii", "ignore").decode()
    return text.casefold()


CANONICAL_BY_KEY = {name_key(c): c for c in CANONICAL_COLUMNS}
CANONICAL_BY_KEY.update({
    "pais": "Pais", "unidad de negocio": "Unidad de negocio", "ruta": "ruta",
    "zona": "Zona", "marca": "Marca",
})


def normalize_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [CANONICAL_BY_KEY.get(name_key(c), clean_name(c)) for c in df.columns]
    return df.loc[:, ~df.columns.duplicated()]


def _temporary_book(content: bytes, suffix: str) -> Path:
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(content)
        return Path(tmp.name)


@st.cache_data(show_spinner=False)
def sheet_names(content: bytes, suffix: str) -> list[str]:
    path = _temporary_book(content, suffix)
    engine = "pyxlsb" if suffix.lower() == ".xlsb" else "openpyxl"
    try:
        with pd.ExcelFile(path, engine=engine) as book:
            return list(book.sheet_names)
    finally:
        path.unlink(missing_ok=True)


@st.cache_data(show_spinner=False)
def read_sheet(content: bytes, suffix: str, sheet_name: str) -> pd.DataFrame:
    path = _temporary_book(content, suffix)
    engine = "pyxlsb" if suffix.lower() == ".xlsb" else "openpyxl"
    try:
        return normalize_frame(pd.read_excel(path, sheet_name=sheet_name, engine=engine))
    finally:
        path.unlink(missing_ok=True)


@st.cache_data(show_spinner=False)
def load_structure() -> pd.DataFrame:
    path = Path(__file__).with_name("estructura.csv")
    if not path.exists():
        raise FileNotFoundError("Falta estructura.csv junto a app.py")
    return normalize_frame(pd.read_csv(path, encoding="utf-8-sig"))


def numeric(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series(0.0, index=df.index)
    return pd.to_numeric(df[column], errors="coerce").fillna(0)


def column_name(df: pd.DataFrame, *candidates: str) -> str | None:
    """Encuentra una columna sin depender de acentos, mayúsculas o espacios."""
    available = {name_key(column): column for column in df.columns}
    return next((available.get(name_key(candidate)) for candidate in candidates), None)


def excel_datetime(values: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(values):
        return pd.to_datetime(values, errors="coerce")
    result = pd.Series(pd.NaT, index=values.index, dtype="datetime64[ns]")
    numbers = pd.to_numeric(values, errors="coerce")
    serials = numbers.between(20_000, 80_000)
    result.loc[serials] = pd.Timestamp("1899-12-30") + pd.to_timedelta(numbers.loc[serials], unit="D")
    result.loc[~serials] = pd.to_datetime(values.loc[~serials], errors="coerce")
    return result


def adjusted_year(values: pd.Series) -> pd.Series:
    years = values.dt.year.astype("Int64")
    special = values.dt.strftime("%Y-%m-%d").isin(["2024-12-30", "2024-12-31"])
    return years.mask(special, 2025)


def same_excel_week(left: pd.Series, right: pd.Series) -> pd.Series:
    valid = left.notna() & right.notna()
    return (
        valid
        & left.dt.isocalendar().week.eq(right.dt.isocalendar().week)
        & adjusted_year(left).eq(adjusted_year(right))
    )


def add_structure(raw: pd.DataFrame, structure: pd.DataFrame) -> pd.DataFrame:
    if all(c in raw.columns for c in ["Territorio", "Subdireccion", "Zona", "Sucursal", "Unidad de negocio"]):
        return raw
    catalog = structure.rename(columns={
        "Ruta": "ruta", "Base": "Unidad de negocio", "País": "Pais",
        "Unidad de Negocio": "Unidad LATAM",
    })
    catalog = catalog.loc[:, ~catalog.columns.duplicated()]
    targets = ["Territorio", "Subdireccion", "Zona", "Sucursal", "Unidad de negocio", "Pais"]
    add = [c for c in targets if c in catalog.columns and c not in raw.columns]
    if "ruta" not in raw.columns:
        raise ValueError("No se encontró la columna 'ruta' en la hoja seleccionada.")
    catalog = catalog[["ruta", *add]].drop_duplicates(subset=["ruta"])
    result = raw.merge(catalog, on="ruta", how="left", validate="many_to_one")
    return result


def add_excel_calculations(base: pd.DataFrame, profile: str) -> pd.DataFrame:
    required = [
        "Corte", "tipo_desembolso", "fecha_desembolso", "MONTO_DESEMBOLSO",
        "dias_de_atraso", "colocado_ci", "colocado_pp", "CAMBIO_COORDINAODRA",
        "coordinadora_prinsipal", "fecha_cambio_coordinadora", "pagosRealizados", "id_y_localidad",
    ]
    missing = [c for c in required if c not in base.columns]
    if missing:
        raise ValueError("Faltan columnas del reporte: " + ", ".join(missing))

    result = base.copy()
    loan_type = result["tipo_desembolso"].fillna("").astype(str).str.strip()
    loan_key = loan_type.str.casefold()
    days = numeric(result, "dias_de_atraso")
    is_fp = loan_key.eq("financiero personal")
    is_pp = loan_key.eq("prestamo personal")

    if profile == "LATAM":
        result["Clientes Totales"] = (~is_fp & ~is_pp).astype(int)
        result["FP"] = is_fp.astype(int)
        result["FP Al Corriente"] = (is_fp & days.eq(0)).astype(int)
        result["FP Atraso"] = (is_fp & days.gt(0)).astype(int)
        result["PP"] = is_pp.astype(int)
        result["PP Al Corriente"] = (is_pp & days.eq(0)).astype(int)
        result["PP Atraso"] = (is_pp & days.gt(0)).astype(int)
    else:
        result["Clientes Totales"] = (~is_fp).astype(int)

    result["Clientes al corriente"] = (result["Clientes Totales"].eq(1) & days.eq(0)).astype(int)
    result["Faltas"] = (result["Clientes Totales"].eq(1) & days.gt(0)).astype(int)

    standard_portfolio = numeric(result, "colocado_ci") + numeric(result, "colocado_pp")
    result["Cartera Total"] = np.where(is_fp, numeric(result, "colocado_con_interes_fp"), standard_portfolio)
    result["Cartera sin atrasos"] = np.where(days.gt(7), 0, result["Cartera Total"])

    change_raw = result["CAMBIO_COORDINAODRA"]
    change = pd.to_numeric(change_raw, errors="coerce")
    principal_raw = result["coordinadora_prinsipal"]
    principal = pd.to_numeric(principal_raw, errors="coerce")
    blank_principal = principal_raw.isna() | principal_raw.astype(str).str.strip().isin(["", "nan", "None"])
    result["Coord Principal"] = (blank_principal | principal.eq(1)).astype(int)

    cutoff = excel_datetime(result["Corte"])
    disbursement = excel_datetime(result["fecha_desembolso"])
    current_disbursement = same_excel_week(cutoff, disbursement)
    amount = numeric(result, "MONTO_DESEMBOLSO")
    kinds = [("Renov", "Renovacion"), ("Nuevos", "Normal"), ("Reconquista", "Reingreso")]
    for label, kind in kinds:
        flag = current_disbursement & loan_key.eq(kind.casefold())
        result[f"CO {label} $"] = np.where(flag, amount, 0)
        result[f"CO {label} #"] = flag.astype(int)

    change_date = excel_datetime(result["fecha_cambio_coordinadora"])
    current_change = same_excel_week(cutoff, change_date)
    result["Coord Nueva"] = (current_change & change.eq(1)).astype(int)
    result["Cambio de Coordinadora"] = (current_change & change.gt(1)).astype(int)

    payments_raw = result["pagosRealizados"]
    payments = pd.to_numeric(payments_raw, errors="coerce")
    no_payments = payments_raw.isna() | payments.fillna(0).eq(0)
    result["Nunca Abonada"] = (no_payments & days.gt(0)).astype(int)
    result["Corte"] = cutoff
    return result


def prepare_base(raw: pd.DataFrame, profile: str, structure: pd.DataFrame) -> pd.DataFrame:
    raw = normalize_frame(raw)
    already_calculated = {"Clientes Totales", "Cartera Total", "Cartera sin atrasos"}.issubset(raw.columns)

    if not already_calculated:
        if profile == "LATAM" and "Pais" not in raw.columns:
            raise ValueError("El archivo parece ser Presico porque no contiene la columna PAIS. Selecciona Presico.")
        if profile == "Presico" and "Pais" in raw.columns:
            raise ValueError("El archivo parece ser LATAM porque contiene la columna PAIS. Selecciona LATAM.")

    base = add_structure(raw, structure)
    if not already_calculated:
        base = add_excel_calculations(base, profile)
    elif "Corte" in base.columns:
        base["Corte"] = excel_datetime(base["Corte"])
    return base


def prepare_vales(raw: pd.DataFrame) -> pd.DataFrame:
    """Normaliza la cartera Vales; no requiere ni consulta la estructura."""
    result = normalize_frame(raw)
    source_columns = {
        "Corte": ("Fecha de Corte",),
        "Distribuidora": ("Distribuidora",),
        "Clientes Totales": ("Clientes con Compras Pendientes",),
        "Clientes al corriente": ("Clientes al Corriente",),
        "Faltas": ("Clientes en atraso",),
        "Status": ("Status",),
        "Status Mora VA": ("Status Mora VA",),
        "Mora Máxima": ("Mora Máxima", "Mora Maxima"),
        "Cartera Total": ("Colocado Neto VA",),
        "ruta": ("Coordinacion", "Coordinación"),
    }
    missing = [
        target for target, candidates in source_columns.items()
        if column_name(result, *candidates) is None
    ]
    if missing:
        raise ValueError(
            "Faltan columnas de la cartera Vales: " + ", ".join(missing)
        )

    for target, candidates in source_columns.items():
        source = column_name(result, *candidates)
        result[target] = result[source]

    for dimension in ["Marca", "Subdireccion", "Zona", "Sucursal"]:
        if dimension not in result.columns:
            result[dimension] = ""

    result["Corte"] = excel_datetime(result["Corte"])
    for measure in [
        "Clientes Totales", "Clientes al corriente", "Faltas", "Mora Máxima", "Cartera Total",
    ]:
        result[measure] = numeric(result, measure)

    status = result["Status"].fillna("").astype(str).str.strip()
    is_restructura = status.map(name_key).eq("restructura")
    status_mora = result["Status Mora VA"]
    status_mora_text = status_mora.fillna("").astype(str).str.strip()
    blank_status_mora = status_mora.isna() | status_mora_text.isin(["", "nan", "None"])
    status_mora_is_one = pd.to_numeric(status_mora, errors="coerce").eq(1)
    clientes_reportados = numeric(result, "Clientes al corriente")
    clientes_en_atraso = numeric(result, "Faltas")

    # Distribuidoras al corriente: filas con Status Mora VA vacío.
    result["Distribuidoras al corriente"] = blank_status_mora.astype(int)
    # Clientes al corriente y faltas excluyen las filas con Status = Restructura.
    result["Clientes al corriente"] = np.where(~is_restructura, clientes_reportados, 0)
    result["Faltas"] = np.where(~is_restructura, clientes_en_atraso, 0)
    # Cartera al corriente: Colocado Neto VA excepto las filas con Status Mora VA = 1.
    result["Cartera sin atrasos"] = np.where(
        ~status_mora_is_one, result["Cartera Total"], 0,
    )
    return result


def build_concentrado(base: pd.DataFrame, profile: str) -> pd.DataFrame:
    if profile == "LATAM":
        dims = ["Corte", "Pais", "Unidad de negocio", "Subdireccion", "Zona", "Sucursal", "ruta", "id_y_localidad"]
    else:
        dims = ["Unidad de negocio", "Territorio", "Subdireccion", "Zona", "Sucursal", "ruta", "id_y_localidad"]

    required = [*dims, "Clientes Totales", "Clientes al corriente", "Faltas", "Cartera Total", "Cartera sin atrasos"]
    missing = [c for c in required if c not in base.columns]
    if missing:
        raise ValueError("No se puede formar el concentrado. Faltan: " + ", ".join(missing))

    work = base.copy()
    sum_cols = [c for c in SUM_COLUMNS if c in work.columns]
    max_cols = [c for c in MAX_COLUMNS if c in work.columns]
    for column in sum_cols + max_cols:
        work[column] = pd.to_numeric(work[column], errors="coerce").fillna(0)

    aggregation = {c: "sum" for c in sum_cols} | {c: "max" for c in max_cols}
    result = work.groupby(dims, dropna=False, sort=False).agg(aggregation).reset_index()
    result["Calidad"] = np.divide(
        result["Cartera sin atrasos"], result["Cartera Total"],
        out=np.zeros(len(result), dtype=float), where=result["Cartera Total"].ne(0),
    )
    principal = result.get("Coord Principal", pd.Series(0, index=result.index))
    new_coordinator = result.get("Coord Nueva", pd.Series(0, index=result.index))
    coordinator_change = result.get("Cambio de Coordinadora", pd.Series(0, index=result.index))
    result["Coord Totales"] = principal

    common_rename = {
        "Clientes Totales": "Clientes totales", "FP Al Corriente": "FP al corriente",
        "FP Atraso": "FP Falta", "PP Al Corriente": "PP al corriente", "PP Atraso": "PP Falta",
        "dias_de_atraso": "Máx. días de atraso",
    }
    result = result.rename(columns=common_rename)

    if profile == "Presico":
        has_coordinator = result["Coord Totales"].eq(1)
        productive = has_coordinator & result["Clientes totales"].ge(21) & result["Calidad"].ge(0.60)
        developing = has_coordinator & result["Clientes totales"].lt(21) & result["Calidad"].ge(0.60)
        result["Máx. de Coordinadora"] = coordinator_change
        result["Máx. de Coord Nueva"] = new_coordinator
        result["Cambio Coordinadora"] = coordinator_change
        legacy_columns = [
            column for column in result.columns
            if "de Coordinadora" in column or "de Coord Nueva" in column
        ]
        result = result.drop(columns=legacy_columns)
        result["Coord prod"] = productive.astype(int)
        result["Coord en desarrollo"] = developing.astype(int)
        result["Coord impro"] = (
            has_coordinator & result["Coord prod"].eq(0) & result["Coord en desarrollo"].eq(0)
        ).astype(int)
        order = [
            "Unidad de negocio", "Territorio", "Subdireccion", "Zona", "Sucursal", "ruta", "id_y_localidad",
            "Clientes totales", "Clientes al corriente", "Faltas", "Máx. días de atraso", "Cartera Total", "Cartera sin atrasos",
            "Calidad", "CO Renov $", "CO Nuevos $", "CO Reconquista $", "CO Renov #", "CO Nuevos #",
            "CO Reconquista #", "Nunca Abonada", "Máx. de Coordinadora", "Máx. de Coord Nueva",
            "Coord Totales", "Coord prod", "Coord en desarrollo", "Coord impro",
            "Cambio Coordinadora", "Coord Nueva",
        ]
    else:
        result["Clientes totales"] = (
            result["Clientes totales"] + result.get("FP", 0) + result.get("PP", 0)
        )
        result["Clientes al corriente"] = (
            result["Clientes al corriente"]
            + result.get("FP al corriente", 0)
            + result.get("PP al corriente", 0)
        )
        result["Faltas"] = result["Faltas"] + result.get("FP Falta", 0) + result.get("PP Falta", 0)
        has_coordinator = result["Coord Totales"].eq(1)
        result["Coord prod"] = (
            has_coordinator & result["Clientes totales"].ge(21) & result["Calidad"].ge(0.60)
        ).astype(int)
        result["Coord en desarrollo"] = (
            has_coordinator & result["Clientes totales"].lt(21) & result["Calidad"].ge(0.60)
        ).astype(int)
        result["Coord impro"] = (
            has_coordinator & result["Coord prod"].eq(0) & result["Coord en desarrollo"].eq(0)
        ).astype(int)
        result["Fecha corte"] = result["Corte"]
        result["Cambio Coordinadora"] = coordinator_change
        result["País"] = result["Pais"]
        order = [
            "Fecha corte", "Unidad de negocio", "Subdireccion", "Zona", "Sucursal", "ruta", "id_y_localidad",
            "Clientes totales", "Clientes al corriente", "Faltas", "Máx. días de atraso", "Cartera Total", "Cartera sin atrasos", "Calidad",
            "CO Renov $", "CO Nuevos $", "CO Reconquista $", "CO Renov #", "CO Nuevos #",
            "CO Reconquista #", "Nunca Abonada", "Coord Totales", "Coord prod", "Coord en desarrollo",
            "Coord impro", "Cambio Coordinadora", "Coord Nueva", "País",
        ]
    return result[[c for c in order if c in result.columns]]


def build_concentrado_vales(base: pd.DataFrame) -> pd.DataFrame:
    dims = ["Corte", "Marca", "Subdireccion", "Zona", "Sucursal", "ruta", "Distribuidora"]
    measures = [
        "Distribuidoras al corriente",
        "Clientes Totales", "Clientes al corriente", "Faltas", "Mora Máxima",
        "Cartera Total", "Cartera sin atrasos",
    ]
    missing = [column for column in [*dims, *measures] if column not in base.columns]
    if missing:
        raise ValueError("No se puede formar el concentrado Vales. Faltan: " + ", ".join(missing))

    work = base.copy()
    for column in measures:
        work[column] = numeric(work, column)
    aggregation = {column: "sum" for column in measures if column != "Mora Máxima"}
    aggregation["Mora Máxima"] = "max"
    result = work.groupby(dims, dropna=False, sort=False).agg(aggregation).reset_index()
    result["Calidad"] = np.divide(
        result["Cartera sin atrasos"], result["Cartera Total"],
        out=np.zeros(len(result), dtype=float), where=result["Cartera Total"].ne(0),
    )
    result = result.rename(columns={
        "Clientes Totales": "Clientes totales",
        "Mora Máxima": "Máx. días de atraso",
    })
    order = [
        "Corte", "Marca", "Subdireccion", "Zona", "Sucursal", "ruta", "Distribuidora",
        "Distribuidoras al corriente",
        "Clientes totales", "Clientes al corriente", "Faltas", "Máx. días de atraso",
        "Cartera Total", "Cartera sin atrasos", "Calidad",
    ]
    return result[order]


def metric_total(df: pd.DataFrame, column: str) -> float:
    if column not in df.columns:
        return 0.0
    return pd.to_numeric(df[column], errors="coerce").fillna(0).sum()


def card_values(df: pd.DataFrame) -> dict[str, float]:
    productive = metric_total(df, "Coord prod")
    developing = metric_total(df, "Coord en desarrollo")
    return {
        "Clientes al corriente": metric_total(df, "Clientes al corriente"),
        "Faltas": metric_total(df, "Faltas"),
        "Clientes totales": metric_total(df, "Clientes totales"),
        "Nunca abonados": metric_total(df, "Nunca Abonada"),
        "Coord. prod. + des.": productive + developing,
        "Coord. productivas": productive,
        "Coord. en desarrollo": developing,
        "Coord. improductivas": metric_total(df, "Coord impro"),
        "Coordinadoras": metric_total(df, "Coord Totales"),
    }


def country_detail(df: pd.DataFrame, country_column: str) -> pd.DataFrame:
    rows = []
    for country, group in df.groupby(country_column, dropna=False, sort=True):
        values = card_values(group)
        rows.append({country_column: country, **values})
    return pd.DataFrame(rows)


def show_detail_cards(df: pd.DataFrame) -> None:
    values = card_values(df)

    st.markdown("#### Indicadores de clientes")
    client_columns = st.columns(4)
    for column, label in zip(
        client_columns,
        ["Clientes al corriente", "Faltas", "Clientes totales", "Nunca abonados"],
    ):
        column.metric(label, f"{values[label]:,.0f}")


    st.markdown("#### Indicadores de coordinadoras")
    coordinator_columns = st.columns(5)
    for column, label in zip(
        coordinator_columns,
        [
            "Coord. prod. + des.",
            "Coord. productivas",
            "Coord. en desarrollo",
            "Coord. improductivas",
            "Coordinadoras",
        ],
    ):
        column.metric(label, f"{values[label]:,.0f}")


def show_vales_cards(df: pd.DataFrame) -> None:
    st.markdown("#### Indicadores de distribuidoras")
    values = {
        "Distribuidoras al corriente": metric_total(df, "Distribuidoras al corriente"),
        "Clientes al corriente": metric_total(df, "Clientes al corriente"),
        "Clientes en atraso": metric_total(df, "Faltas"),
        "Clientes totales": metric_total(df, "Clientes totales"),
    }
    columns = st.columns(4)
    for column, label in zip(columns, values):
        column.metric(label, f"{values[label]:,.0f}")


@st.cache_data(show_spinner=False)
def excel_bytes(df: pd.DataFrame) -> bytes:
    stream = BytesIO()
    with pd.ExcelWriter(stream, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Concentrado")
        sheet = writer.book["Concentrado"]
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        for column in sheet.columns:
            width = min(42, max(11, max(len(str(cell.value or "")) for cell in column) + 2))
            sheet.column_dimensions[column[0].column_letter].width = width
    return stream.getvalue()


def main() -> None:
    st.title("Concentrados de cartera · LATAM, Presico y Vales")
    st.caption("Selecciona el proceso, carga el reporte y descarga el concentrado.")

    st.subheader("1. Tipo de concentrado")
    profile = st.radio(
        "Selecciona el origen de la cartera",
        options=["LATAM", "Presico", "Vales"], horizontal=True, label_visibility="collapsed",
    )

    st.subheader("2. Archivo")
    uploaded = st.file_uploader(
        f"Reporte de {profile}", type=["xlsb", "xlsx"], key=f"uploader_{profile}",
        help=(
            "Para Vales carga el archivo con la hoja de cartera. "
            "Para LATAM o Presico puede ser el reporte diario o un libro formulado con Base."
        ),
    )
    if uploaded is None:
        st.info(f"Selecciona el archivo de {profile} para continuar.")
        st.stop()

    content = uploaded.getvalue()
    suffix = Path(uploaded.name).suffix.lower()
    try:
        available_sheets = sheet_names(content, suffix)
    except Exception as exc:
        st.error(f"No fue posible abrir el archivo: {exc}")
        st.stop()

    default_sheet = "Base" if "Base" in available_sheets else available_sheets[0]
    if len(available_sheets) > 1:
        source_sheet = st.selectbox(
            "Hoja que contiene los datos",
            available_sheets,
            index=available_sheets.index(default_sheet),
        )
    else:
        source_sheet = default_sheet
        st.caption(f"Hoja detectada: {source_sheet}")

    with st.sidebar:
        st.header("Reglas de Vales" if profile == "Vales" else "Reglas de coordinadoras")
        if profile == "Vales":
            st.caption("Distribuidoras al corriente: Status Mora VA vacío. Cartera al corriente: Colocado Neto VA con Status Mora VA distinto de 1.")
        else:
            st.caption("Productiva: 21 o más clientes totales y calidad mínima de 60%.")
            st.caption("En desarrollo: menos de 21 clientes totales y calidad mínima de 60%.")

    try:
        with st.spinner(f"Procesando {profile}. En archivos grandes puede tardar unos minutos…"):
            raw = read_sheet(content, suffix, source_sheet)
            if profile == "Vales":
                base = prepare_vales(raw)
                concentrado = build_concentrado_vales(base)
            else:
                base = prepare_base(raw, profile, load_structure())
                concentrado = build_concentrado(base, profile)
    except Exception as exc:
        st.error(str(exc))
        st.info("Verifica el tipo LATAM/Presico y la hoja seleccionada.")
        st.stop()

    group_label = "distribuidoras" if profile == "Vales" else "localidades"
    st.success(f"{profile}: {len(base):,} registros procesados · {len(concentrado):,} {group_label}")

    filter_columns = [
        c for c in ["Marca", "País", "Pais", "Territorio", "Unidad de negocio", "Subdireccion", "Zona", "Sucursal", "ruta"]
        if c in concentrado.columns
    ]
    if profile == "Vales" and "Distribuidora" in concentrado.columns:
        filter_columns.append("Distribuidora")
    filtered = concentrado
    with st.sidebar:
        st.header("Filtros")
        for position, column in enumerate(filter_columns):
            options = sorted(filtered[column].dropna().astype(str).unique().tolist())
            choices = ["Todos", *options]
            key = f"filter_{profile}_{position}"
            if st.session_state.get(key, "Todos") not in choices:
                st.session_state[key] = "Todos"
            label = "País" if column in {"País", "Pais"} else column
            selected = st.selectbox(label, choices, key=key)
            if selected != "Todos":
                filtered = filtered[filtered[column].astype(str).eq(selected)]

    clients = filtered.get("Clientes totales", pd.Series(dtype=float)).sum()
    current = filtered.get("Clientes al corriente", pd.Series(dtype=float)).sum()
    portfolio = filtered.get("Cartera Total", pd.Series(dtype=float)).sum()
    clean_portfolio = filtered.get("Cartera sin atrasos", pd.Series(dtype=float)).sum()
    quality = clean_portfolio / portfolio if portfolio else 0

    if profile == "Vales":
        top_metrics = [
            ("Distribuidoras al corriente", metric_total(filtered, "Distribuidoras al corriente")),
            ("Clientes al corriente", current),
            ("Distribuidoras totales", len(filtered)),
            ("Cartera total", portfolio),
            ("Cartera al corriente", clean_portfolio),
            ("Calidad", quality),
        ]
        metric_columns = st.columns(len(top_metrics))
        for column, (label, value) in zip(metric_columns, top_metrics):
            if label.startswith("Cartera"):
                column.metric(label, f"${value:,.0f}")
            elif label == "Calidad":
                column.metric(label, f"{value:.1%}")
            else:
                column.metric(label, f"{value:,.0f}")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Localidades", f"{len(filtered):,}")
        c2.metric("Cartera al corriente", f"${clean_portfolio:,.0f}")
        c3.metric("Calidad", f"{quality:.1%}")

    if profile == "Vales":
        show_vales_cards(filtered)
    else:
        show_detail_cards(filtered)

    tab_names = ["Concentrado", "Resumen"]
    if profile == "LATAM":
        tab_names.append("Detalle por país")
    tab_names.append("Control")
    tabs = st.tabs(tab_names)
    tab1, tab2 = tabs[:2]
    country_tab = tabs[2] if profile == "LATAM" else None
    tab3 = tabs[-1]

    with tab1:
        st.dataframe(filtered, use_container_width=True, height=560, hide_index=True)
    with tab2:
        group = next(
            (c for c in ["País", "Pais", "Territorio", "Unidad de negocio", "Subdireccion", "Zona"] if c in filtered.columns),
            None,
        )
        if group:
            summary = filtered.groupby(group, dropna=False)[["Cartera Total", "Cartera sin atrasos"]].sum()
            summary["Calidad"] = summary["Cartera sin atrasos"].div(summary["Cartera Total"]).fillna(0)
            st.bar_chart(summary[["Cartera Total", "Cartera sin atrasos"]])
            st.dataframe(summary.reset_index(), use_container_width=True, hide_index=True)
    if country_tab is not None:
        with country_tab:
            country_column = next(
                (c for c in ["País", "Pais"] if c in filtered.columns),
                None,
            )
            if country_column and not filtered.empty:
                countries = sorted(filtered[country_column].dropna().astype(str).unique().tolist())
                if countries:
                    if st.session_state.get("latam_country_detail") not in countries:
                        st.session_state["latam_country_detail"] = countries[0]
                    selected_country = st.selectbox(
                        "País",
                        countries,
                        key="latam_country_detail",
                    )
                    selected_rows = filtered[
                        filtered[country_column].astype(str).eq(selected_country)
                    ]
                    show_detail_cards(selected_rows)
                    st.markdown("#### Comparativo de todos los países")
                    detail = country_detail(filtered, country_column)
                    st.dataframe(detail, use_container_width=True, hide_index=True)
                else:
                    st.info("No hay países disponibles con los filtros seleccionados.")
            else:
                st.info("No hay información por país disponible.")
    with tab3:
        missing_route = int(base["ruta"].isna().sum()) if "ruta" in base.columns else len(base)
        missing_locality = int(base["id_y_localidad"].isna().sum()) if "id_y_localidad" in base.columns else len(base)
        missing_structure = int(base["Sucursal"].isna().sum()) if "Sucursal" in base.columns else len(base)
        st.write({
            "filas_base": len(base), "sin_ruta": missing_route,
            "sin_localidad": missing_locality, "sin_estructura": missing_structure,
        })

    st.download_button(
        "Descargar Excel", excel_bytes(filtered), f"Concentrado_{profile}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        "Descargar CSV", filtered.to_csv(index=False).encode("utf-8-sig"), f"Concentrado_{profile}.csv",
        mime="text/csv",
    )


if __name__ == "__main__":
    main()
