# -*- coding: utf-8 -*-
"""Editor de Excel -> CRM
Convierte archivos .xlsx a las 26 columnas estandar del CRM con mapeo configurable.
"""

import os
import json
import zipfile
from io import BytesIO

import pandas as pd
import streamlit as st

# ── Constantes ──────────────────────────────────────────────────────────────────
CRM_COLUMNS = [
    "DOC_DNI_RUC",        "NOM_CLI",            "CARTERA",            "COD_CREDITO",
    "NOM_AGENCIA",        "DEUDA_CAP",          "DEUDA_TOTAL",        "TLF_CELULAR_CLIENTE",
    "DIR_CASA",           "DISTRITO",           "RANGO_DIAS_MORA",    "FEC_ULT_PAGO_ACTUAL",
    "NOM_CONYUGE",        "NOM_AVAL",           "TLF_CELULAR_AVAL",   "NOM_CONYUGE_AVAL",
    "DIR_CASA_AVAL",      "DISTRITO_AVAL",      "EXPEDIENTE",         "JUZGADO",
    "CONDICION",          "REFERENCIA",         "PROCESO_JUDICIAL",   "FEC_DEMANDA",
    "MONTO_DEMANDA",      "FEC_INGRESO_JUDICIAL",
]

NUMERIC_COLS  = {"DEUDA_CAP", "DEUDA_TOTAL", "MONTO_DEMANDA"}
PROFILES_FILE = "mapping_profiles.json"
NONE_LABEL    = "-- sin mapeo --"


# ── Utilidades de datos ─────────────────────────────────────────────────────────

def load_profiles():
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_profiles(profiles):
    with open(PROFILES_FILE, "w", encoding="utf-8") as f:
        json.dump(profiles, f, ensure_ascii=False, indent=2)


def auto_map(source_cols):
    lookup = {c.upper().strip(): c for c in source_cols}
    return {crm: lookup[crm] for crm in CRM_COLUMNS if crm in lookup}


def is_non_numeric(val):
    if pd.isna(val) or str(val).strip() in ("", "nan", "None"):
        return False
    try:
        float(str(val).replace(",", ".").replace(" ", ""))
        return False
    except ValueError:
        return True


def apply_mapping(df, mapping):
    n = len(df)
    data = {}
    for col in CRM_COLUMNS:
        src = mapping.get(col, "")
        data[col] = df[src].tolist() if src and src in df.columns else [""] * n
    return pd.DataFrame(data)


def to_excel_bytes(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="CRM")
    return buf.getvalue()


def get_alerts(df):
    alerts = []
    empty_mask = df["DOC_DNI_RUC"].isna() | df["DOC_DNI_RUC"].astype(str).str.strip().eq("")
    n_empty = int(empty_mask.sum())
    if n_empty:
        alerts.append(("warning", "**DOC_DNI_RUC**: {} registro(s) vacios".format(n_empty)))
    for col in NUMERIC_COLS:
        bad = df[col].apply(is_non_numeric)
        if bad.any():
            examples = df.loc[bad, col].astype(str).unique()[:3].tolist()
            alerts.append(("error", "**{}**: valores no numericos -> {}".format(col, examples)))
    return alerts


def merge_and_deduplicate(dfs, mapping):
    frames = []
    for fname, df in dfs.items():
        transformed = apply_mapping(df, mapping)
        transformed["_source_file"] = fname
        frames.append(transformed)

    combined = pd.concat(frames, ignore_index=True)
    n_total  = len(combined)

    combined["_fecha_ord"] = pd.to_datetime(
        combined["FEC_ULT_PAGO_ACTUAL"].astype(str).str.strip(),
        dayfirst=True, errors="coerce"
    )

    empty_mask = combined["COD_CREDITO"].astype(str).str.strip().eq("")
    df_sin     = combined[empty_mask].copy()
    df_con     = combined[~empty_mask].copy()

    dup_mask   = df_con.duplicated(subset=["COD_CREDITO"], keep=False)
    duplicados = df_con[dup_mask].copy()

    df_con   = df_con.sort_values("_fecha_ord", ascending=False, na_position="last")
    df_dedup = df_con.drop_duplicates(subset=["COD_CREDITO"], keep="first")

    result    = pd.concat([df_dedup, df_sin], ignore_index=True)
    n_removed = n_total - len(result)

    result     = result.drop(columns=["_fecha_ord", "_source_file"])
    duplicados = duplicados.drop(columns=["_fecha_ord"])

    return result, n_removed, duplicados


# ── App principal ───────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Editor CRM Excel",
        page_icon="=",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown(
        "<style>"
        ".block-container{padding-top:1.5rem}"
        "div[data-testid='stSelectbox'] label p{font-weight:600;font-size:.8rem}"
        "</style>",
        unsafe_allow_html=True,
    )

    st.title("Editor de Excel -> CRM")
    st.caption("Carga archivos .xlsx, mapea sus columnas a las 26 columnas del CRM y descarga el resultado.")

    # Estado de sesion
    for key, val in [("profiles", None), ("mapping", {}), ("dfs", {}), ("file_sheets", {})]:
        if key not in st.session_state:
            st.session_state[key] = val if val is not None else load_profiles()

    # ── Barra lateral ────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("Archivos Excel")
        uploaded = st.file_uploader(
            "Sube uno o mas archivos .xlsx",
            type=["xlsx"],
            accept_multiple_files=True,
        )

        if uploaded:
            current_names = {f.name for f in uploaded}
            st.session_state.dfs         = {k: v for k, v in st.session_state.dfs.items()         if k in current_names}
            st.session_state.file_sheets = {k: v for k, v in st.session_state.file_sheets.items() if k in current_names}

            for uf in uploaded:
                if uf.name not in st.session_state.dfs:
                    try:
                        xl     = pd.ExcelFile(uf)
                        sheets = xl.sheet_names
                        st.session_state.file_sheets[uf.name] = sheets
                        df_read = pd.read_excel(uf, sheet_name=sheets[0], dtype=str).fillna("")
                        st.session_state.dfs[uf.name] = df_read
                        st.success("{} - {:,} filas ({} hojas)".format(uf.name, len(df_read), len(sheets)))
                    except Exception as e:
                        st.error("{}: {}".format(uf.name, e))
        else:
            st.session_state.dfs         = {}
            st.session_state.file_sheets = {}

        st.divider()

        # Perfiles de mapeo
        st.header("Perfiles de Mapeo")
        profiles = st.session_state.profiles
        p_names  = list(profiles.keys())

        if p_names:
            ca, cb = st.columns([3, 1])
            sel_load = ca.selectbox("Perfiles guardados", [""] + p_names,
                                    label_visibility="collapsed", key="sel_load")
            cb.write(""); cb.write("")
            if cb.button("Cargar", key="btn_load") and sel_load:
                st.session_state.mapping = profiles[sel_load].copy()
                st.toast("Perfil '{}' cargado".format(sel_load))
                st.rerun()

        cc, cd = st.columns([3, 1])
        new_name = cc.text_input("Nombre del perfil", placeholder="Ej: Banco X",
                                  label_visibility="collapsed", key="inp_save")
        cd.write(""); cd.write("")
        if cd.button("Guardar", key="btn_save") and new_name.strip():
            profiles[new_name.strip()] = st.session_state.mapping.copy()
            save_profiles(profiles)
            st.toast("Perfil '{}' guardado".format(new_name.strip()))
            st.rerun()

        if p_names:
            ce, cf = st.columns([3, 1])
            sel_del = ce.selectbox("Eliminar perfil", [""] + p_names,
                                    label_visibility="collapsed", key="sel_del")
            cf.write(""); cf.write("")
            if cf.button("Eliminar", key="btn_del") and sel_del:
                del profiles[sel_del]
                save_profiles(profiles)
                st.toast("Perfil '{}' eliminado".format(sel_del))
                st.rerun()

    # ── Contenido principal ──────────────────────────────────────────────────────
    if not st.session_state.dfs:
        st.info("Sube uno o mas archivos .xlsx desde el panel izquierdo para comenzar.")
        with st.expander("Ver las 26 columnas destino del CRM"):
            st.code("\n".join("{:2}. {}".format(i + 1, c) for i, c in enumerate(CRM_COLUMNS)))
        return

    file_names = sorted(st.session_state.dfs.keys())
    if len(file_names) > 1:
        active_file = st.selectbox("Archivo activo:", file_names)
    else:
        active_file = file_names[0]
        st.info("Archivo cargado: **{}**".format(active_file))

    # Selector de hoja
    sheets = st.session_state.file_sheets.get(active_file, [])
    if len(sheets) > 1:
        sel_sheet = st.selectbox("Hoja:", sheets, key="active_sheet")
        cache_key = "{}_{}".format(active_file, sel_sheet)
        if cache_key not in st.session_state:
            raw_bytes = None
            for uf in (uploaded or []):
                if uf.name == active_file:
                    raw_bytes = uf
                    break
            if raw_bytes is not None:
                st.session_state[cache_key] = pd.read_excel(raw_bytes, sheet_name=sel_sheet, dtype=str).fillna("")
            else:
                st.session_state[cache_key] = st.session_state.dfs[active_file]
        active_df = st.session_state[cache_key]
    else:
        active_df = st.session_state.dfs[active_file]

    src_cols = list(active_df.columns)
    options  = [NONE_LABEL] + src_cols

    if not st.session_state.mapping:
        st.session_state.mapping = auto_map(src_cols)

    # ── Tabs ─────────────────────────────────────────────────────────────────────
    tab_map, tab_prev, tab_dl = st.tabs([
        "Mapeo de Columnas",
        "Vista Previa",
        "Descargar",
    ])

    # ── TAB 1: Mapeo ──────────────────────────────────────────────────────────────
    with tab_map:
        h1, h2, h3 = st.columns([4, 1, 1])
        h1.subheader("Columna origen -> Columna CRM destino")
        if h2.button("Auto-mapear", use_container_width=True):
            st.session_state.mapping = auto_map(src_cols)
            st.rerun()
        if h3.button("Limpiar todo", use_container_width=True):
            st.session_state.mapping = {}
            st.rerun()

        st.caption("{} - {} columnas origen - {:,} filas".format(active_file, len(src_cols), len(active_df)))

        with st.expander("Columnas disponibles en el archivo origen"):
            cols4 = st.columns(4)
            for i, c in enumerate(src_cols):
                cols4[i % 4].code(c, language=None)

        st.divider()
        new_mapping = {}
        grid = st.columns(3)
        for idx, crm_col in enumerate(CRM_COLUMNS):
            with grid[idx % 3]:
                cur     = st.session_state.mapping.get(crm_col, "")
                cur_opt = cur if (cur and cur in options) else NONE_LABEL
                sel     = st.selectbox(crm_col, options, index=options.index(cur_opt),
                                       key="map_{}".format(crm_col))
                new_mapping[crm_col] = "" if sel == NONE_LABEL else sel

        st.session_state.mapping = new_mapping
        n_mapped = sum(1 for v in new_mapping.values() if v)
        pct      = n_mapped / len(CRM_COLUMNS)
        st.divider()
        mc1, mc2 = st.columns([1, 3])
        mc1.metric("Columnas mapeadas", "{} / {}".format(n_mapped, len(CRM_COLUMNS)))
        mc2.progress(pct, text="{:.0f}% completado".format(pct * 100))

    # ── TAB 2: Vista previa ───────────────────────────────────────────────────────
    with tab_prev:
        st.subheader("Vista Previa -- primeros 20 registros transformados")
        preview = apply_mapping(active_df.head(20), st.session_state.mapping)

        alerts = get_alerts(preview)
        if alerts:
            for lvl, msg in alerts:
                icon = "AVISO" if lvl == "warning" else "ERROR"
                (st.warning if lvl == "warning" else st.error)("{} {}".format(icon, msg))
        else:
            st.success("Sin alertas en los primeros 20 registros")

        def _hl(val):
            return "background-color:#ffe4e4;color:#cc0000;" if str(val).strip() == "" else ""

        st.dataframe(
            preview.style.map(_hl, subset=["DOC_DNI_RUC"]),
            use_container_width=True,
            height=460,
        )

    # ── TAB 3: Descarga ───────────────────────────────────────────────────────────
    with tab_dl:
        st.subheader("Descargar Excel con las 26 columnas CRM")

        full        = apply_mapping(active_df, st.session_state.mapping)
        n_mapped_dl = sum(1 for v in st.session_state.mapping.values() if v)
        n_empty_dni = int((full["DOC_DNI_RUC"].isna() |
                           full["DOC_DNI_RUC"].astype(str).str.strip().eq("")).sum())

        d1, d2, d3 = st.columns(3)
        d1.metric("Total filas exportadas", "{:,}".format(len(full)))
        d2.metric("Columnas mapeadas",      "{} / 26".format(n_mapped_dl))
        d3.metric("DOC_DNI_RUC vacios",     n_empty_dni)

        full_alerts = get_alerts(full)
        if full_alerts:
            with st.expander("Alertas en el dataset completo", expanded=True):
                for lvl, msg in full_alerts:
                    (st.warning if lvl == "warning" else st.error)(msg)

        base = os.path.splitext(active_file)[0]
        st.download_button(
            label="Descargar Excel CRM",
            data=to_excel_bytes(full),
            file_name="{}_CRM.xlsx".format(base),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        # Combinar y deduplicar (solo con mas de 1 archivo)
        if len(st.session_state.dfs) > 1:
            st.divider()
            st.subheader("Combinar {} archivos con deduplicacion".format(len(st.session_state.dfs)))
            st.caption(
                "Si el mismo COD_CREDITO aparece en mas de un archivo, "
                "prevalece el registro con FEC_ULT_PAGO_ACTUAL mas reciente."
            )

            merged, n_removed, duplicados_df = merge_and_deduplicate(
                st.session_state.dfs, st.session_state.mapping
            )

            m1, m2, m3 = st.columns(3)
            m1.metric("Registros totales (antes)",
                      "{:,}".format(sum(len(d) for d in st.session_state.dfs.values())))
            m2.metric("Duplicados eliminados", "{:,}".format(n_removed))
            m3.metric("Registros finales",     "{:,}".format(len(merged)))

            if n_removed:
                with st.expander("Ver {} registro(s) duplicado(s) reemplazados".format(n_removed)):
                    show_cols = ["_source_file", "COD_CREDITO", "DOC_DNI_RUC",
                                 "NOM_CLI", "FEC_ULT_PAGO_ACTUAL", "DEUDA_TOTAL"]
                    show_cols = [c for c in show_cols if c in duplicados_df.columns]
                    st.dataframe(
                        duplicados_df[show_cols]
                        .sort_values(["COD_CREDITO", "FEC_ULT_PAGO_ACTUAL"])
                        .reset_index(drop=True),
                        use_container_width=True,
                        height=300,
                    )
            else:
                st.success("No se encontraron COD_CREDITO duplicados entre los archivos.")

            merged_alerts = get_alerts(merged)
            for lvl, msg in merged_alerts:
                (st.warning if lvl == "warning" else st.error)(msg)

            st.download_button(
                label="Descargar Excel Combinado (deduplicado)",
                data=to_excel_bytes(merged),
                file_name="combinado_CRM.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            st.divider()
            st.subheader("Exportar archivos por separado ({})".format(len(st.session_state.dfs)))
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, fdf in st.session_state.dfs.items():
                    transformed = apply_mapping(fdf, st.session_state.mapping)
                    zf.writestr(
                        "{}_CRM.xlsx".format(os.path.splitext(fname)[0]),
                        to_excel_bytes(transformed),
                    )
            st.download_button(
                label="Descargar todos por separado como ZIP",
                data=zip_buf.getvalue(),
                file_name="exportacion_CRM.zip",
                mime="application/zip",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
