
Convierte archivos .xlsx a las 26 columnas estándar del CRM con mapeo configurable.
"""

import os
import json
import zipfile
from io import BytesIO

import pandas as pd
import streamlit as st

from config import CRM_COLUMNS, NUMERIC_COLS, NONE_LABEL, PROFILES_FILE
from utils import load_profiles, save_profiles, auto_map, is_non_numeric, apply_mapping, get_alerts




# ─── Utilidades restantes ───────────────────────────────────────────────────────

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="CRM")
    return buf.getvalue()


def merge_and_deduplicate(dfs: dict, mapping: dict) -> tuple:
    """
    Aplica el mapeo a TODOS los archivos, los concatena y deduplica por COD_CREDITO.
    Cuando hay duplicados, prevalece el registro con FEC_ULT_PAGO_ACTUAL más reciente.
    Retorna (df_resultado, n_duplicados_eliminados, detalle_duplicados).
    """
    frames = []
    for fname, df in dfs.items():
        transformed = apply_mapping(df, mapping)
        transformed["_source_file"] = fname
        frames.append(transformed)

    combined = pd.concat(frames, ignore_index=True)
    n_total = len(combined)

    # Parsear FEC_ULT_PAGO_ACTUAL con varios formatos posibles
    combined["_fecha_ord"] = pd.to_datetime(
        combined["FEC_ULT_PAGO_ACTUAL"].astype(str).str.strip(),
        dayfirst=True, errors="coerce"
    )

    # Registros sin COD_CREDITO no se deduplicarán (se conservan todos)
    empty_mask  = combined["COD_CREDITO"].astype(str).str.strip().eq("")
    df_sin_cod  = combined[empty_mask].copy()
    df_con_cod  = combined[~empty_mask].copy()

    # Detectar duplicados antes de eliminar (para el reporte)
    dup_mask   = df_con_cod.duplicated(subset=["COD_CREDITO"], keep=False)
    duplicados = df_con_cod[dup_mask].copy()

    # Ordenar: fecha más reciente primero; NaT al final
    df_con_cod = df_con_cod.sort_values("_fecha_ord", ascending=False, na_position="last")
    df_dedup   = df_con_cod.drop_duplicates(subset=["COD_CREDITO"], keep="first")

    result = pd.concat([df_dedup, df_sin_cod], ignore_index=True)
    n_removed = n_total - len(result)

    # Limpiar columnas auxiliares
    result     = result.drop(columns=["_fecha_ord", "_source_file"])
    duplicados = duplicados.drop(columns=["_fecha_ord"])

    return result, n_removed, duplicados


def get_alerts(df: pd.DataFrame) -> list:
    """Valida DOC_DNI_RUC vacío y montos no numéricos. Retorna lista de (nivel, mensaje)."""
    alerts = []

    empty_mask = df["DOC_DNI_RUC"].isna() | df["DOC_DNI_RUC"].astype(str).str.strip().eq("")
    n_empty = int(empty_mask.sum())
    if n_empty:
        alerts.append(("warning", f"**DOC_DNI_RUC**: {n_empty} registro(s) vacíos"))

    for col in NUMERIC_COLS:
        bad = df[col].apply(is_non_numeric)
        if bad.any():
            examples = df.loc[bad, col].astype(str).unique()[:3].tolist()
            alerts.append(("error", f"**{col}**: valores no numéricos detectados → {examples}"))

    return alerts


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="CRM")
    return buf.getvalue()


def merge_and_deduplicate(dfs: dict, mapping: dict) -> tuple:
    """
    Aplica el mapeo a TODOS los archivos, los concatena y deduplica por COD_CREDITO.
    Cuando hay duplicados, prevalece el registro con FEC_ULT_PAGO_ACTUAL más reciente.
    Retorna (df_resultado, n_duplicados_eliminados, detalle_duplicados).
    """
    frames = []
    for fname, df in dfs.items():
        transformed = apply_mapping(df, mapping)
        transformed["_source_file"] = fname
        frames.append(transformed)

    combined = pd.concat(frames, ignore_index=True)
    n_total = len(combined)

    # Parsear FEC_ULT_PAGO_ACTUAL con varios formatos posibles
    combined["_fecha_ord"] = pd.to_datetime(
        combined["FEC_ULT_PAGO_ACTUAL"].astype(str).str.strip(),
        dayfirst=True, errors="coerce"
    )

    # Registros sin COD_CREDITO no se deduplicarán (se conservan todos)
    empty_mask  = combined["COD_CREDITO"].astype(str).str.strip().eq("")
    df_sin_cod  = combined[empty_mask].copy()
    df_con_cod  = combined[~empty_mask].copy()

    # Detectar duplicados antes de eliminar (para el reporte)
    dup_mask   = df_con_cod.duplicated(subset=["COD_CREDITO"], keep=False)
    duplicados = df_con_cod[dup_mask].copy()

    # Ordenar: fecha más reciente primero; NaT al final
    df_con_cod = df_con_cod.sort_values("_fecha_ord", ascending=False, na_position="last")
    df_dedup   = df_con_cod.drop_duplicates(subset=["COD_CREDITO"], keep="first")

    result = pd.concat([df_dedup, df_sin_cod], ignore_index=True)
    n_removed = n_total - len(result)

    # Limpiar columnas auxiliares
    result     = result.drop(columns=["_fecha_ord", "_source_file"])
    duplicados = duplicados.drop(columns=["_fecha_ord"])

    return result, n_removed, duplicados


def get_alerts(df: pd.DataFrame) -> list:
    """Valida DOC_DNI_RUC vacío y montos no numéricos. Retorna lista de (nivel, mensaje)."""
    alerts = []

    empty_mask = df["DOC_DNI_RUC"].isna() | df["DOC_DNI_RUC"].astype(str).str.strip().eq("")
    n_empty = int(empty_mask.sum())
    if n_empty:
        alerts.append(("warning", f"**DOC_DNI_RUC**: {n_empty} registro(s) vacíos"))

    for col in NUMERIC_COLS:
        bad = df[col].apply(is_non_numeric)
        if bad.any():
            examples = df.loc[bad, col].astype(str).unique()[:3].tolist()
            alerts.append(("error", f"**{col}**: valores no numéricos detectados → {examples}"))

    return alerts


# ─── App Streamlit ─────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Editor CRM Excel",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown("""
        <style>
        .block-container { padding-top: 1.5rem; }
        div[data-testid="stSelectbox"] label p { font-weight: 600; font-size: 0.80rem; }
        </style>
    """, unsafe_allow_html=True)

    st.title("📊 Editor de Excel → CRM")
    st.caption("Carga archivos **.xlsx**, mapea sus columnas a las **26 columnas del CRM** y descarga el resultado.")

    # ── Estado de sesión ────────────────────────────────────────────────────────
    if "profiles" not in st.session_state:
        st.session_state.profiles = load_profiles()
    if "mapping" not in st.session_state:
        st.session_state.mapping = {}
    if "dfs" not in st.session_state:
        st.session_state.dfs = {}

    # ── Barra lateral ────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("📁 Archivos Excel")

        uploaded = st.file_uploader(
            "Sube uno o más archivos .xlsx (multi-hoja soportado)",
            type=["xlsx"],
            accept_multiple_files=True,
        )

        if "file_sheets" not in st.session_state:
            st.session_state.file_sheets = {}

        if uploaded:
            current_names = {f.name for f in uploaded}
            # Limpiar DFs y sheets removidos
            st.session_state.dfs = {k: v for k, v in st.session_state.dfs.items() if k in current_names}
            st.session_state.file_sheets = {k: v for k, v in st.session_state.file_sheets.items() if k in current_names}
            
            for uf in uploaded:
                if uf.name not in st.session_state.dfs:
                    try:
                        xl = pd.ExcelFile(uf)
                        sheets = xl.sheet_names
                        st.session_state.file_sheets[uf.name] = sheets
                        
                        # Por defecto primera hoja
                        df_read = pd.read_excel(uf, sheet_name=sheets[0], dtype=str).fillna("")
                        st.session_state.dfs[uf.name] = df_read
                        st.success(f"✅ **{uf.name}** — {len(df_read):,} filas ({len(sheets)} hojas)")
                    except Exception as e:
                        st.error(f"❌ {uf.name}: {e}")


            for name, df_s in sorted(st.session_state.dfs.items()):
                st.success(f"✅ **{name}** — {len(df_s):,} filas")
        else:
            st.session_state.dfs = {}

        st.divider()

        # ── Perfiles de mapeo ─────────────────────────────────────────────────
        st.header("💾 Perfiles de Mapeo")
        profiles = st.session_state.profiles
        p_names = list(profiles.keys())

        # Cargar perfil
        if p_names:
            ca, cb = st.columns([3, 1])
            sel_load = ca.selectbox(
                "Perfiles guardados", [""] + p_names,
                label_visibility="collapsed", key="sel_load",
            )
            cb.write(""); cb.write("")
            if cb.button("↩️", key="btn_load", help="Cargar perfil seleccionado") and sel_load:
                st.session_state.mapping = profiles[sel_load].copy()
                st.toast(f"Perfil «{sel_load}» cargado ✅")
                st.rerun()

        # Guardar perfil
        cc, cd = st.columns([3, 1])
        new_name = cc.text_input(
            "Nombre del perfil", placeholder="Ej: Banco X",
            label_visibility="collapsed", key="inp_save",
        )
        cd.write(""); cd.write("")
        if cd.button("💾", key="btn_save", help="Guardar mapeo actual") and new_name.strip():
            profiles[new_name.strip()] = st.session_state.mapping.copy()
            save_profiles(profiles)
            st.toast(f"Perfil «{new_name.strip()}» guardado 💾")
            st.rerun()

        # Eliminar perfil
        if p_names:
            ce, cf = st.columns([3, 1])
            sel_del = ce.selectbox(
                "Eliminar perfil", [""] + p_names,
                label_visibility="collapsed", key="sel_del",
            )
            cf.write(""); cf.write("")
            if cf.button("🗑️", key="btn_del", help="Eliminar perfil") and sel_del:
                del profiles[sel_del]
                save_profiles(profiles)
                st.toast(f"Perfil «{sel_del}» eliminado 🗑️")
                st.rerun()

    # ── Contenido principal ──────────────────────────────────────────────────────
    if not st.session_state.dfs:
        st.info("👆 Sube uno o más archivos **.xlsx** desde el panel izquierdo para comenzar.")
        with st.expander("📋 Ver las 26 columnas destino del CRM"):
            st.code("\n".join(f"{i+1:2}. {c}" for i, c in enumerate(CRM_COLUMNS)))
        return

    # Selector de archivo activo
    file_names = sorted(st.session_state.dfs.keys())
    if file_names:
        active_file = st.selectbox("📄 Archivo activo:", file_names, key="active_file")
        sheets = st.session_state.file_sheets.get(active_file, [])
        if sheets:
            sheet_name = st.selectbox("📄 Hoja:", sheets, key="active_sheet")
            active_df_key = (active_file, sheet_name)
            
            if active_df_key not in st.session_state:
                df_read = pd.read_excel(st.session_state.file_sheets[active_file], sheet_name=sheet_name, dtype=str).fillna("")
                st.session_state[active_df_key] = df_read
            
            active_df = st.session_state[active_df_key]
            st.info(f"📄 **{active_file}** — Hoja: **{sheet_name}** — {len(active_df):,} filas")
        else:
            active_df = st.session_state.dfs[active_file]
    else:
        active_df = pd.DataFrame()

    src_cols  = list(active_df.columns)
    options   = [NONE_LABEL] + src_cols

    # Auto-mapeo inicial si el mapeo está vacío
    if not st.session_state.mapping:
        st.session_state.mapping = auto_map(src_cols)

    # ── Tabs ─────────────────────────────────────────────────────────────────────
    tab_map, tab_stats, tab_prev, tab_dl = st.tabs([
        "🔗  Mapeo de Columnas",
        "📈 Stats & Filtros",
        "👁️  Vista Previa",
        "⬇️  Descargar",
    ])

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1 · Mapeo de columnas
    # ══════════════════════════════════════════════════════════════════════════
    with tab_map:
        h1, h2, h3 = st.columns([4, 1, 1])
        h1.subheader("Columna origen  →  Columna CRM destino")
        if h2.button("🔄 Auto-mapear", use_container_width=True,
                     help="Detecta columnas con el mismo nombre que las del CRM"):
            st.session_state.mapping = auto_map(src_cols)
            st.rerun()
        if h3.button("🧹 Limpiar todo", use_container_width=True,
                     help="Quitar todos los mapeos"):
            st.session_state.mapping = {}
            st.rerun()

        st.caption(
            f"Archivo: **{active_file}** · {len(src_cols)} columnas origen · "
            f"{len(active_df):,} filas  —  "
            "Para cada campo del CRM, elige la columna del archivo origen que le corresponde."
        )

        with st.expander("📋 Columnas disponibles en el archivo origen"):
            disp_cols = st.columns(4)
            for i, c in enumerate(src_cols):
                disp_cols[i % 4].code(c, language=None)

        st.divider()

        new_mapping = {}
        grid = st.columns(3)
        for idx, crm_col in enumerate(CRM_COLUMNS):
            with grid[idx % 3]:
                cur = st.session_state.mapping.get(crm_col, "")
                cur_opt = cur if (cur and cur in options) else NONE_LABEL
                sel = st.selectbox(
                    crm_col, options,
                    index=options.index(cur_opt),
                    key=f"map_{crm_col}",
                )
                new_mapping[crm_col] = "" if sel == NONE_LABEL else sel

        st.session_state.mapping = new_mapping

        n_mapped = sum(1 for v in new_mapping.values() if v)
        pct = n_mapped / len(CRM_COLUMNS)
        st.divider()
        mc1, mc2 = st.columns([1, 3])
        mc1.metric("Columnas mapeadas", f"{n_mapped} / {len(CRM_COLUMNS)}")
        mc2.progress(pct, text=f"{pct*100:.0f}% completado")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2 · Vista previa
    # ══════════════════════════════════════════════════════════════════════════
    with tab_prev:
        st.subheader("👁️  Vista Previa y Edición")
        
        if len(active_df) > 0:
            preview_full = apply_mapping(active_df.head(100), st.session_state.mapping)  # Más filas para edición
            alerts = get_alerts(preview_full)
            if alerts:
                for lvl, msg in alerts:
                    icon = "⚠️" if lvl == "warning" else "🚨"
                    (st.warning if lvl == "warning" else st.error)(f"{icon} {msg}")
            else:
                st.success("✅ Sin alertas en preview")

            # Edición inline con streamlit-aggrid
            try:
                from st_aggrid import GridOptionsBuilder, AgGrid, GridUpdateMode
                gb = GridOptionsBuilder.from_dataframe(preview_full)
                gb.configure_default_column(minWidth=120, autoHeight=True)
                gb.configure_column("DOC_DNI_RUC", cellStyle={"backgroundColor": "rgba(255,228,228,0.5)"} if any(preview_full["DOC_DNI_RUC"].str.strip()=="") else None)
                gridOptions = gb.build()
                
                grid_response = AgGrid(
                    preview_full,
                    gridOptions=gridOptions,
                    height=400,
                    update_mode=GridUpdateMode.SELECTION_CHANGED,
                    fit_columns_on_grid_load=True
                )
                
                if grid_response['selected_rows']:
                    st.session_state.edited_preview = pd.DataFrame(grid_response['selected_rows'])
                    st.success("✅ Cambios editados guardados en session")
            except ImportError:
                # Fallback si aggrid no instalado
                def _hl(val):
                    return "background-color:#ffe4e4;color:#cc0000;" if str(val).strip() == "" else ""
                st.data_editor(
                    preview_full.head(20).style.map(_hl, subset=["DOC_DNI_RUC"]),
                    use_container_width=True,
                    num_rows="dynamic",
                    key="edited_preview_fallback"
                )
        else:
            st.warning("Carga un archivo para ver preview")


    # ══════════════════════════════════════════════════════════════════════════
    # TAB 3 · Descarga
    # ══════════════════════════════════════════════════════════════════════════
    with tab_dl:
        st.subheader("Descargar Excel con las 26 columnas CRM")

        full = apply_mapping(active_df, st.session_state.mapping)
        n_mapped_dl  = sum(1 for v in st.session_state.mapping.values() if v)
        n_empty_dni  = int(
            (full["DOC_DNI_RUC"].isna() |
             full["DOC_DNI_RUC"].astype(str).str.strip().eq("")).sum()
        )

        d1, d2, d3 = st.columns(3)
        d1.metric("Total filas exportadas", f"{len(full):,}")
        d2.metric("Columnas mapeadas",      f"{n_mapped_dl} / 26")
        d3.metric("DOC_DNI_RUC vacíos",     n_empty_dni)

        full_alerts = get_alerts(full)
        if full_alerts:
            with st.expander("⚠️ Alertas en el dataset completo", expanded=True):
                for lvl, msg in full_alerts:
                    icon = "⚠️" if lvl == "warning" else "🚨"
                    (st.warning if lvl == "warning" else st.error)(f"{icon} {msg}")

        base = os.path.splitext(active_file)[0]
        st.download_button(
            label="⬇️ Descargar Excel CRM",
            data=to_excel_bytes(full),
            file_name=f"{base}_CRM.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        # ── Combinar y deduplicar (solo si hay más de 1 archivo) ─────────────
        if len(st.session_state.dfs) > 1:
            st.divider()
            st.subheader(f"🔀 Combinar {len(st.session_state.dfs)} archivos con deduplicación")
            st.caption(
                "Une todos los archivos en uno solo. Si el mismo **COD_CREDITO** (número de cuenta) aparece "
                "en más de un archivo, prevalece el registro con **FEC_ULT_PAGO_ACTUAL** más reciente."
            )

            merged, n_removed, duplicados_df = merge_and_deduplicate(
                st.session_state.dfs, st.session_state.mapping
            )

            m1, m2, m3 = st.columns(3)
            m1.metric("Registros totales (antes)",
                      f"{sum(len(d) for d in st.session_state.dfs.values()):,}")
            m2.metric("Duplicados eliminados",     f"{n_removed:,}",
                      delta=f"-{n_removed}" if n_removed else None,
                      delta_color="inverse" if n_removed else "off")
            m3.metric("Registros finales",         f"{len(merged):,}")

            if n_removed:
                with st.expander(
                    f"🔍 Ver los {n_removed} registro(s) duplicado(s) que fueron reemplazados",
                    expanded=False,
                ):
                    st.caption(
                        "Se muestra el grupo completo de cada número de cuenta duplicado. "
                        "La fila con la fecha más reciente fue la que se conservó."
                    )
                    show_cols = ["_source_file", "COD_CREDITO", "DOC_DNI_RUC", "NOM_CLI",
                                 "FEC_ULT_PAGO_ACTUAL", "DEUDA_TOTAL"]
                    show_cols = [c for c in show_cols if c in duplicados_df.columns]
                    st.dataframe(
                        duplicados_df[show_cols]
                        .sort_values(["COD_CREDITO", "FEC_ULT_PAGO_ACTUAL"])
                        .reset_index(drop=True),
                        use_container_width=True,
                        height=300,
                    )
            else:
                st.success("✅ No se encontraron COD_CREDITO duplicados entre los archivos.")

            merged_alerts = get_alerts(merged)
            if merged_alerts:
                for lvl, msg in merged_alerts:
                    icon = "⚠️" if lvl == "warning" else "🚨"
                    (st.warning if lvl == "warning" else st.error)(f"{icon} {msg}")

            st.download_button(
                label="🔀 Descargar Excel Combinado (deduplicado)",
                data=to_excel_bytes(merged),
                file_name="combinado_CRM.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            st.divider()
            st.subheader(f"📦 Exportar archivos por separado ({len(st.session_state.dfs)})")
            st.caption("Cada archivo se convierte individualmente con el mapeo actual.")

            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, fdf in st.session_state.dfs.items():
                    transformed = apply_mapping(fdf, st.session_state.mapping)
                    zf.writestr(
                        f"{os.path.splitext(fname)[0]}_CRM.xlsx",
                        to_excel_bytes(transformed),
                    )

            st.download_button(
                label="📦 Descargar todos por separado como ZIP",
                data=zip_buf.getvalue(),
                file_name="exportacion_CRM.zip",
                mime="application/zip",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
