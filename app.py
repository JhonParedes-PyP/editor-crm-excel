#!/usr/bin/env python3
"""
Editor de Excel → CRM
Convierte archivos .xlsx a las 26 columnas estándar del CRM con mapeo configurable.
"""

import os
import json
import zipfile
from io import BytesIO

import pandas as pd
import streamlit as st

# ─── Columnas destino del CRM (orden fijo) ────────────────────────────────────
CRM_COLUMNS = [
    "DOC_DNI_RUC",       "NOM_CLI",           "CARTERA",           "COD_CREDITO",
    "NOM_AGENCIA",       "DEUDA_CAP",         "DEUDA_TOTAL",       "TLF_CELULAR_CLIENTE",
    "DIR_CASA",          "DISTRITO",          "RANGO_DIAS_MORA",   "FEC_ULT_PAGO_ACTUAL",
    "NOM_CONYUGE",       "NOM_AVAL",          "TLF_CELULAR_AVAL",  "NOM_CONYUGE_AVAL",
    "DIR_CASA_AVAL",     "DISTRITO_AVAL",     "EXPEDIENTE",        "JUZGADO",
    "CONDICION",         "REFERENCIA",        "PROCESO_JUDICIAL",  "FEC_DEMANDA",
    "MONTO_DEMANDA",     "FEC_INGRESO_JUDICIAL",
]

NUMERIC_COLS = {"DEUDA_CAP", "DEUDA_TOTAL", "MONTO_DEMANDA"}
PROFILES_FILE = "mapping_profiles.json"
NONE_LABEL    = "— sin mapeo —"


# ─── Utilidades de datos ───────────────────────────────────────────────────────

def load_profiles() -> dict:
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_profiles(profiles: dict) -> None:
    with open(PROFILES_FILE, "w", encoding="utf-8") as f:
        json.dump(profiles, f, ensure_ascii=False, indent=2)


def auto_map(source_cols: list) -> dict:
    """Mapeo automático por nombre idéntico (insensible a mayúsculas/espacios)."""
    lookup = {c.upper().strip(): c for c in source_cols}
    return {crm: lookup[crm] for crm in CRM_COLUMNS if crm in lookup}


def is_non_numeric(val) -> bool:
    """True si el valor no puede convertirse a número."""
    if pd.isna(val) or str(val).strip() in ("", "nan", "None"):
        return False
    try:
        float(str(val).replace(",", ".").replace(" ", ""))
        return False
    except ValueError:
        return True


def apply_mapping(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """Retorna un DataFrame con exactamente las 26 columnas CRM en orden fijo."""
    n = len(df)
    data = {}
    for col in CRM_COLUMNS:
        src = mapping.get(col, "")
        data[col] = df[src].tolist() if src and src in df.columns else [""] * n
    return pd.DataFrame(data)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="CRM")
    return buf.getvalue()


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
            "Sube uno o más archivos .xlsx",
            type=["xlsx"],
            accept_multiple_files=True,
            help="Se leerá la primera hoja de cada archivo",
        )

        if uploaded:
            current_names = {f.name for f in uploaded}
            # Quitar archivos que el usuario removió del uploader
            st.session_state.dfs = {k: v for k, v in st.session_state.dfs.items()
                                     if k in current_names}
            # Leer archivos nuevos
            for uf in uploaded:
                if uf.name not in st.session_state.dfs:
                    with st.spinner(f"Leyendo {uf.name}…"):
                        try:
                            df_read = pd.read_excel(uf, sheet_name=0, dtype=str).fillna("")
                            st.session_state.dfs[uf.name] = df_read
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
    if len(file_names) > 1:
        active_file = st.selectbox("📄 Archivo activo:", file_names)
    else:
        active_file = file_names[0]
        st.info(f"📄 Archivo cargado: **{active_file}**")

    active_df = st.session_state.dfs[active_file]
    src_cols  = list(active_df.columns)
    options   = [NONE_LABEL] + src_cols

    # Auto-mapeo inicial si el mapeo está vacío
    if not st.session_state.mapping:
        st.session_state.mapping = auto_map(src_cols)

    # ── Tabs ─────────────────────────────────────────────────────────────────────
    tab_map, tab_prev, tab_dl = st.tabs([
        "🔗  Mapeo de Columnas",
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
        st.subheader("Vista Previa — primeros 20 registros transformados")
        preview = apply_mapping(active_df.head(20), st.session_state.mapping)

        alerts = get_alerts(preview)
        if alerts:
            for lvl, msg in alerts:
                icon = "⚠️" if lvl == "warning" else "🚨"
                (st.warning if lvl == "warning" else st.error)(f"{icon} {msg}")
        else:
            st.success("✅ Sin alertas en los primeros 20 registros")

        # Resaltar celdas vacías de DOC_DNI_RUC en rojo
        def _hl(val):
            return "background-color:#ffe4e4;color:#cc0000;" if str(val).strip() == "" else ""

        st.dataframe(
            preview.style.map(_hl, subset=["DOC_DNI_RUC"]),
            use_container_width=True,
            height=460,
        )

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

        # ── Exportar todos los archivos como ZIP ──────────────────────────────
        if len(st.session_state.dfs) > 1:
            st.divider()
            st.subheader(f"Exportar todos los archivos ({len(st.session_state.dfs)})")
            st.caption("Se aplica el mismo mapeo actual a todos los archivos cargados.")

            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, fdf in st.session_state.dfs.items():
                    transformed = apply_mapping(fdf, st.session_state.mapping)
                    zf.writestr(
                        f"{os.path.splitext(fname)[0]}_CRM.xlsx",
                        to_excel_bytes(transformed),
                    )

            st.download_button(
                label="📦 Descargar todos como ZIP",
                data=zip_buf.getvalue(),
                file_name="exportacion_CRM.zip",
                mime="application/zip",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
