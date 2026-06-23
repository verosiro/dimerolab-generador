"""Streamlit UI del generador de Planilla cobro.

Subí el .xlsm del día con la hoja `ProtocolosDigitales` ya pegada de
AppSheet. La app genera la `Planilla cobro` precargada con una fila por
estudio y te devuelve el .xlsm modificado.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from generador import correr


DATA_DIR = Path(__file__).parent / "data"


st.set_page_config(
    page_title="Generador Planilla cobro · DimeroLab",
    page_icon="🧪",
    layout="wide",
)


def check_pin() -> bool:
    """PIN opcional vía st.secrets. Si no hay PIN configurado, pasa directo."""
    try:
        pin = st.secrets["PIN"]
    except (KeyError, FileNotFoundError, Exception):
        pin = None
    if not pin:
        return True
    if st.session_state.get("pin_ok"):
        return True
    intento = st.text_input("PIN", type="password")
    if intento and intento == pin:
        st.session_state["pin_ok"] = True
        st.rerun()
    if intento:
        st.error("PIN incorrecto")
    return False


if not check_pin():
    st.stop()


st.title("🧪 Generador de Planilla cobro")
st.caption(
    "Subí el .xlsm del día con `ProtocolosDigitales` cargado. "
    "La app llena la `Planilla cobro` (una fila por estudio) y te devuelve el archivo."
)

with st.expander("Cómo funciona", expanded=False):
    st.markdown(
        """
        - Lee la hoja `ProtocolosDigitales` (la que pegás de AppSheet).
        - Para cada protocolo, genera **una fila por cada estudio** (P_Estudio, Estudio_2..5).
        - **Traduce nombres** que tienen alias conocido (T4 → T4 Total, Hemograma → Hemograma completo, etc.).
        - Si un estudio **no tiene alias**, deja el nombre del vete tal cual y **resalta la fila en amarillo** para que la revises.
        - **Borra el contenido previo** de `Planilla cobro` antes de escribir (asumí que la subís vacía).
        - **No toca** ninguna otra hoja del .xlsm (macros, perfiles, planillas de trabajo, etc. quedan iguales).
        """
    )

uploaded = st.file_uploader("Archivo .xlsm del día", type=["xlsm"])

if uploaded is None:
    st.info("Esperando archivo...")
    st.stop()

try:
    res = correr(uploaded.getvalue(), DATA_DIR / "aliases.json")
except ValueError as e:
    st.error(f"❌ {e}")
    st.stop()
except Exception as e:
    st.exception(e)
    st.stop()

m = res.metricas

st.subheader("Resumen")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Protocolos", m["protocolos_unicos"])
c2.metric("Estudios totales", m["estudios_totales"])
c3.metric("Reconocidos (catálogo + alias)", m["estudios_canonico"] + m["estudios_alias"])
c4.metric("A revisar", m["estudios_a_revisar"])

st.subheader("Distribución a planillas")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Planilla cobro", m["filas_planilla_cobro"])
c2.metric("Hemograma", m["filas_hemograma"])
c3.metric("Química", m["filas_quimica"])
c4.metric("Orinas", m["filas_orinas"])
c5.metric("Coagulograma", m["filas_coagulograma"])
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Serología", m["filas_serologia"])
c2.metric("Hemoparásitos", m["filas_hemoparasitos"])
c3.metric("VetCheck", m["filas_vetcheck"])
c4.metric("Hematología", m["filas_hematologia"])
c5.metric("Derivaciones", m["filas_derivaciones"])

col_dl1, col_dl2 = st.columns(2)
col_dl1.download_button(
    "⬇️ .xlsm del día (todas las planillas)",
    data=res.xlsm_bytes,
    file_name=uploaded.name.replace(".xlsm", " — generado.xlsm"),
    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
    type="primary",
)
if res.templado_bytes:
    col_dl2.download_button(
        f"⬇️ Templado V2 ({len(res.membretes)} membretes)",
        data=res.templado_bytes,
        file_name=uploaded.name.replace(".xlsm", " — derivaciones.xlsm"),
        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
    )

if m.get("derivaciones_por_destino"):
    st.subheader("Derivaciones por destino")
    cols = st.columns(min(len(m["derivaciones_por_destino"]), 6) or 1)
    for i, (destino, n) in enumerate(sorted(
        m["derivaciones_por_destino"].items(), key=lambda x: -x[1]
    )):
        cols[i % len(cols)].metric(destino, n)
    st.caption(
        "Se agregaron a la planilla Derivaciones según `data/derivables.json`. "
        "'EN EL LABO' = no sale del labo, pero igual queda registrado para imprimir membrete. "
        "'(definir)' = derivación que viene de un perfil con texto compuesto y no se pudo "
        "matchear con el catálogo — completala a mano en el Templado."
    )

if res.membretes:
    with st.expander(f"Ver {len(res.membretes)} membretes para imprimir", expanded=False):
        df_mem = pd.DataFrame(res.membretes)[
            ["codigo", "veterinaria", "paciente", "estudio", "destino"]
        ]
        df_mem.columns = ["Código", "Veterinaria", "Paciente", "Estudio", "Destino"]
        st.dataframe(df_mem, use_container_width=True, hide_index=True)

if res.a_revisar:
    st.subheader(f"⚠️ {len(res.a_revisar)} estudios a revisar")
    st.caption(
        "Estos estudios no figuran en el catálogo del .xlsm ni tienen alias. "
        "Quedaron en Planilla cobro con el nombre del vete tal cual y resaltados "
        "en amarillo. Revisalos y, si es una traducción recurrente, agregala al "
        "`data/aliases.json` para la próxima vez."
    )
    df = pd.DataFrame(
        [
            {
                "Código": f.codigo,
                "Veterinaria": f.veterinaria,
                "Estudio (texto del vete)": f.estudio_original,
            }
            for f in res.a_revisar
        ]
    )
    st.dataframe(df, use_container_width=True, hide_index=True)
else:
    st.success("✅ Todos los estudios fueron reconocidos.")

with st.expander("Métricas técnicas"):
    st.json(m)
