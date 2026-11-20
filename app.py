"""El Generador del Excel del día se jubiló el 20/11/2026.

Lo que hacía (llenar las planillas del .xlsm diario, los membretes de
derivación) ahora lo hace el LIMS del laboratorio. Esta página queda solo
para avisarle a quien todavía tenga guardada la dirección. La lógica vieja
sigue en `generador.py` y en el historial del repositorio.
"""

import streamlit as st

st.set_page_config(page_title="Generador DimeroLab · fuera de servicio", page_icon="🧪")

st.title("Este generador ya no está en uso")
st.write(
    "Desde el 20 de noviembre de 2026, el Excel del día, los membretes de derivación "
    "y el archivo para el iNCCA se hacen desde el LIMS del laboratorio."
)
st.info("Si necesitás la dirección del LIMS o un usuario, pedíselos a Administración.")
