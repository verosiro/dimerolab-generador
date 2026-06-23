# Generador Planilla cobro · DimeroLab

App web que pre-carga la hoja `Planilla cobro` del .xlsm diario a partir de
lo que cargaron los vetes en AppSheet (hoja `ProtocolosDigitales`).

Es la Fase A de la visión Dimero LIMS: automatizar el llenado manual del
Excel del día para reducir errores de transcripción.

## Cómo usar

1. Pegá la tabla `Ingresos` de AppSheet en `ProtocolosDigitales` del .xlsm del día.
2. Subí el archivo a la app.
3. Descargá el .xlsm con `Planilla cobro` llena. Las filas resaltadas en
   amarillo son estudios sin traducción automática — revisalas.

## Correr localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estructura

- `app.py` — UI Streamlit.
- `generador.py` — lógica pura: parser xlsm + traducción + escritura.
- `data/aliases.json` — traducciones manuales nombre solicitado → nombre canónico.
- `data/quimicas_combinables.json` — catálogo de químicas combinables (reservado para etapa siguiente).

## Reglas que aplica

1. **Una fila por estudio.** Si un protocolo tiene 3 estudios (P_Estudio, Estudio_2, Estudio_3), genera 3 filas.
2. **Aliases manuales** (`data/aliases.json`): T4 → T4 Total, etc.
3. **Fecha en español** ("21 mayo jueves").
4. **Si no hay alias**, deja el texto del vete y resalta la fila para revisar.
5. **Borra cualquier contenido previo** debajo del header de `Planilla cobro`.
6. **No toca** ninguna otra hoja del .xlsm (macros, perfiles, etc.).

## Próximas etapas

- Detección de "N químicas" cuando los pedidos individuales son combinables.
- Expandir perfiles en sus componentes si la planilla del día lo necesita.
- Tomar `Propietario/Especie/Raza` automáticamente del cache de AppSheet si vienen vacíos.
