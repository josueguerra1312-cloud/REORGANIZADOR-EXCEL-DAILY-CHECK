# Reorganizador Excel Daily Check

App en Python y Streamlit para convertir el Excel diario de programa de mantenimiento al formato solicitado.

## Estructura correcta del repositorio

Coloca estos archivos directamente en la raiz del repositorio:

```text
app.py
procesador_programa.py
requirements.txt
README.md
.gitignore
```

No subas una carpeta contenedora. `app.py` y `procesador_programa.py` deben estar en la misma ubicacion.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Desplegar en Streamlit Cloud

1. Sube los archivos a GitHub.
2. Verifica que `requirements.txt` este en la raiz.
3. En Streamlit Cloud selecciona `app.py` como archivo principal.
4. Si Streamlit conserva errores anteriores, usa `Manage app`, luego `Clear cache` y `Reboot app`.

## Logica de ordenamiento

Dentro de cada grupo AC + WO, las tareas se ordenan asi:

1. WEEKLY CHECK
2. DAILY CHECK
3. EO normales
4. Defectos o controles

La descripcion usa `eo_description` completo y agrega `P/N` y `S/N` solo cuando ambos campos existen.
