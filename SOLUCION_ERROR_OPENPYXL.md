# Correccion ModuleNotFoundError openpyxl

El error `ModuleNotFoundError: No module named openpyxl` indica que Streamlit Cloud no instalo la dependencia `openpyxl`.

Verifica:

1. Que `requirements.txt` este en la raiz del repositorio, al mismo nivel que `app.py` y `procesador.py`.
2. Que el archivo se llame exactamente `requirements.txt`, en minusculas y sin extension adicional como `.txt.txt`.
3. Que dentro contenga:

```txt
streamlit>=1.36.0
pandas>=2.0.0
openpyxl>=3.1.2
xlrd>=2.0.1
```

Despues haz commit/push y reinicia la app en Streamlit Cloud desde Manage app > Reboot app.
