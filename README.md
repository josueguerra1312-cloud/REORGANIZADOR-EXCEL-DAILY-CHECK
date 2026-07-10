# Reorganizador Excel Daily Check

Esta versión corrige el error `ModuleNotFoundError: No module named openpyxl` agregando `openpyxl` en `requirements.txt`.

## Archivos en raíz

```text
app.py
procesador_programa.py
requirements.txt
README.md
.gitignore
```

## Importante para Streamlit Cloud

El archivo debe llamarse exactamente:

```text
requirements.txt
```

Debe estar en la raíz del repositorio, en la misma carpeta que `app.py`.

Si Streamlit ya había fallado antes, entra a:

```text
Manage app > Reboot app
```

o borra y vuelve a crear el deploy para forzar la instalación de dependencias.
