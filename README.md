# ✝ Sistema Académico — Iglesia Pentecostal Fuente de Gracia

Sistema académico web para Instituto Bíblico. Construido con Streamlit + Excel.

## Archivos
```
app.py              ← Aplicación principal
requirements.txt    ← Dependencias
estudiantes.xlsx    ← Base de datos (se crea automático)
```

## Correr localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## 🚀 Subir a Streamlit Community Cloud (GRATIS)

### Paso 1 — Crear cuenta en GitHub
Ve a https://github.com y crea una cuenta gratuita si no tienes.

### Paso 2 — Crear repositorio
1. Clic en **"New repository"**
2. Nombre: `fuente-de-gracia-academico`
3. Márcalo como **Public**
4. Clic **"Create repository"**

### Paso 3 — Subir los archivos
1. En el repositorio creado, clic **"uploading an existing file"**
2. Arrastra y suelta estos archivos:
   - `app.py`
   - `requirements.txt`
3. Clic **"Commit changes"**

### Paso 4 — Conectar con Streamlit Cloud
1. Ve a https://share.streamlit.io
2. Inicia sesión con tu cuenta de GitHub
3. Clic **"New app"**
4. Selecciona tu repositorio `fuente-de-gracia-academico`
5. En **"Main file path"** escribe: `app.py`
6. Clic **"Deploy!"**

### Paso 5 — ¡Listo!
En 2-3 minutos tendrás un link como:
```
https://fuente-de-gracia-academico.streamlit.app
```
¡Compártelo con los maestros!

---

## ⚠️ Nota sobre Excel en la nube
Streamlit Cloud **no guarda archivos permanentemente** entre sesiones.
Para uso real con múltiples maestros, se recomienda:
- Exportar el Excel frecuentemente con el botón de descarga
- Subir el Excel actualizado al repositorio de GitHub cuando sea necesario

Para persistencia total, el siguiente paso sería conectar Google Sheets.
