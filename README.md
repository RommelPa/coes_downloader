# 📘 COES Downloader

Automatizador avanzado para descarga masiva de archivos **IEOD** y
**Programa Diario (POD)** desde el portal del **COES Perú**.

## 🚀 Características Principales

✔ Descarga automatizada de:

-   \*\*AnexoA\_\*.xlsx\*\* (IEOD)\
-   **XLSX internos de archivos ZIP del IEOD**\
-   **Anexo1_Despacho_YYYYMMDD.xlsx** (Programa Diario)

✔ Navegación por:

-   Año (solo **2025 en adelante**, detectados automáticamente desde el
    portal)
-   Mes
-   Día

✔ Multithreading integrado: - Hasta **5 descargas simultáneas** -
Velocidad **3× a 7× más rápida**

✔ Interfaz moderna con **PySide6**\
✔ Barra de progreso en tiempo real\
✔ Organización automática por carpetas (Año → Mes → Día)\
✔ Compatible con futuras actualizaciones del COES\
✔ Código limpio, modular y mantenible

## 📂 Estructura del Proyecto

    app/
     ├── main.py                   # Interfaz gráfica PySide6
     └── downloader/
          ├── fetch_coes.py        # Lógica de scraping y descargas
          └── __init__.py
    requirements.txt
    COESDownloader.spec             # Configuración PyInstaller
    README.md

## 📦 Instalación

### 1️⃣ Clonar el repositorio

``` bash
git clone https://github.com/RommelPa/coes_downloader.git
cd coes_downloader
```

### 2️⃣ Crear entorno virtual

``` bash
python -m venv venv
```

### 3️⃣ Activar entorno

Windows:

``` bash
venv\Scripts\activate
```

Linux/Mac:

``` bash
source venv/bin/activate
```

### 4️⃣ Instalar dependencias

``` bash
pip install -r requirements.txt
```

## ▶️ Ejecutar la aplicación

``` bash
python -m app.main
```

## 🛠️ Generar el ejecutable (.exe)

### Usar el .spec personalizado:

``` bash
python -m pyinstaller COESDownloader.spec
```

El ejecutable final estará en:

    dist/COESDownloader/COESDownloader.exe