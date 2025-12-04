import requests
from bs4 import BeautifulSoup
import zipfile
import re

URL_VISTADATOS = "https://www.coes.org.pe/Portal/browser/vistadatos"
URL_DOWNLOAD = "https://www.coes.org.pe/Portal/browser/download"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "X-Requested-With": "XMLHttpRequest",
}

def normalizar_mes(nombre: str) -> str:
    nombre = nombre.strip().replace("-", " ")
    nombre = re.sub(r"\s+", " ", nombre)
    partes = nombre.split(" ")

    if len(partes) == 2:
        num, mes = partes
        return f"{num.zfill(2)}_{mes.capitalize()}"

    if "_" in nombre:
        num, mes = nombre.split("_", 1)
        return f"{num.zfill(2)}_{mes.capitalize()}"

    return nombre

def listar_directorio(ruta: str, indicador: str):
    data = {
        "baseDirectory": ruta,
        "url": ruta,
        "indicador": indicador,
        "initialLink": "IEOD",
        "orderFolder": "D",
    }
    r = requests.post(URL_VISTADATOS, headers=HEADERS, data=data)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser")

def obtener_anios():
    soup = listar_directorio("Post Operación/Reportes/IEOD/", "S")
    items = soup.select("li.infolist-item a.infolist-link")
    return [a.text.strip() for a in items if a.text.strip().isdigit()]

def obtener_meses(anio: str):
    ruta = f"Post Operación/Reportes/IEOD/{anio}/"
    soup = listar_directorio(ruta, "S")

    meses = []
    for li in soup.select("li.infolist-item a.infolist-link"):
        txt = li.text.strip()
        if any(c.isdigit() for c in txt):
            meses.append(normalizar_mes(txt))
    return meses

def obtener_dias(anio: str, mes: str):
    ruta = f"Post Operación/Reportes/IEOD/{anio}/{mes}/"
    soup = listar_directorio(ruta, "S")
    return [li.text.strip() for li in soup.select("li.infolist-item a.infolist-link") if li.text.strip().isdigit()]

def obtener_archivos_del_dia(anio: str, mes: str, dia: str):
    ruta = f"Post Operación/Reportes/IEOD/{anio}/{mes}/{dia}/"

    data = {
        "baseDirectory": ruta,
        "url": ruta,
        "indicador": "N",
        "initialLink": "IEOD",
        "orderFolder": "D",
    }

    r = requests.post(URL_VISTADATOS, headers=HEADERS, data=data)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    archivos = []
    for tr in soup.select("tr.selector-file-contextual"):
        celdas = tr.find_all("td")
        ruta_archivo = tr.get("id")
        if not ruta_archivo:
            continue
        ruta_archivo = ruta_archivo.replace("&#243;", "ó").replace("&#237;", "í")

        archivos.append({
            "ruta": ruta_archivo,
            "nombre": celdas[2].text.strip(),
        })
    return archivos

def pod_obtener_archivo_despacho(anio: str, mes: str, dia: str):
    num, nombre = mes.split("_", 1)
    mes_pod = f"{num}_{nombre.upper()}"

    dia_pod = f"Día {dia.zfill(2)}"

    ruta = f"Operación/Programa de Operación/Programa Diario/{anio}/{mes_pod}/{dia_pod}/"

    data = {
        "baseDirectory": ruta,
        "url": ruta,
        "indicador": "N",
        "initialLink": "IEOD",
        "orderFolder": "D",
    }

    try:
        r = requests.post(URL_VISTADATOS, headers=HEADERS, data=data)
        r.raise_for_status()
    except Exception:
        return []

    soup = BeautifulSoup(r.text, "html.parser")

    archivos = []

    for tr in soup.select("tr.selector-file-contextual"):
        celdas = tr.find_all("td")
        if len(celdas) < 3:
            continue

        nombre_archivo = celdas[2].text.strip().lower()

        if nombre_archivo.startswith("anexo1_despacho") and nombre_archivo.endswith(".xlsx"):
            ruta_archivo = tr.get("id", "").replace("&#243;", "ó").replace("&#237;", "í")

            archivos.append({
                "ruta": ruta_archivo,
                "nombre": celdas[2].text.strip(),
            })

    return archivos

def crear_sesion_coes():
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Accept": "*/*",
    })
    try:
        s.get("https://www.coes.org.pe/Portal/PostOperacion/Reportes/Ieod", timeout=10)
    except:
        pass
    return s

def descargar_archivo(ruta_archivo: str, destino_local: str):
    ruta_codificada = requests.utils.quote(ruta_archivo, safe="/:")
    url = f"{URL_DOWNLOAD}?url={ruta_codificada}"

    with requests.get(url, headers=HEADERS, stream=True) as r:
        r.raise_for_status()
        with open(destino_local, "wb") as f:
            for chunk in r.iter_content(8192):
                if chunk:
                    f.write(chunk)

def descargar_zip_real_con_sesion(session, ruta_zip: str, destino_zip: str):
    params = {"url": ruta_zip}
    r = session.get(URL_DOWNLOAD, params=params, stream=True)
    r.raise_for_status()

    with open(destino_zip, "wb") as f:
        for chunk in r.iter_content(8192):
            if chunk:
                f.write(chunk)

    if not zipfile.is_zipfile(destino_zip):
        return False
    return True