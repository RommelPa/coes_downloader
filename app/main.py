import os
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QListWidget, QLabel,
    QPushButton, QFileDialog, QMessageBox, QProgressBar
)

from app.downloader.fetch_coes import (
    obtener_anios,
    obtener_meses,
    obtener_dias,
    obtener_archivos_del_dia,
    pod_obtener_archivo_despacho,
    descargar_archivo,
    crear_sesion_coes,
    descargar_zip_real_con_sesion,
)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("COES Downloader")
        self.resize(550, 650)

        layout = QVBoxLayout()

        layout.addWidget(QLabel("Años disponibles:"))
        self.list_anios = QListWidget()
        layout.addWidget(self.list_anios)

        layout.addWidget(QLabel("Meses del año:"))
        self.list_meses = QListWidget()
        layout.addWidget(self.list_meses)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        layout.addWidget(self.progress)

        self.btn_descargar_mes = QPushButton("Descargar mes")
        layout.addWidget(self.btn_descargar_mes)

        self.setLayout(layout)

        self.cargar_anios()

        self.list_anios.itemSelectionChanged.connect(self.cargar_meses)
        self.btn_descargar_mes.clicked.connect(self.descargar_mes)

        self.session = crear_sesion_coes()

    def descargar_archivo_tipo(self, tipo, archivo, carpeta_dia):
        """
        Ejecuta la descarga según el tipo (ieod, zip o pod).
        Se ejecuta en un hilo del ThreadPoolExecutor.
        Devuelve True si se guardó correctamente.
        """
        try:
            nombre = archivo["nombre"].lower()

            if tipo == "ieod" and nombre.startswith("anexoa") and nombre.endswith(".xlsx"):
                destino = os.path.join(carpeta_dia, archivo["nombre"])
                descargar_archivo(archivo["ruta"], destino)
                return True

            if tipo == "ieod" and nombre.endswith(".zip"):
                zip_temp = os.path.join(carpeta_dia, archivo["nombre"])
                ok = descargar_zip_real_con_sesion(self.session, archivo["ruta"], zip_temp)

                if not ok:
                    return False

                with zipfile.ZipFile(zip_temp, "r") as z:
                    xlsx_files = [p for p in z.namelist() if p.lower().endswith(".xlsx")]
                    if not xlsx_files:
                        return False

                    destino = os.path.join(carpeta_dia, os.path.basename(xlsx_files[0]))
                    with open(destino, "wb") as out:
                        out.write(z.read(xlsx_files[0]))

                try:
                    os.remove(zip_temp)
                except:
                    pass

                return True

            if tipo == "pod":
                destino = os.path.join(carpeta_dia, archivo["nombre"])
                descargar_archivo(archivo["ruta"], destino)
                return True

        except:
            return False

        return False

    def cargar_anios(self):
        """
        Mostrar solo 2025 y años posteriores.
        """
        self.list_anios.clear()

        try:
            anios = obtener_anios()
            anios_filtrados = [a for a in anios if a.isdigit() and int(a) >= 2025]

            if not anios_filtrados:
                self.list_anios.addItem("2025")
                return

            for anio in sorted(anios_filtrados, reverse=True):
                self.list_anios.addItem(anio)

        except Exception:
            self.list_anios.addItem("2025")

    def cargar_meses(self):
        self.list_meses.clear()
        if not self.list_anios.selectedItems():
            return

        anio = self.list_anios.selectedItems()[0].text()
        meses = obtener_meses(anio)
        for m in meses:
            self.list_meses.addItem(m)

    def descargar_mes(self):
        if not self.list_anios.selectedItems() or not self.list_meses.selectedItems():
            QMessageBox.warning(self, "Error", "Seleccione año y mes.")
            return

        anio = self.list_anios.selectedItems()[0].text()
        mes = self.list_meses.selectedItems()[0].text()

        carpeta = QFileDialog.getExistingDirectory(self, "Carpeta destino")
        if not carpeta:
            return

        dias = obtener_dias(anio, mes)

        tareas = []
        total_archivos = 0

        for d in dias:

            carpeta_dia = os.path.join(carpeta, d)
            os.makedirs(carpeta_dia, exist_ok=True)

            archivos_ieod = obtener_archivos_del_dia(anio, mes, d)

            archivos_pod = pod_obtener_archivo_despacho(anio, mes, d)

            for a in archivos_ieod:
                nombre = a["nombre"].lower()
                if nombre.endswith(".pdf"):
                    continue
                tareas.append(("ieod", a, carpeta_dia))
                total_archivos += 1

            for a in archivos_pod:
                tareas.append(("pod", a, carpeta_dia))
                total_archivos += 1

        if total_archivos == 0:
            QMessageBox.information(self, "Nada que descargar", "No se encontraron archivos XLSX.")
            return

        progreso = 0
        errores = 0
        guardados = 0

        executor = ThreadPoolExecutor(max_workers=5)
        futures = []

        for tipo, archivo, carpeta_dia in tareas:
            f = executor.submit(self.descargar_archivo_tipo, tipo, archivo, carpeta_dia)
            futures.append(f)

        for future in as_completed(futures):
            ok = future.result()

            if ok:
                guardados += 1
            else:
                errores += 1

            progreso += 1
            porcentaje = int(progreso * 100 / total_archivos)
            self.progress.setValue(porcentaje)
            QApplication.processEvents()

        self.progress.setValue(100)
        QMessageBox.information(
            self,
            "Descarga completa",
            f"Archivos guardados: {guardados}\nErrores: {errores}"
        )
        self.progress.setValue(0)

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()