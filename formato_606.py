# -*- coding: utf-8 -*-
"""
formato_606.py
==============
Genera el archivo TXT con formato 606 (DGII) a partir de las facturas en PDF
guardadas en la carpeta 'facturas_606' y registra los gastos en la hoja
'GASTOS606' del libro 'ingresos_gastos.xlsx'.

Los PDF son imágenes escaneadas, por lo que se aplica OCR (Tesseract) para
leer el texto y se usan las coordenadas de cada palabra para asociar cada
etiqueta ('RNC:', 'e-NCF:', 'Subtotal Gravado:', 'Total ITBIS:', ...)
con el valor que está a su derecha en el mismo renglón visual.

Ejecutar:  python formato_606.py
Requiere:  PySide6, PyMuPDF (fitz), pytesseract, Pillow, openpyxl y Tesseract-OCR.
"""

import io
import os
import re
import shutil
import sys
import unicodedata
from calendar import monthrange
from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import fitz  # PyMuPDF
import openpyxl
import pytesseract
from PIL import Image
from PySide6.QtCore import QThread, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

# ------------------------------------------------------------------ rutas
RUTA_PROYECTO = Path(__file__).resolve().parent
CARPETA_PDFS = RUTA_PROYECTO / "facturas_606"
RUTA_EXCEL = RUTA_PROYECTO / "ingresos_gastos.xlsx"

# ------------------------------------------------------------ constantes
FORMATO606 = "606"            # fijo: formato del reporte
RNC_CLIENTE = "00105336614"   # fijo: RNC del reporte (aparece en el encabezado)
HOJA_GASTOS = "GASTOS606"
ENCABEZADO_EXCEL = ["AÑO", "MES", "RNC", "NCF", "MONTO_SERVICIOS",
                    "MONTO_BIENES", "TOTAL", "ITBIS"]

# RNC de proveedores cuyo gasto es de tipo '02' (servicios); el resto es '09'
RNCS_TIPO_GASTO_02 = {"101820217", "101001577"}
FORMA_PAGO = "04"             # columna 23, valor fijo

NOMBRES_MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                 "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre",
                 "Diciembre"]

DPI_OCR = 300


# ============================================================ fechas
def mes_anterior(hoy=None):
    """(año, mes) de la fecha del sistema con un mes menos.
    Si el mes actual es enero, se le resta también uno al año."""
    hoy = hoy or date.today()
    anyo, mes = hoy.year, hoy.month - 1
    if mes == 0:
        anyo, mes = anyo - 1, 12
    return anyo, mes


def fecha_encabezado(hoy=None):
    """Variable 'fecha_encabezado': 'yyyymm' de la fecha del sistema con un
    mes menos. Ej.: 11/11/2026 -> '202610'; 01/01/2027 -> '202612'."""
    anyo, mes = mes_anterior(hoy)
    return f"{anyo:04d}{mes:02d}"


def fecha_anyo(hoy=None):
    """Mismo procedimiento que 'fecha_encabezado' pero en formato 'yyyymmdd'.
    Se usa de respaldo cuando el PDF no trae fecha de emisión/vencimiento."""
    anyo, mes = mes_anterior(hoy)
    hoy = hoy or date.today()
    dia = min(hoy.day, monthrange(anyo, mes)[1])
    return f"{anyo:04d}{mes:02d}{dia:02d}"


def parsear_fecha(texto):
    """Convierte una fecha tipo '05-11-2026' o '05/11/2026' a 'yyyymmdd'."""
    m = re.search(r"(\d{1,2})\s*[-/]\s*(\d{1,2})\s*[-/]\s*(\d{2,4})", texto)
    if not m:
        return None
    dia, mes, anyo = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if anyo < 100:
        anyo += 2000
    try:
        return date(anyo, mes, dia).strftime("%Y%m%d")
    except ValueError:
        return None


# ============================================================ montos
def parsear_monto(texto):
    """Convierte un monto tipo 'RD$ 19,548.24', '1.118,64' o '2951' a float."""
    texto = re.sub(r"[^\d\.,\-]", "", str(texto))
    if texto in ("", "-"):
        return 0.0
    negativo = texto.startswith("-")
    texto = texto.lstrip("-")
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):      # 1.234,56
            texto = texto.replace(".", "").replace(",", ".")
        else:                                        # 1,234.56
            texto = texto.replace(",", "")
    elif "," in texto:
        partes = texto.split(",")
        if len(partes[-1]) in (1, 2):                # 2951,50
            texto = "".join(partes[:-1]) + "." + partes[-1]
        else:                                        # 2,951 (miles)
            texto = texto.replace(",", "")
    elif "." in texto:
        partes = texto.split(".")
        if len(partes) > 2:                          # 1.234.567 (miles)
            texto = "".join(partes)
        elif len(partes[-1]) == 3:                   # 19.548 (miles)
            texto = "".join(partes)
    try:
        valor = float(texto)
    except ValueError:
        return 0.0
    return -valor if negativo else valor


def fmt_monto(valor):
    """Formato del TXT: texto redondeado a entero (0.5 sube), sin separadores."""
    return str(int(Decimal(str(round(valor + 0.0, 2))).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP)))


# ============================================================ OCR
def _idioma_disponible(idioma, ruta_tesseract):
    """Comprueba si existe el archivo <idioma>.traineddata de Tesseract."""
    rutas = []
    prefijo = os.environ.get("TESSDATA_PREFIX")
    if prefijo:
        rutas.append(Path(prefijo))
    if ruta_tesseract:
        rutas.append(Path(ruta_tesseract).parent / "tessdata")
    rutas.append(Path(os.environ.get("PROGRAMFILES", "C:\\Program Files"))
                 / "Tesseract-OCR" / "tessdata")
    return any((base / f"{idioma}.traineddata").exists() for base in rutas)


def preparar_tesseract():
    """Localiza Tesseract y devuelve el idioma disponible ('spa' o 'eng')."""
    cmd = shutil.which("tesseract")
    if not cmd:
        candidato = (Path(os.environ.get("LOCALAPPDATA", ""))
                     / "Programs" / "Tesseract-OCR" / "tesseract.exe")
        if candidato.exists():
            cmd = str(candidato)
    if cmd:
        pytesseract.pytesseract.tesseract_cmd = cmd
    return "spa" if _idioma_disponible("spa", cmd) else "eng"


def normalizar(texto):
    """Mayúsculas y sin acentos, para comparar etiquetas aunque el OCR falle."""
    texto = unicodedata.normalize("NFD", texto.upper())
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def pdf_a_lineas(ruta_pdf, idioma="spa", log=print):
    """Aplica OCR al PDF y reconstruye los renglones visuales de cada página.
    Devuelve una lista de dicts {'pagina', 'cy', 'texto'}."""
    lineas = []
    with fitz.open(ruta_pdf) as documento:
        total = len(documento)
        for num, pagina in enumerate(documento, 1):
            log(f"   OCR página {num}/{total}…")
            pix = pagina.get_pixmap(dpi=DPI_OCR)
            imagen = Image.open(io.BytesIO(pix.tobytes("png")))
            datos = pytesseract.image_to_data(
                imagen, lang=idioma, output_type=pytesseract.Output.DICT)

            palabras = []
            for i in range(len(datos["text"])):
                texto = datos["text"][i].strip()
                if not texto or int(float(datos["conf"][i])) < 0:
                    continue
                x = datos["left"][i]
                y = datos["top"][i]
                h = datos["height"][i]
                palabras.append({"t": texto, "x": x, "h": h,
                                 "cy": y + h / 2})
            if not palabras:
                continue

            # Agrupar palabras por renglón visual (centro-y parecido),
            # sin importar el orden en que Tesseract entregó los bloques.
            palabras.sort(key=lambda p: p["cy"])
            altura_mediana = sorted(p["h"] for p in palabras)[len(palabras) // 2]
            tolerancia = max(altura_mediana * 0.6, 10)
            grupos = []
            for p in palabras:
                if grupos and abs(p["cy"] - grupos[-1]["cy_ref"]) <= tolerancia:
                    g = grupos[-1]
                    g["palabras"].append(p)
                    g["cy_ref"] = ((g["cy_ref"] * (len(g["palabras"]) - 1)
                                    + p["cy"]) / len(g["palabras"]))
                else:
                    grupos.append({"cy_ref": p["cy"], "palabras": [p]})
            for g in grupos:
                texto = " ".join(p["t"] for p in
                                 sorted(g["palabras"], key=lambda p: p["x"]))
                lineas.append({"pagina": num, "cy": g["cy_ref"], "texto": texto})
    return lineas


# ============================================================ extracción
def monto_despues_de(texto, patron_etiqueta):
    """Valor numérico que está a la derecha de una etiqueta.
    Devuelve None si la etiqueta no está o si su valor está en otro renglón."""
    m = re.search(patron_etiqueta, texto)
    if not m:
        return None
    resto = texto[m.end():]
    if not resto.strip():
        return None                                   # valor en otro renglón
    if re.fullmatch(r"[\s\-–—]*", resto):
        return 0.0                                    # '-' significa cero
    m_num = re.search(r"-?[\d][\d\.,]*", resto)
    if not m_num:
        return None
    return parsear_monto(m_num.group(0))


def extraer_datos_pdf(ruta_pdf, idioma="spa", log=print):
    """Lee un PDF con OCR y devuelve los campos de la factura."""
    campos = {
        "rnc": None, "rnc_cliente": None, "ncf": None, "ncf_modificado": None,
        "fecha_emision": None, "fecha_vencimiento": None,
        "subtotal_gravado": None, "subtotal_exento": None, "itbis": None,
        "isce": None, "iscav": None, "otros_impuestos": None,
        "propina_legal": None,
    }
    reglas = (
        ("rnc", r"\bRNC\s*:\s*(\d{6,})"),
        ("rnc_cliente", r"\bRNC\s+CLIENTE\s*:\s*(\d{6,})"),
        ("ncf", r"\bE\s*-\s*NCF\s*:\s*([A-Z0-9]{8,})"),
        ("ncf_modificado", r"NCF\s+MODIFICADO\s*:\s*([A-Z0-9]{8,})"),
        ("fecha_emision", r"FECHA\s+EMISION\s*:\s*(\S+(?:\s+\S+)?)"),
        ("fecha_vencimiento", r"FECHA\s+VENCIMIENTO\s*:\s*(\S+(?:\s+\S+)?)"),
        ("subtotal_gravado", r"SUBTOTAL\s+GRAVADO\s*:"),
        ("subtotal_exento", r"SUBTOTAL\s+EXENTO\s*:"),
        ("itbis", r"TOTAL\s+ITBIS\s*:"),
        ("isce", r"TOTAL\s+I.{0,3}SCE\s*:"),
        ("iscav", r"TOTAL\s+I.{0,3}SCAV\s*:"),
        ("otros_impuestos", r"\bCD.{0,2}T\s*:"),
        ("propina_legal", r"PROPINA\s+LEGAL\s*:"),
    )
    for linea in pdf_a_lineas(ruta_pdf, idioma, log):
        texto = normalizar(linea["texto"])
        for campo, patron in reglas:
            if campo in ("fecha_emision", "fecha_vencimiento"):
                if campos[campo] is None:
                    m = re.search(patron, texto)
                    if m:
                        campos[campo] = parsear_fecha(m.group(1))
            elif campo in ("subtotal_gravado", "subtotal_exento", "itbis",
                           "isce", "iscav", "otros_impuestos", "propina_legal"):
                if campos[campo] is None:
                    campos[campo] = monto_despues_de(texto, patron)
            elif campos[campo] is None:
                m = re.search(patron, texto)
                if m:
                    campos[campo] = m.group(1)
    return campos


# ============================================================ fila 606
def construir_fila_detalle(campos, log=print):
    """Construye la fila de detalle (23 columnas separadas por '|').
    Devuelve (fila_txt, datos_para_excel)."""
    rnc_proveedor = re.sub(r"\D", "", campos["rnc"] or "")
    tipo_rnc = "1" if len(rnc_proveedor) == 9 else "2"
    tipo_gastos = "02" if rnc_proveedor in RNCS_TIPO_GASTO_02 else "09"

    ncf = (campos["ncf"] or "").upper()
    ncf_modificado = ""
    if ncf.startswith("E34"):   # nota de crédito: se reporta el NCF modificado
        ncf_modificado = (campos["ncf_modificado"] or "").upper()

    subtotal = round((campos["subtotal_gravado"] or 0)
                     + (campos["subtotal_exento"] or 0), 2)
    monto_servicio = subtotal if tipo_gastos == "02" else 0.0
    monto_bienes = subtotal if tipo_gastos == "09" else 0.0
    monto_total = round(monto_servicio + monto_bienes, 2)
    itbis = round(campos["itbis"] or 0, 2)
    selectivo = round((campos["isce"] or 0) + (campos["iscav"] or 0), 2)
    otros = round(campos["otros_impuestos"] or 0, 2)
    propina = round(campos["propina_legal"] or 0, 2)

    # Columnas 6 y 7: 'Fecha Emisión' de la factura en 'yyyymmdd'.
    # Si el PDF no trae la fecha, se usa la fecha del proceso (fecha_anyo).
    fecha_comprobante = campos["fecha_emision"] or fecha_anyo()
    fecha_vencimiento = fecha_comprobante

    columnas = [
        rnc_proveedor,               # 1  RNC del proveedor
        tipo_rnc,                    # 2  '1' si 9 dígitos, '2' si 11
        tipo_gastos,                 # 3  '02' o '09' según el RNC
        ncf,                         # 4  e-NCF
        ncf_modificado,              # 5  NCF modificado (solo si el NCF es E34…)
        fecha_comprobante,           # 6  fecha comprobante 'yyyymmdd'
        fecha_vencimiento,           # 7  fecha vencimiento 'yyyymmdd'
        fmt_monto(monto_servicio),   # 8  monto de servicios
        fmt_monto(monto_bienes),     # 9  monto de bienes
        str(int(fmt_monto(monto_servicio)) + int(fmt_monto(monto_bienes))),  # 10 monto total
        fmt_monto(itbis),            # 11 ITBIS facturado
        "0", "0", "0",               # 12, 13, 14
        fmt_monto(itbis),            # 15 ITBIS
        "0",                         # 16
        "",                          # 17 (en blanco)
        "0", "0",                    # 18, 19
        fmt_monto(selectivo),        # 20 selectivo al consumo (ISCe + ISCav)
        fmt_monto(otros),            # 21 otros impuestos (CDT)
        fmt_monto(propina),          # 22 propina legal
        FORMA_PAGO,                  # 23 forma de pago
    ]

    if not rnc_proveedor:
        log("   AVISO: no se encontró el RNC del proveedor en el PDF.")
    if not ncf:
        log("   AVISO: no se encontró el e-NCF en el PDF.")

    fila_excel = {
        "rnc": rnc_proveedor,
        "ncf": ncf,
        "monto_servicios": monto_servicio,
        "monto_bienes": monto_bienes,
        "total": monto_total,
        "itbis": itbis,
    }
    return "|".join(columnas), fila_excel


# ============================================================ proceso
def exportar_excel(filas, ruta_xlsx, periodo, log=print):
    """Agrega las filas a la hoja GASTOS606 a partir de la fila 2 y guarda."""
    ruta_xlsx = Path(ruta_xlsx)
    if ruta_xlsx.exists():
        wb = openpyxl.load_workbook(ruta_xlsx)
    else:
        wb = openpyxl.Workbook()
    ws = wb[HOJA_GASTOS] if HOJA_GASTOS in wb.sheetnames else wb.create_sheet(HOJA_GASTOS)
    if ws.max_row < 1 or all(celda.value in (None, "") for celda in ws[1]):
        ws.append(ENCABEZADO_EXCEL)

    fila = 2                       # primera fila de datos (después del encabezado)
    while ws.cell(fila, 1).value not in (None, ""):
        fila += 1

    anyo = int(periodo[:4])
    mes = NOMBRES_MESES[int(periodo[4:6])]
    for i, f in enumerate(filas):
        r = fila + i
        ws.cell(r, 1, anyo)                       # AÑO (número)
        ws.cell(r, 2, mes)                        # MES (nombre en español)
        ws.cell(r, 3, f["rnc"])                   # RNC
        ws.cell(r, 4, f["ncf"])                   # NCF
        ws.cell(r, 5, f["monto_servicios"])       # MONTO_SERVICIOS
        ws.cell(r, 6, f["monto_bienes"])          # MONTO_BIENES
        ws.cell(r, 7, f["total"])                 # TOTAL
        ws.cell(r, 8, f["itbis"])                 # ITBIS
    try:
        wb.save(ruta_xlsx)
    except PermissionError:
        raise RuntimeError(
            "No se pudo guardar 'ingresos_gastos.xlsx'. Cierre el archivo en "
            "Excel y vuelva a procesar (el TXT ya quedó generado).")
    log(f"Excel actualizado: {ruta_xlsx} (hoja {HOJA_GASTOS}, desde la fila {fila})")


def procesar_carpeta(carpeta, log=print, ruta_txt=None, ruta_excel=None):
    """Recorre los PDF de la carpeta, arma el TXT 606 y actualiza el Excel."""
    carpeta = Path(carpeta)
    ruta_txt = Path(ruta_txt) if ruta_txt else RUTA_PROYECTO / (
        f"DGII_F_606_{RNC_CLIENTE}_{fecha_encabezado()}.txt")
    ruta_excel = Path(ruta_excel) if ruta_excel else RUTA_EXCEL

    pdfs = sorted(p for p in carpeta.glob("*.pdf") if not p.name.startswith("~$"))
    if not pdfs:
        raise RuntimeError(f"No se encontraron archivos PDF en:\n{carpeta}")

    periodo = fecha_encabezado()
    log(f"Período a reportar: {periodo}")
    log(f"PDF a procesar: {len(pdfs)}")

    filas_detalle = []
    filas_excel = []
    idioma = preparar_tesseract()
    for pdf in pdfs:
        log("")
        log(f"Leyendo: {pdf.name}")
        campos = extraer_datos_pdf(pdf, idioma, log)
        fila_txt, fila_excel = construir_fila_detalle(campos, log)
        filas_detalle.append(fila_txt)
        filas_excel.append(fila_excel)
        datos_fila = fila_txt.split("|")
        log("   RNC: {} | tipo RNC: {} | gasto: {} | e-NCF: {} | modificado: {}".format(
            datos_fila[0], datos_fila[1], datos_fila[2],
            fila_excel["ncf"], datos_fila[4]))
        log("   Emisión: {} | Vencimiento: {} | Gravado: {} | Exento: {} | ITBIS: {}".format(
            campos["fecha_emision"] or "-", campos["fecha_vencimiento"] or "-",
            fmt_monto(campos["subtotal_gravado"] or 0),
            fmt_monto(campos["subtotal_exento"] or 0),
            fmt_monto(campos["itbis"] or 0)))

    encabezado = f"{FORMATO606}|{RNC_CLIENTE}|{periodo}|{len(filas_detalle)}"
    contenido = "\n".join([encabezado] + filas_detalle) + "\n"
    ruta_txt.write_text(contenido, encoding="utf-8")
    log("")
    log(f"Archivo TXT creado: {ruta_txt}")
    log(f"Encabezado: {encabezado}")

    exportar_excel(filas_excel, ruta_excel, periodo, log)
    return ruta_txt, len(filas_detalle)


# ============================================================ interfaz
class ProcesadorHilo(QThread):
    """Procesa los PDF en segundo plano para no congelar la ventana."""
    senal_log = Signal(str)
    senal_exito = Signal(str, int)
    senal_error = Signal(str)

    def __init__(self, carpeta):
        super().__init__()
        self.carpeta = carpeta

    def run(self):
        try:
            ruta_txt, n = procesar_carpeta(self.carpeta, log=self.senal_log.emit)
            self.senal_exito.emit(str(ruta_txt), n)
        except Exception as exc:
            self.senal_error.emit(str(exc))


class VentanaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Formato 606 - DGII")
        self.resize(860, 580)
        self.hilo = None

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        periodo = fecha_encabezado()
        info = QLabel(f"RNC Cliente (fijo): {RNC_CLIENTE}      "
                      f"Período a reportar: {periodo}")
        layout.addWidget(info)

        fila_carpeta = QHBoxLayout()
        fila_carpeta.addWidget(QLabel("Carpeta de PDFs:"))
        self.txt_carpeta = QLineEdit(str(CARPETA_PDFS))
        fila_carpeta.addWidget(self.txt_carpeta, 1)
        btn_examinar = QPushButton("Examinar…")
        btn_examinar.clicked.connect(self.elegir_carpeta)
        fila_carpeta.addWidget(btn_examinar)
        layout.addLayout(fila_carpeta)

        self.btn_procesar = QPushButton("Procesar facturas (PDF → TXT 606 + Excel)")
        self.btn_procesar.setMinimumHeight(42)
        self.btn_procesar.clicked.connect(self.procesar)
        layout.addWidget(self.btn_procesar)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText(
            "Aquí se mostrará el detalle del proceso…")
        layout.addWidget(self.log, 1)

        self.btn_abrir = QPushButton("Abrir carpeta del proyecto (TXT y Excel)")
        self.btn_abrir.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(RUTA_PROYECTO))))
        layout.addWidget(self.btn_abrir)

    def elegir_carpeta(self):
        carpeta = QFileDialog.getExistingDirectory(
            self, "Seleccione la carpeta con los PDF", self.txt_carpeta.text())
        if carpeta:
            self.txt_carpeta.setText(carpeta)

    def procesar(self):
        carpeta = self.txt_carpeta.text().strip()
        if not Path(carpeta).is_dir():
            QMessageBox.warning(self, "Carpeta no válida",
                                f"La carpeta no existe:\n{carpeta}")
            return
        self.btn_procesar.setEnabled(False)
        self.log.appendPlainText("=" * 60)
        self.hilo = ProcesadorHilo(carpeta)
        self.hilo.senal_log.connect(self.log.appendPlainText)
        self.hilo.senal_exito.connect(self.al_terminar)
        self.hilo.senal_error.connect(self.al_error)
        self.hilo.start()

    def al_terminar(self, ruta_txt, cantidad):
        self.btn_procesar.setEnabled(True)
        self.log.appendPlainText("")
        self.log.appendPlainText(f"✔ Proceso terminado: {cantidad} comprobante(s).")
        QMessageBox.information(
            self, "Proceso terminado",
            f"Se generó el archivo:\n{ruta_txt}\n\n"
            f"Comprobantes procesados: {cantidad}\n"
            f"Excel actualizado: {RUTA_EXCEL.name}")

    def al_error(self, mensaje):
        self.btn_procesar.setEnabled(True)
        self.log.appendPlainText("")
        self.log.appendPlainText(f"✖ ERROR: {mensaje}")
        QMessageBox.critical(self, "Error", mensaje)


def main():
    app = QApplication(sys.argv)
    ventana = VentanaPrincipal()
    ventana.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
