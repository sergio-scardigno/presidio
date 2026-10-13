"""Anonimiza documentos con Microsoft Presidio.

Uso:
    python anonimizar.py                 # procesa todo input/ y deja el resultado en output/
    python anonimizar.py archivo.csv     # procesa solo ese archivo (resultado en output/)

Opciones:
    --idioma es|en
    --modo reemplazar|mascara|hash|borrar
    --seudonimos                 mismo valor -> mismo sustituto (PERSON_1, EMAIL_ADDRESS_3...)
    --umbral 0.35                confianza mínima (0-1)
    --columnas a,b,c             (CSV/XLSX) solo analiza estas columnas
    --ocultar-columnas a,b       (CSV/XLSX) sustituye la columna entera sin analizarla

Formatos: .txt .md .csv .tsv .json .log .docx .xlsx .pdf
El PDF se exporta como .txt, .pdf y .md. En el PDF se conserva la página original
y se tapan los datos detectados. El .md ordena el texto en cláusulas.
Si la página no tiene texto (escaneo), se lee con el OCR en español de Windows.
"""
import argparse
import csv
import difflib
import hashlib
import io
import re
import sys

import numpy as np
from collections import defaultdict
from pathlib import Path

import pymupdf
import winocr
from PIL import Image, ImageDraw, ImageFont
from presidio_analyzer import AnalyzerEngine, BatchAnalyzerEngine, Pattern, PatternRecognizer, RecognizerResult
from presidio_analyzer.nlp_engine import NlpEngineProvider
from presidio_anonymizer import AnonymizerEngine
from presidio_anonymizer.entities import OperatorConfig

BASE = Path(__file__).parent
INPUT, OUTPUT = BASE / "input", BASE / "output"
MODELOS = {"es": "es_core_news_lg", "en": "en_core_web_lg"}
TEXTO_PLANO = {".txt", ".md", ".json", ".log", ".xml", ".html"}
TABLAS = {".csv", ".tsv"}
MAX_BLOQUE = 100_000  # caracteres por bloque de texto (spaCy admite 1M como máximo)
FILAS_LOTE = 5_000


class Anonimizador:
    def __init__(self, idioma, modo, umbral, seudonimos):
        provider = NlpEngineProvider(nlp_configuration={
            "nlp_engine_name": "spacy",
            "models": [{"lang_code": k, "model_name": v} for k, v in MODELOS.items()],
        })
        self.analyzer = AnalyzerEngine(nlp_engine=provider.create_engine(),
                                       supported_languages=list(MODELOS))
        for reconocedor in reconocedores_argentina():
            self.analyzer.registry.add_recognizer(reconocedor)
        self.batch = BatchAnalyzerEngine(self.analyzer)
        self.anonymizer = AnonymizerEngine()
        self.idioma, self.modo, self.umbral, self.seudonimos = idioma, modo, umbral, seudonimos
        self.mapa = defaultdict(dict)  # tipo -> {valor original: seudónimo}
        self.conteo = defaultdict(int)
        tipos = self.analyzer.get_supported_entities(idioma)
        self.ops = {t: self._operador(t) for t in tipos} if seudonimos else self._operadores_modo()

    def _operadores_modo(self):
        if self.modo == "mascara":
            op = OperatorConfig("mask", {"masking_char": "*", "chars_to_mask": 1000, "from_end": False})
        elif self.modo == "hash":
            op = OperatorConfig("hash", {"hash_type": "sha256"})
        elif self.modo == "borrar":
            op = OperatorConfig("redact", {})
        else:
            return None  # por defecto Presidio reemplaza por <TIPO_ENTIDAD>
        return {"DEFAULT": op}

    def _operador(self, tipo):
        return OperatorConfig("custom", {"lambda": lambda valor: self.seudonimo(tipo, valor)})

    def seudonimo(self, tipo, valor):
        tabla = self.mapa[tipo]
        if valor not in tabla:
            tabla[valor] = f"{tipo}_{len(tabla) + 1}"
        return tabla[valor]

    def ocultar(self, tipo, valor):
        """Sustituye un valor completo sin analizarlo (columnas marcadas como sensibles)."""
        if not valor or not valor.strip():
            return valor
        self.conteo[tipo] += 1
        if self.seudonimos:
            return self.seudonimo(tipo, valor)
        if self.modo == "mascara":
            return "*" * len(valor)
        if self.modo == "borrar":
            return ""
        if self.modo == "hash":
            return hashlib.sha256(valor.encode()).hexdigest()
        return f"<{tipo}>"

    def etiqueta(self, tipo, valor):
        """Texto que reemplaza un dato, sin contarlo (el conteo lo hace _aplicar)."""
        if self.seudonimos:
            return self.seudonimo(tipo, valor)
        if self.modo == "mascara":
            return "*" * len(valor)
        if self.modo == "borrar":
            return ""
        if self.modo == "hash":
            return hashlib.sha256(valor.encode()).hexdigest()[:16]
        return f"<{tipo}>"

    def _aplicar(self, texto, resultados):
        for r in resultados:
            self.conteo[r.entity_type] += 1
        return self.anonymizer.anonymize(text=texto, analyzer_results=resultados, operators=self.ops).text

    def texto(self, texto):
        """Anonimiza un texto de cualquier longitud, dividiéndolo en bloques por líneas."""
        if not texto or not texto.strip():
            return texto
        bloques, actual = [], ""
        for linea in texto.splitlines(keepends=True):
            if actual and len(actual) + len(linea) > MAX_BLOQUE:
                bloques.append(actual)
                actual = ""
            actual += linea
        bloques.append(actual)
        return "".join(self._aplicar(b, self.analyzer.analyze(text=b, language=self.idioma,
                                                               score_threshold=self.umbral))
                       for b in bloques)

    def valores(self, valores, contexto, cache):
        """Anonimiza en lote los valores de una columna; usa la cabecera como contexto."""
        pendientes = [v for v in dict.fromkeys(valores) if v not in cache and v and v.strip()]
        if pendientes:
            analisis = self.batch.analyze_iterator(pendientes, language=self.idioma, batch_size=256,
                                                   score_threshold=self.umbral, context=[contexto])
            for v, res in zip(pendientes, analisis):
                cache[v] = self._aplicar(v, res)
        return [cache.get(v, v) for v in valores]


def lista(texto):
    return {c.strip().lower() for c in texto.split(",") if c.strip()} if texto else set()


def anonimizar_filas(filas, cabecera, anon, columnas, ocultas, caches):
    """Procesa un lote de filas columna a columna."""
    for i, nombre in enumerate(cabecera):
        clave = nombre.strip().lower()
        col = [f[i] if i < len(f) else "" for f in filas]
        if clave in ocultas:
            nuevos = [anon.ocultar(nombre.strip().upper().replace(" ", "_"), v) for v in col]
        elif not columnas or clave in columnas:
            nuevos = anon.valores(col, nombre, caches[i])
        else:
            continue
        for f, v in zip(filas, nuevos):
            if i < len(f):
                f[i] = v


def leer_texto(ruta):
    for cod in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return ruta.read_text(encoding=cod)
        except UnicodeDecodeError:
            continue


def procesar_csv(entrada, salida, anon, columnas, ocultas):
    contenido = leer_texto(entrada)
    try:
        dialecto = csv.Sniffer().sniff(contenido[:20_000], delimiters=",;\t|")
    except csv.Error:
        dialecto = csv.excel_tab if entrada.suffix.lower() == ".tsv" else csv.excel
    lector = csv.reader(contenido.splitlines(), dialecto)
    cabecera = next(lector)
    faltan = (columnas | ocultas) - {c.strip().lower() for c in cabecera}
    if faltan:
        print(f"  Aviso: columnas no encontradas: {', '.join(sorted(faltan))}")
    caches = defaultdict(dict)
    total = 0
    with open(salida, "w", encoding="utf-8-sig", newline="") as f:
        escritor = csv.writer(f, delimiter=dialecto.delimiter, quotechar=dialecto.quotechar or '"')
        escritor.writerow(cabecera)
        lote = []
        for fila in lector:
            lote.append(fila)
            if len(lote) >= FILAS_LOTE:
                anonimizar_filas(lote, cabecera, anon, columnas, ocultas, caches)
                escritor.writerows(lote)
                total += len(lote)
                print(f"  {total} filas...", flush=True)
                lote = []
        if lote:
            anonimizar_filas(lote, cabecera, anon, columnas, ocultas, caches)
            escritor.writerows(lote)
            total += len(lote)
    print(f"  {total} filas procesadas")


def procesar_xlsx(entrada, salida, anon, columnas, ocultas):
    from openpyxl import load_workbook
    wb = load_workbook(entrada)
    for ws in wb.worksheets:
        filas = list(ws.iter_rows())
        if not filas:
            continue
        cabecera = [str(c.value or "") for c in filas[0]]
        for i, nombre in enumerate(cabecera):
            clave = nombre.strip().lower()
            celdas = [f[i] for f in filas[1:] if i < len(f)
                      and isinstance(f[i].value, str) and not f[i].value.startswith("=")]
            if clave in ocultas:
                for c in celdas:
                    c.value = anon.ocultar(nombre.strip().upper().replace(" ", "_"), c.value)
            elif not columnas or clave in columnas:
                for c, v in zip(celdas, anon.valores([c.value for c in celdas], nombre, {})):
                    c.value = v
    wb.save(salida)


def parrafo_docx(p, anon):
    original = p.text
    nuevo = anon.texto(original)
    if nuevo != original and p.runs:
        p.runs[0].text = nuevo
        for run in p.runs[1:]:
            run.text = ""


def procesar_docx(entrada, salida, anon):
    from docx import Document
    doc = Document(entrada)

    def bloques(contenedor):
        yield from contenedor.paragraphs
        for tabla in contenedor.tables:
            for fila in tabla.rows:
                for celda in fila.cells:
                    yield from bloques(celda)

    for p in bloques(doc):
        parrafo_docx(p, anon)
    for seccion in doc.sections:
        for parte in (seccion.header, seccion.footer):
            for p in bloques(parte):
                parrafo_docx(p, anon)
    doc.save(salida)


def reconocedores_argentina():
    """DNI y CUIT/CUIL, que Presidio no trae para Argentina."""
    return [
        PatternRecognizer(
            supported_entity="AR_DNI",
            supported_language="es",
            patterns=[
                Pattern("dni_puntos", r"\b\d{1,2}\.\d{3}\.\d{3}\b", 0.2),
                Pattern("dni_frase", r"(?i)identidad\s+\d[\d.\s,]{5,18}\d", 0.8),
            ],
            context=["dni", "documento", "identidad"],
        ),
        PatternRecognizer(
            supported_entity="AR_CUIT",
            supported_language="es",
            patterns=[
                Pattern("cuit", r"\b\d{2}-\d{8}-\d\b", 0.9),
                Pattern("cuit_frase", r"(?i)cui[tl]\s*\d{2}\D+\d{8}\D+\d", 0.85),
            ],
            context=["cuit", "cuil"],
        ),
    ]


def letras(token):
    return "".join(c for c in token if c.isalpha()).upper()


PALABRAS_COMUNES = {
    "escritura", "sociedad", "limitada", "gerente", "gerentes", "cuotas", "capital",
    "articulo", "contrato", "escribano", "inscripcion", "registro", "comercio",
    "responsabilidad", "constitucion", "trescientos", "administracion", "suscribe",
    "comparecen", "domiciliado", "domiciliada", "argentino", "argentina", "soltera",
    "casado", "abogado", "procuradora", "publico", "notarial", "primero", "segundo",
    "tercero", "cuarto", "quinto", "sexto", "septimo", "octavo", "noveno", "decimo",
}


def en_mayusculas(token):
    letras_orig = [c for c in token if c.isalpha()]
    return bool(letras_orig) and sum(c.isupper() for c in letras_orig) >= len(letras_orig) * 0.8


def sufijo_comun(a, b):
    n = 0
    for x, y in zip(reversed(a), reversed(b)):
        if x != y:
            break
        n += 1
    return n


def parecido(palabra, semillas):
    """Apellido en mayúsculas deformado por el OCR respecto de uno ya visto."""
    limpia = letras(palabra)
    if len(limpia) < 7 or limpia.lower() in PALABRAS_COMUNES or not en_mayusculas(palabra):
        return False
    return any(
        difflib.SequenceMatcher(None, limpia, semilla).ratio() >= 0.64 and sufijo_comun(limpia, semilla) >= 3
        for semilla in semillas
    )


def ampliar_con_parecidos(texto, palabras, resultados, semillas):
    cubierto = [False] * len(palabras)
    for resultado in resultados:
        if resultado.entity_type != "PERSON":
            continue
        for token in texto[resultado.start:resultado.end].split():
            limpia = letras(token)
            if len(limpia) >= 7 and en_mayusculas(token) and limpia.lower() not in PALABRAS_COMUNES:
                semillas.append(limpia)
        for i, palabra in enumerate(palabras):
            if palabra["end"] > resultado.start and palabra["start"] < resultado.end:
                cubierto[i] = True
    extra = []
    for resultado in resultados:
        if resultado.entity_type != "PERSON":
            continue
        siguientes = [i for i, palabra in enumerate(palabras) if palabra["start"] >= resultado.end]
        if not siguientes or cubierto[siguientes[0]]:
            continue
        palabra = palabras[siguientes[0]]
        limpia = letras(palabra["text"])
        if (
            len(limpia) < 7
            or palabra["text"].endswith(":")
            or not en_mayusculas(palabra["text"])
            or limpia.lower() in PALABRAS_COMUNES
        ):
            continue
        extra.append(RecognizerResult("PERSON", palabra["start"], palabra["end"], 0.9))
        semillas.append(limpia)
        cubierto[siguientes[0]] = True
        print(f"    apellido junto al nombre: {palabra['text']}", flush=True)
    cambio = True
    while cambio and semillas:
        cambio = False
        for i, palabra in enumerate(palabras):
            if cubierto[i] or not parecido(palabra["text"], semillas):
                continue
            extra.append(RecognizerResult("PERSON", palabra["start"], palabra["end"], 0.9))
            semillas.append(letras(palabra["text"]))
            cubierto[i] = True
            cambio = True
            print(f"    apellido parecido: {palabra['text']}", flush=True)
    return extra


def sin_solapes(resultados):
    """Si dos detecciones se pisan, conserva el tramo que sigue después de la primera."""
    aceptados, fin = [], -1
    for resultado in sorted(resultados, key=lambda r: (r.start, -(r.end - r.start))):
        inicio = max(resultado.start, fin)
        if inicio >= resultado.end:
            continue
        if inicio != resultado.start:
            resultado = RecognizerResult(resultado.entity_type, inicio, resultado.end, resultado.score)
        aceptados.append(resultado)
        fin = resultado.end
    return aceptados


def quitar_subrayado(imagen):
    """Saca líneas horizontales largas para que el OCR lea la palabra subrayada."""
    arreglo = np.array(imagen)
    oscuro = np.array(imagen.convert("L")) < 90
    for y in range(oscuro.shape[0]):
        fila = np.concatenate(([False], oscuro[y], [False]))
        cambios = np.diff(fila.astype(np.int8))
        for inicio, fin in zip(np.where(cambios == 1)[0], np.where(cambios == -1)[0]):
            if fin - inicio > 80:
                arreglo[y, inicio:fin] = 255
    return Image.fromarray(arreglo)


def ocr_imagen(imagen):
    try:
        return winocr.recognize_pil_sync(imagen, "es-MX")
    except AssertionError:
        sys.exit("Falta el OCR en español de Windows (paquete de idioma es-MX).")


def palabras_de_ocr(resultado):
    palabras, partes, pos = [], [], 0
    for i, linea in enumerate(resultado.get("lines") or []):
        if i:
            partes.append("\n")
            pos += 1
        for j, word in enumerate(linea.get("words") or []):
            if j:
                partes.append(" ")
                pos += 1
            texto = word.get("text") or ""
            caja = word.get("bounding_rect") or {}
            palabras.append({
                "text": texto,
                "start": pos,
                "end": pos + len(texto),
                "box": {
                    "x": float(caja.get("x", 0)),
                    "y": float(caja.get("y", 0)),
                    "width": float(caja.get("width", 0)),
                    "height": float(caja.get("height", 0)),
                },
            })
            partes.append(texto)
            pos += len(texto)
    return "".join(partes), palabras


def palabras_nativas(pagina, escala):
    palabras, partes, pos = [], [], 0
    linea_anterior = None
    for x0, y0, x1, y1, word, _bloque, linea, _n in pagina.get_text("words"):
        if palabras and linea != linea_anterior:
            partes.append("\n")
            pos += 1
        elif palabras:
            partes.append(" ")
            pos += 1
        linea_anterior = linea
        palabras.append({
            "text": word,
            "start": pos,
            "end": pos + len(word),
            "box": {
                "x": x0 * escala,
                "y": y0 * escala,
                "width": (x1 - x0) * escala,
                "height": (y1 - y0) * escala,
            },
        })
        partes.append(word)
        pos += len(word)
    return "".join(partes), palabras


def contenido_pagina(pagina):
    """Imagen de la página, texto y cajas de cada palabra. El tercer valor indica si hubo OCR."""
    pix = pagina.get_pixmap(dpi=200)
    imagen = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    if len((pagina.get_text("text") or "").strip()) >= 40 and pagina.get_text("words"):
        escala = pix.width / float(pagina.rect.width)
        texto, palabras = palabras_nativas(pagina, escala)
        return imagen, texto, palabras, False
    texto, palabras = palabras_de_ocr(ocr_imagen(quitar_subrayado(imagen)))
    return imagen, texto, palabras, True


def misma_linea(a, b):
    ay, ah = a["box"]["y"], a["box"]["height"]
    by, bh = b["box"]["y"], b["box"]["height"]
    arriba = max(ay, by)
    abajo = min(ay + ah, by + bh)
    return abajo - arriba > 0.4 * min(ah, bh)


def grupos_de_linea(seleccion):
    grupos, actual = [], [seleccion[0]]
    for palabra in seleccion[1:]:
        if misma_linea(actual[-1], palabra):
            actual.append(palabra)
        else:
            grupos.append(actual)
            actual = [palabra]
    grupos.append(actual)
    return grupos


def fuente(tamaño, cache):
    if tamaño in cache:
        return cache[tamaño]
    for ruta in (r"C:\Windows\Fonts\times.ttf", r"C:\Windows\Fonts\arial.ttf"):
        if Path(ruta).exists():
            cache[tamaño] = ImageFont.truetype(ruta, tamaño)
            return cache[tamaño]
    cache[tamaño] = ImageFont.load_default()
    return cache[tamaño]


def color_papel(imagen):
    ancho, alto = imagen.size
    muestras = [(8, 8), (ancho - 8, 8), (8, alto - 8), (ancho // 2, 8)]
    return max((imagen.getpixel(p) for p in muestras), key=lambda c: sum(c[:3]))


def tapar(dibujo, palabras, resultado, etiqueta, papel, cache):
    seleccion = [p for p in palabras if p["end"] > resultado.start and p["start"] < resultado.end and p["box"]["width"] > 0]
    if not seleccion or len(seleccion) > 40:
        return
    for i, grupo in enumerate(grupos_de_linea(seleccion)):
        x0 = min(p["box"]["x"] for p in grupo) - 2
        y0 = min(p["box"]["y"] for p in grupo) - 1
        x1 = max(p["box"]["x"] + p["box"]["width"] for p in grupo) + 2
        y1 = max(p["box"]["y"] + p["box"]["height"] for p in grupo) + 1
        dibujo.rectangle([x0, y0, x1, y1], fill=papel)
        if i or not etiqueta:
            continue
        tamaño = max(8, int((y1 - y0) * 0.72))
        letra = fuente(tamaño, cache)
        while tamaño > 8 and dibujo.textlength(etiqueta, font=letra) > (x1 - x0 - 4):
            tamaño -= 1
            letra = fuente(tamaño, cache)
        caja = dibujo.textbbox((0, 0), etiqueta, font=letra)
        alto = caja[3] - caja[1]
        dibujo.text((x0 + 2, y0 + max(0, (y1 - y0 - alto) / 2) - caja[1]), etiqueta, fill="black", font=letra)


def analizar_pagina(texto, palabras, anon, semillas):
    if not texto.strip():
        return [], ""
    resultados = anon.analyzer.analyze(text=texto, language=anon.idioma, score_threshold=anon.umbral)
    resultados.extend(ampliar_con_parecidos(texto, palabras, resultados, semillas))
    resultados = sin_solapes(resultados)
    return resultados, anon._aplicar(texto, resultados)


def es_ruido(linea):
    """Margen, sello y numeración de la foja notarial, que no son parte del contrato."""
    if re.fullmatch(r"A\s+<PHONE_NUMBER>", linea):
        return True
    if re.search(r"tBR|BROS del|R\.?\s*863|Q:s65|^OE\b|^Esce|^DE Es", linea):
        return True
    if re.match(r"Dos mil trescientos", linea, re.I):
        return True
    if re.fullmatch(r"[a-nA-N]\)", linea):
        return False
    if re.fullmatch(r"del\s+[A-Z]", linea):
        return True
    if "<" in linea:
        return False
    letras = sum(c.isalpha() for c in linea)
    if letras == 0:
        return True
    if letras < 4 and not re.fullmatch(r"[a-záéíóúñ]{1,3}", linea, re.I):
        return True
    compacto = re.sub(r"\s", "", linea)
    if compacto and letras * 2 < len(compacto) and len(linea) < 24:
        return True
    return False


def texto_continuo(paginas):
    lineas = []
    for pagina in paginas:
        for linea in pagina.splitlines():
            linea = linea.strip()
            if linea and not es_ruido(linea):
                lineas.append(linea)
    texto = re.sub(r"\s+", " ", " ".join(lineas)).strip()
    texto = re.sub(r"\s+([,.;:)])", r"\1", texto)
    return re.sub(r"([¿¡(])\s+", r"\1", texto)


def incisos(texto):
    texto = re.sub(r"\s+I\)\s+", " l) ", texto)
    partes = re.split(r"\s+([a-n])\)\s+", texto, flags=re.I)
    if len(partes) < 7:
        return texto
    items = [f"- **{letra.lower()})** {contenido.strip()}" for letra, contenido in zip(partes[1::2], partes[2::2])]
    return partes[0].strip() + "\n\n" + "\n".join(items)


def secciones(texto):
    marcas = (
        ("la sociedad se denomina", "Denominación y duración"),
        ("tiene por objeto", "Objeto social"),
        ("capital social es", "Capital social"),
        ("La administración y representación", "Administración y representación"),
        ("gerente convocará", "Reuniones de socios"),
        ("SEPTIMIO", "Séptimo. Ejercicio social"),
        ("OCTAVO", "Octavo. Cesión de cuotas"),
        ("NOVENO", "Noveno. Fallecimiento o incapacidad"),
        ("DECIMO", "Décimo. Disolución y liquidación"),
        ("SEGUNDO:", "Suscripción, domicilio y gerencia"),
        ("escritura matriz", "Certificación de la copia"),
    )
    encontrados, desde = [], 0
    for clave, titulo in marcas:
        idx = texto.lower().find(clave.lower(), desde)
        if idx < 0:
            continue
        encontrados.append((idx, titulo))
        desde = idx + len(clave)
    if not encontrados:
        return [("Texto", texto)]
    bloques = []
    if encontrados[0][0] > 40:
        bloques.append(("Comparecencia", texto[:encontrados[0][0]].strip()))
    for i, (idx, titulo) in enumerate(encontrados):
        fin = encontrados[i + 1][0] if i + 1 < len(encontrados) else len(texto)
        bloques.append((titulo, texto[idx:fin].strip()))
    return bloques


def paginas_a_markdown(paginas, nombre):
    """Arma un Markdown continuo, con las cláusulas en el orden de la escritura."""
    bloques = secciones(texto_continuo(paginas))
    partes = [
        "---",
        "documento: escritura de constitución de sociedad",
        "uso: lectura del texto anonimizado",
        "anonimizado: true",
        "origen: OCR de un PDF escaneado",
        "etiquetas: <PERSON> persona; <LOCATION> lugar; <ORGANIZATION> organización; <AR_DNI> documento; <AR_CUIT> CUIT o CUIL; <PHONE_NUMBER> teléfono; <URL> web",
        "---",
        "",
        f"# {nombre}",
        "",
        "Texto anonimizado, en el orden de la escritura. Las etiquetas sustituyen datos personales. El escaneo dejó errores de lectura; no hay que reconstruir lo tapado.",
        "",
    ]
    for titulo, cuerpo in bloques:
        partes.append(f"## {titulo}")
        partes.append("")
        partes.append(incisos(cuerpo) if titulo == "Objeto social" else cuerpo)
        partes.append("")
    return "\n".join(partes).strip() + "\n"


def procesar_pdf(entrada, salida, anon):
    origen = pymupdf.open(entrada)
    destino = pymupdf.open()
    textos = []
    apellidos = []
    for n, pagina in enumerate(origen, start=1):
        imagen, crudo, palabras, uso_ocr = contenido_pagina(pagina)
        if uso_ocr:
            print(f"  página {n}: OCR ({len(crudo)} caracteres)", flush=True)
        resultados, anonimizado = analizar_pagina(crudo, palabras, anon, apellidos)
        textos.append(anonimizado)
        dibujo = ImageDraw.Draw(imagen)
        papel = color_papel(imagen)
        cache = {}
        for resultado in resultados:
            valor = crudo[resultado.start:resultado.end]
            tapar(dibujo, palabras, resultado, anon.etiqueta(resultado.entity_type, valor), papel, cache)
        buffer = io.BytesIO()
        imagen.save(buffer, format="JPEG", quality=80)
        nueva = destino.new_page(width=pagina.rect.width, height=pagina.rect.height)
        nueva.insert_image(nueva.rect, stream=buffer.getvalue())
    salida_pdf = salida.with_suffix(".pdf")
    destino.save(salida_pdf)
    destino.close()
    origen.close()
    salida.write_text("\n\n--- página ---\n\n".join(textos), encoding="utf-8")
    salida_md = salida.with_suffix(".md")
    salida_md.write_text(paginas_a_markdown(textos, "Escritura de constitución"), encoding="utf-8")
    print(f"  -> {salida_pdf}")
    print(f"  -> {salida_md}")


def procesar(entrada, salida, anon, args):
    ext = entrada.suffix.lower()
    if ext == ".pdf":
        salida = salida.with_suffix(".txt")
    salida.parent.mkdir(parents=True, exist_ok=True)
    if ext in TABLAS:
        procesar_csv(entrada, salida, anon, args.columnas, args.ocultar_columnas)
    elif ext == ".xlsx":
        procesar_xlsx(entrada, salida, anon, args.columnas, args.ocultar_columnas)
    elif ext in TEXTO_PLANO:
        salida.write_text(anon.texto(leer_texto(entrada)), encoding="utf-8")
    elif ext == ".docx":
        procesar_docx(entrada, salida, anon)
    elif ext == ".pdf":
        procesar_pdf(entrada, salida, anon)
    else:
        print(f"  Formato no soportado, se omite: {entrada.name}")
        return
    print(f"  -> {salida}")


def main():
    ap = argparse.ArgumentParser(description="Anonimiza documentos con Presidio")
    ap.add_argument("archivo", type=Path, nargs="?", help="si se omite, procesa todo input/")
    ap.add_argument("--idioma", choices=list(MODELOS), default="es")
    ap.add_argument("--modo", choices=["reemplazar", "mascara", "hash", "borrar"], default="reemplazar")
    ap.add_argument("--seudonimos", action="store_true")
    ap.add_argument("--umbral", type=float, default=0.35)
    ap.add_argument("--columnas", type=lista, default=set())
    ap.add_argument("--ocultar-columnas", type=lista, default=set())
    args = ap.parse_args()

    if args.archivo:
        if not args.archivo.exists():
            sys.exit(f"No existe: {args.archivo}")
        trabajos = [(args.archivo, OUTPUT / args.archivo.name)]
    else:
        INPUT.mkdir(exist_ok=True)
        trabajos = [(p, OUTPUT / p.relative_to(INPUT)) for p in sorted(INPUT.rglob("*")) if p.is_file()]
        if not trabajos:
            sys.exit(f"No hay archivos en {INPUT}")

    print("Cargando modelos...")
    anon = Anonimizador(args.idioma, args.modo, args.umbral, args.seudonimos)
    for entrada, salida in trabajos:
        print(f"Procesando {entrada.name}")
        try:
            procesar(entrada, salida, anon, args)
        except Exception as e:
            print(f"  ERROR: {e}")

    print("\nEntidades anonimizadas:")
    for tipo, n in sorted(anon.conteo.items(), key=lambda x: -x[1]):
        print(f"  {tipo}: {n}")


if __name__ == "__main__":
    main()
