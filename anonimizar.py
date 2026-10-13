"""Anonimiza documentos con Microsoft Presidio.

Uso:
    python anonimizar.py                 # procesa todo input/ y deja el resultado en output/
    python anonimizar.py archivo.pdf     # procesa solo ese archivo (resultado en output/)
    python anonimizar.py --desanonimizar resp_ia.txt --mapa output/archivo.mapa.json

Opciones:
    --idioma es|en
    --modo reemplazar|mascara|hash|borrar
    --seudonimos                 numera cada entidad (<PERSONA_1>, <DNI_1>...) igual en todo el documento y guarda .mapa.json
    --desanonimizar archivo.txt  restaura los datos reales en un texto usando un mapa de equivalencias
    --mapa archivo.mapa.json     archivo con el glosario de correspondencias para --desanonimizar
    --sellos                     (PDF) tapa los sellos del escribano en hojas de protocolo notarial
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
import json
import re
import sys
import unicodedata

import numpy as np
from collections import defaultdict
from pathlib import Path

import pymupdf
import winocr
from PIL import Image, ImageDraw, ImageFont
from presidio_analyzer import AnalyzerEngine, BatchAnalyzerEngine, Pattern, PatternRecognizer, RecognizerResult
from presidio_analyzer.nlp_engine import NlpEngineProvider

BASE = Path(__file__).parent
INPUT, OUTPUT = BASE / "input", BASE / "output"
MODELOS = {"es": "es_core_news_lg", "en": "en_core_web_lg"}
TEXTO_PLANO = {".txt", ".md", ".json", ".log", ".xml", ".html"}
TABLAS = {".csv", ".tsv"}
MAX_BLOQUE = 100_000
FILAS_LOTE = 5_000

ETIQUETAS_ES = {
    "PERSON": "PERSONA",
    "LOCATION": "LUGAR",
    "DIRECCION": "DIRECCION",
    "ORGANIZATION": "ORGANIZACION",
    "AR_DNI": "DNI",
    "AR_CUIT": "CUIT",
    "EMAIL_ADDRESS": "EMAIL",
    "PHONE_NUMBER": "TELEFONO",
    "DATE_TIME": "FECHA",
    "IBAN_CODE": "CBU",
    "IP_ADDRESS": "IP",
    "URL": "WEB",
    "NRP": "DOCUMENTO",
}


# ==============================================================================
# STOPWORDS Y REGLAS DE CALIBRACIÓN JURÍDICA
# ==============================================================================

STOPWORDS_ESTRUCTURA = {
    "apartado", "clausula", "articulo", "inciso", "capitulo", "titulo", "seccion", "anexo",
    "primero", "segundo", "tercero", "cuarto", "quinto", "sexto", "septimo", "octavo", "noveno", "decimo",
    "undecimo", "duodecimo", "primera", "segunda", "tercera", "cuarta", "quinta", "sexta", "septima", "octava",
    "novena", "decima", "pri", "ero", "priero", "ercero", "quinto", "cuarto", "segundo", "tercero",
    "cuerda", "concuerda", "matriz", "folio", "foja", "protocolo", "testimonio", "notarial",
    "escritura", "numero", "comparecen", "compareciente", "comparecientes", "otorgante", "otorgantes",
    "suscribe", "suscriben", "objeto", "denominacion", "duracion", "domicilio", "sede", "capital",
    "cuotas", "cuota", "administracion", "representacion", "fiscalizacion", "gobierno", "gerencia",
    "gerente", "gerentes", "socio", "socios", "socia", "socias", "asamblea", "reunion", "reuniones",
    "ejercicio", "balance", "utilidades", "reservas", "disolucion", "liquidacion", "liquidador",
    "liquidadores", "cesion", "preferencia", "fallecimiento", "incapacidad", "herederos", "sucesion",
    "incorporacion", "extinguido", "disuelto", "liquidado", "aprobado", "inscripto", "suscripto",
    "integrado", "legal", "especial", "real", "constituido", "social", "quienes", "quien",
    "cuales", "cual", "mismo", "misma", "dicho", "dicha", "presente", "siguiente", "titular",
}

STOPWORDS_TIPO_SOCIEDAD = {
    "sociedad", "sociedades", "responsabilidad", "limitada", "anonima", "acciones",
    "s.r.l.", "s.a.", "s.a.s.", "srl", "sa", "sas", "constitucion", "estatuto", "contrato",
    "sociedad de responsabilidad limitada", "sociedad anonima",
}

ORGANISMOS_PUBLICOS_LEYES = {
    "ley", "leyes", "la ley", "decreto", "resolucion", "codigo", "codigo civil", "codigo comercial",
    "codigo civil y comercial", "inspeccion general de justicia", "igj", "registro publico de comercio",
    "registro publico", "banco de la nacion argentina", "banco nacion", "banco", "nacion argentina",
    "nacion argentinas", "afip", "dgi", "anses", "arca", "poder judicial", "justicia",
    "estado nacional", "republica argentina", "boletin oficial",
}

TITULOS_PROFESIONES_ESTADO = {
    "escribano", "escribana", "notario", "notaria", "escribano publico", "escribana publica", "publico", "publica",
    "abogado", "abogada", "procurador", "procuradora", "contador", "contadora", "licenciado", "licenciada",
    "ingeniero", "ingeniera", "comerciante", "empleado", "empleada", "dr", "dra", "lic", "ing",
    "senor", "senora", "sr", "sra", "don", "dona", "presidente", "vicepresidente", "director", "directora",
    "casado", "casada", "soltero", "soltera", "divorciado", "divorciada", "viudo", "viuda",
    "argentino", "argentina", "argentinos", "argentinas", "argent inor", "nacido", "nacida",
}

TERMINOS_DOCUMENTOS = {
    "documento", "documento nacional de identidad", "dni", "le", "lc", "ci", "pasaporte",
    "identidad", "cuit", "cuil", "cdi",
}

TERMINOS_GENERICOS_OBJETO = {
    "call center", "software", "hardware", "marketing", "turnos", "videollamadas",
    "procesamiento", "datos", "servicios", "asesoramiento", "consultoria", "internet", "web",
}

LUGARES_GENERICOS = {
    "ciudad", "provincia", "partido", "municipio", "republica", "pais", "nacion", "localidad",
}

RUIDO_SELLOS = {
    "oe esc", "oe esc 10 t", "ve esc", "ve esc q08941", "t bros", "r.863", "q:s65", "aro", "ribano",
}

CONECTORES_Y_ARTICULOS = {
    "el", "la", "los", "las", "lo", "de", "del", "y", "e", "a", "al", "en", "con", "por", "para", "ei",
}

TODAS_STOPWORDS = (
    STOPWORDS_ESTRUCTURA | STOPWORDS_TIPO_SOCIEDAD | ORGANISMOS_PUBLICOS_LEYES |
    TITULOS_PROFESIONES_ESTADO | TERMINOS_DOCUMENTOS | TERMINOS_GENERICOS_OBJETO |
    LUGARES_GENERICOS | RUIDO_SELLOS | CONECTORES_Y_ARTICULOS
)

BORDES_RECORTE = CONECTORES_Y_ARTICULOS | {
    "senor", "senora", "don", "dona", "escribano", "escribana", "publico", "publica",
    "ante", "mi", "que", "se", "su", "sus", "como", "titular", "socio", "socia", "gerente",
}

PREFIJOS_DATOS = {
    "DIRECCION": re.compile(
        r"(?i)^(?:domiciliad[oa]s?\s+en|(?:(?:con|fija(?:ndo)?|se\s+fija\s+el)\s+)?domicilio\s+(?:real|legal|especial|constituido|social)?\s*en|sit[oa]s?\s+en|(?:con\s+)?sede\s+(?:social\s+)?en)\s*"
    ),
    "AR_DNI": re.compile(r"(?i)^(?:identidad|d\.?\s*n\.?\s*i\.?|documento\s+nacional\s+de\s+identidad)\s*(?:n[°º.]?\s*)?"),
    "AR_CUIT": re.compile(r"(?i)^c\.?\s*u\.?\s*i\.?\s*[tl]\.?\s*(?:n[°º.]?\s*)?"),
    "ORGANIZATION": re.compile(r'(?i)^(?:(?:se\s+)?denomina\s*|nombre\s+de\s+la\s+sociedad\s*|sociedad\s*)["\']?\s*'),
}


def norm(texto):
    t = texto.strip().strip(",.;:()[]{}¿?¡!\"'/-")
    t = re.sub(r"\s+", " ", t).lower()
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")


# Tipos numéricos: se comparan solo por sus dígitos, así "29.494 786" y "29494786" son el mismo dato.
NUMERICOS = {"AR_DNI", "AR_CUIT", "PHONE_NUMBER", "IBAN_CODE"}


def clave(tipo, valor):
    """Forma canónica de un valor para decidir si dos menciones son la misma entidad."""
    if tipo in NUMERICOS:
        digitos = re.sub(r"\D", "", valor)
        if tipo == "AR_DNI":
            return digitos.lstrip("0")
        if tipo == "PHONE_NUMBER":
            return digitos[-8:]  # ignora prefijos +54, 011, 15
        return digitos
    return norm(limpiar_direccion_val(valor))


def sim(a, b):
    return difflib.SequenceMatcher(None, norm(a), norm(b)).ratio()


# Espacio entre palabras, tolerando la numeración de margen notarial que el OCR intercala entre ellas
SEPARADOR_OCR = r"\s+(?:\d{1,2}\s+)?"

VOCALES = {"a": "aáà", "e": "eéè", "i": "iíì", "o": "oóò", "u": "uúüù"}


def patron_palabra(palabra):
    """Regex de una palabra que ignora tildes (Pérez = Perez)."""
    partes = []
    for c in palabra:
        base = norm(c) if c.isalpha() else ""
        partes.append(f"[{VOCALES[base]}]" if base in VOCALES else re.escape(c))
    return "".join(partes)


def limpiar_direccion_val(v):
    # Elimina saltos de línea producidos por numeración de margen notarial (ej. 'Lavalle\n27\n1290...')
    return re.sub(r"\n\s*\d{1,2}\s*\n", " ", v)


def es_falso_positivo(tipo, valor, texto_completo="", inicio=0, estricto=True):
    val_limpio = valor.strip().strip(",.;:()[]{}¿?¡!\"'/-")
    n = norm(val_limpio)
    if not n or len(n) < 2:
        return True
    if n in TODAS_STOPWORDS:
        return True

    words = [norm(w) for w in val_limpio.split() if norm(w)]
    if not words:
        return True
    if all(w in TODAS_STOPWORDS for w in words):
        return True

    if not any(c.isalpha() for c in val_limpio) and tipo not in {"AR_DNI", "AR_CUIT", "PHONE_NUMBER"}:
        return True

    # Comparación difusa contra vocabulario de estructura jurídica (ej. ERCERO, PRIERO)
    if len(words) == 1 and len(n) >= 4:
        for est in STOPWORDS_ESTRUCTURA:
            if abs(len(n) - len(est)) <= 1 and difflib.SequenceMatcher(None, n, est).ratio() >= 0.8:
                return True

    n_sin_espacios = n.replace(" ", "")
    if n_sin_espacios in STOPWORDS_ESTRUCTURA:
        return True

    if tipo == "PERSON":
        tokens_orig = [w for w in val_limpio.split() if w]
        if estricto and not any(w[0].isupper() for w in tokens_orig):
            return True
        if len(words) == 1:
            if n in STOPWORDS_ESTRUCTURA or n in TITULOS_PROFESIONES_ESTADO or n in LUGARES_GENERICOS:
                return True

    elif tipo == "ORGANIZATION":
        if len(words) == 1 and not re.search(r"\b(srl|sa|sas)\b", n):
            return True
        if n in STOPWORDS_TIPO_SOCIEDAD or n in ORGANISMOS_PUBLICOS_LEYES:
            return True

    elif tipo == "LOCATION":
        if n in {"la ley", "ley", "ciudad", "provincia", "banco", "segundo", "call center", "argentina"}:
            return True
        if len(words) == 1 and n in LUGARES_GENERICOS:
            return True

    elif tipo == "PHONE_NUMBER":
        digits = "".join(c for c in val_limpio if c.isdigit())
        previo = texto_completo[max(0, inicio - 30):inicio]
        # Formato de DNI (12.345.678) o número de foja notarial ("A 091331742")
        if re.fullmatch(r"\d{1,2}\.\d{3}\.\d{3}", val_limpio):
            return True
        if 8 <= len(digits) <= 10 and re.search(r"(?:^|\s)A\s*$", previo[-6:], re.I):
            return True
        # Una tira de dígitos sin separadores solo es teléfono si el texto lo dice
        if digits == val_limpio and not re.search(r"(?i)tel|cel|m[oó]vil|whats|contacto|llamar", previo):
            return True

    elif tipo == "URL":
        if not re.search(r"https?://|www\.|\.(com|net|org|ar|io)\b", val_limpio, re.I):
            return True

    return False


def recortar_bordes(texto, inicio, fin):
    while inicio < fin and texto[inicio] in " \t\r\n,.;:()[]{}¿?¡!\"'/-":
        inicio += 1
    while fin > inicio and texto[fin - 1] in " \t\r\n,.;:()[]{}¿?¡!\"'/-":
        fin -= 1

    tokens = [(m.start() + inicio, m.end() + inicio) for m in re.finditer(r"\S+", texto[inicio:fin])]
    while tokens and norm(texto[tokens[0][0]:tokens[0][1]]) in BORDES_RECORTE:
        tokens.pop(0)
    while tokens and norm(texto[tokens[-1][0]:tokens[-1][1]]) in BORDES_RECORTE:
        tokens.pop()

    if not tokens:
        return inicio, inicio
    nuevo_ini, nuevo_fin = tokens[0][0], tokens[-1][1]
    while nuevo_ini < nuevo_fin and texto[nuevo_ini] in " \t\r\n,.;:()[]{}¿?¡!\"'/-":
        nuevo_ini += 1
    while nuevo_fin > nuevo_ini and texto[nuevo_fin - 1] in " \t\r\n,.;:()[]{}¿?¡!\"'/-":
        nuevo_fin -= 1
    return nuevo_ini, nuevo_fin


def depurar_resultados(texto, resultados, estricto=True):
    limpios = []
    for r in resultados:
        tipo, inicio, fin = r.entity_type, r.start, r.end
        pref = PREFIJOS_DATOS.get(tipo)
        if pref:
            m = pref.match(texto[inicio:fin])
            if m:
                inicio += m.end()

        inicio, fin = recortar_bordes(texto, inicio, fin)
        if inicio >= fin:
            continue

        valor = texto[inicio:fin]
        if es_falso_positivo(tipo, valor, texto, inicio, estricto=estricto):
            continue

        if tipo == "LOCATION" and re.search(r"\b\d{1,5}\b", valor):
            tipo = "DIRECCION"

        # Si una LOCATION es una calle (viene después de "domicilio en", "calle"...) y continúa con altura/piso
        previo = norm(texto[max(0, inicio - 40):inicio])
        if tipo == "LOCATION" and re.search(r"\b(domicili\w*|sede|sit[oa]s?|calle|avenida|av)(\W+\w+){0,3}$", previo):
            resto = texto[fin:fin + 60]
            m_alt = re.match(r"(?:[ \t]+|(?:[ \t]*\n(?:\s*\d{1,2}\s*\n)?[ \t]*))(\d{1,5}\b(?:[ \t]*,?[ \t]*(?:piso|p\.|pb)\s*\d+)?(?:[ \t]*,?[ \t]*(?:dto|depto|oficina|of\.)\s*[a-z0-9]+)?)", resto, re.I)
            if m_alt:
                fin += m_alt.end()
                tipo = "DIRECCION"
                valor = texto[inicio:fin]

        # Si PERSON tiene solo nombres de pila y le sigue un apellido
        if tipo == "PERSON":
            resto = texto[fin:fin + 50]
            m_ape = re.match(r"(?:[ \t]+|(?:[ \t]*\n(?:\s*\d{1,2}\s*\n)?[ \t]*))([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ'’\-]{2,25}(?:\s+(?:de|del|la|los|de\s+la)\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ'’\-]+)?)\b", resto)
            if m_ape:
                ape_candidato = m_ape.group(1)
                n_ape = norm(ape_candidato)
                if n_ape not in TODAS_STOPWORDS and not any(w in TODAS_STOPWORDS for w in n_ape.split()):
                    fin += m_ape.end()
                    valor = texto[inicio:fin]

        limpios.append(RecognizerResult(tipo, inicio, fin, r.score))

    limpios.sort(key=lambda x: (x.start, -(x.end - x.start)))
    aceptados, pos_fin = [], -1
    for r in limpios:
        if r.start < pos_fin:
            continue
        aceptados.append(r)
        pos_fin = r.end

    return aceptados


def sin_solapes(resultados):
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


# ==============================================================================
# RECONOCEDORES ESPECÍFICOS DE ARGENTINA
# ==============================================================================

def reconocedores_argentina():
    """DNI, CUIT/CUIL, Direcciones y Sociedades en Argentina."""
    return [
        PatternRecognizer(
            supported_entity="AR_DNI",
            supported_language="es",
            patterns=[
                Pattern("dni_frase", r"(?i)(?:identidad|d\.?n\.?i\.?|documento\s+nacional\s+de\s+identidad)\s*(?:n[°º.]?\s*)?(\d{1,2}[\d.\s,]{5,12}\d)", 0.95),
            ],
            context=["dni", "documento", "identidad"],
        ),
        PatternRecognizer(
            supported_entity="AR_CUIT",
            supported_language="es",
            patterns=[
                Pattern("cuit", r"\b\d{2}[-\s]?\d{8}[-\s—–]?\d\b", 0.95),
                Pattern("cuit_frase", r"(?i)cui[tl]\s*\d{2}[\D\s]+\d{8}[\D\s]+\d", 0.95),
            ],
            context=["cuit", "cuil"],
        ),
        PatternRecognizer(
            supported_entity="DIRECCION",
            supported_language="es",
            patterns=[
                Pattern(
                    "domicilio_con_formula",
                    r"(?i)\b(?:domiciliad[oa]s?\s+en|(?:(?:con|fija(?:ndo)?|se\s+fija\s+el)\s+)?domicilio\s+(?:real|legal|especial|constituido|social)?\s*en|sit[oa]s?\s+en|(?:con\s+)?sede\s+(?:social\s+)?en)\s+"
                    r"([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ0-9\s\.\-]{2,35}?"
                    r"(?:\n\s*\d{1,2}\s*\n\s*)?"
                    r"\b\d{1,5}\b"
                    r"(?:\s*(?:,\s*)?(?:piso|p\.|pb|planta\s+baja)\s*\d{1,3}\b)?"
                    r"(?:\s*(?:,\s*)?(?:dto\.?|depto\.?|departamento|oficina|of\.)\s*[a-z0-9]{1,4}\b)?)"
                    r"(?:\s*(?:,\s*)?(?-i:([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+de\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)?)))?",
                    0.95,
                ),
                Pattern(
                    "calle_con_tipo",
                    r"(?i)\b(?:av(?:enida|\.)?|calle|pje\.?|pasaje|bv(?:ld|\.)?|boulevard|ruta\s+\d+|diagonal)\s+"
                    r"([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ0-9\s\.\-]{2,35}?\b\d{1,5}\b"
                    r"(?:\s*(?:,\s*)?(?:piso|p\.|pb|planta\s+baja)\s*\d{1,3}\b)?"
                    r"(?:\s*(?:,\s*)?(?:dto\.?|depto\.?|departamento|oficina|of\.)\s*[a-z0-9]{1,4}\b)?)"
                    r"(?:\s*(?:,\s*)?(?-i:([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+de\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)?)))?",
                    0.95,
                ),
            ],
            context=["domicilio", "domiciliado", "domiciliada", "sede", "calle", "avenida", "direccion"],
        ),
        PatternRecognizer(
            supported_entity="ORGANIZATION",
            supported_language="es",
            patterns=[
                Pattern(
                    "sociedad_nombre",
                    r'(?i)"([A-ZÁÉÍÓÚÑ0-9\s\.\-]{3,40}?\b(?:S\.?\s?R\.?\s?[LI]\.?|S\.?\s?A\.?\s?S\.?|S\.A\.))"',
                    0.95,
                ),
            ],
            context=["sociedad", "denomina", "srl", "sa"],
        ),
    ]


# ==============================================================================
# CLASE PRINCIPAL: ANONIMIZADOR CON CORREFERENCIA
# ==============================================================================

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
        self.idioma, self.modo, self.umbral, self.seudonimos = idioma, modo, umbral, seudonimos
        self.conteo = defaultdict(int)
        self.nuevo_documento()

    def nuevo_documento(self):
        """Reinicia el estado de clusters y seudónimos para un nuevo documento."""
        self.clusters = defaultdict(list)
        self.indices = defaultdict(int)
        self.mapa_reemplazos = {}
        self.equivalencias_dict = {}
        self.patrones = []

    def registrar_cluster(self, etype, val):
        val_clean = val.strip().strip(",.;:()[]{}¿?¡!\"'/-")
        if etype not in NUMERICOS:
            val_clean = re.sub(r"\s+", " ", limpiar_direccion_val(val_clean))
        n_val = clave(etype, val_clean)
        if not n_val:
            return f"<{ETIQUETAS_ES.get(etype, etype)}>"
        val_words = n_val.split()

        best_c, best_score = None, 0.0
        for c in self.clusters[etype]:
            if n_val in c["norm_variants"]:
                best_c, best_score = c, 1.0
                break
            c_norm = clave(etype, c["canonical"])
            if etype == "PERSON":
                c_words = c_norm.split()
                if len(val_words) >= 2 and len(c_words) >= len(val_words) and val_words == c_words[:len(val_words)]:
                    score = 0.95
                elif len(c_words) >= 2 and len(val_words) >= len(c_words) and c_words == val_words[:len(c_words)]:
                    score = 0.95
                elif len(val_words) == 1 and any(sim(val_words[0], s) >= 0.80 for s in c["surnames"]):
                    score = 0.92
                elif len(val_words) >= 2 and len(c_words) >= 2:
                    fn_sim = sim(" ".join(val_words[:-1]), " ".join(c_words[:-1]))
                    if fn_sim >= 0.85:
                        sur_sim = max((sim(val_words[-1], s) for s in c["surnames"]), default=0.0)
                        score = 0.92 if sur_sim >= 0.65 else 0.0
                    else:
                        score = sim(n_val, c_norm)
                else:
                    score = sim(n_val, c_norm)

            elif etype == "DIRECCION":
                # Misma calle y misma altura, aunque una mención agregue piso/oficina o localidad
                m_a = re.search(r"\b\d+\b", n_val)
                m_b = re.search(r"\b\d+\b", c_norm)
                if m_a and m_b and m_a.group() == m_b.group():
                    score = max(sim(n_val[:m_a.start()], c_norm[:m_b.start()]), 0.0)
                else:
                    score = 0.0

            elif etype in {"ORGANIZATION", "LOCATION"}:
                score = sim(n_val, c_norm)
            else:
                score = 0.0  # numéricos y demás: solo coincidencia exacta de la clave

            if score > best_score:
                best_score = score
                best_c = c

        threshold = 0.70 if etype in {"PERSON", "DIRECCION", "ORGANIZATION"} else 0.85
        if best_c and best_score >= threshold:
            best_c["variants"].add(val_clean)
            best_c["norm_variants"].add(n_val)
            best_c["frecuencia"][val_clean] += 1
            if etype == "PERSON" and val_words:
                best_c["surnames"].add(val_words[-1])
            # Canónico: la forma más completa y, entre iguales, la más repetida (las lecturas
            # erróneas del OCR suelen aparecer una sola vez)
            best_c["canonical"] = max(best_c["frecuencia"], key=lambda v: (len(v.split()), best_c["frecuencia"][v], -sum(not (ch.isalnum() or ch.isspace()) for ch in v), len(v)))
            self.mapa_reemplazos[(etype, n_val)] = best_c["tag"]
            self.equivalencias_dict[best_c["tag"]] = best_c["canonical"]
            return best_c["tag"]

        self.indices[etype] += 1
        tag = f"<{ETIQUETAS_ES.get(etype, etype)}_{self.indices[etype]}>"
        new_c = {
            "tag": tag,
            "canonical": val_clean,
            "variants": {val_clean},
            "frecuencia": defaultdict(int, {val_clean: 1}),
            "norm_variants": {n_val},
            "surnames": {val_words[-1]} if (etype == "PERSON" and val_words) else set(),
        }
        self.clusters[etype].append(new_c)
        self.mapa_reemplazos[(etype, n_val)] = tag
        self.equivalencias_dict[tag] = val_clean
        return tag

    def preparar_documento(self, lista_detecciones):
        """Agrupa las entidades del documento completo (primero las menciones más largas)
        y arma los patrones para encontrar las demás apariciones de cada una."""
        primera = {}
        for i, (etype, val) in sorted(enumerate(lista_detecciones), key=lambda x: -len(x[1][1])):
            tag = self.registrar_cluster(etype, val)
            primera[tag] = min(primera.get(tag, i), i)
        self._renumerar(primera)
        self.patrones = self._patrones_propagacion()

    def _renumerar(self, primera):
        """Numera las entidades de cada tipo según su primera aparición en el documento."""
        nuevo = {}
        for etype, clist in self.clusters.items():
            clist.sort(key=lambda c: primera.get(c["tag"], float("inf")))
            for n, c in enumerate(clist, start=1):
                nuevo[c["tag"]] = f"<{ETIQUETAS_ES.get(etype, etype)}_{n}>"
                c["tag"] = nuevo[c["tag"]]
        self.mapa_reemplazos = {k: nuevo.get(t, t) for k, t in self.mapa_reemplazos.items()}
        self.equivalencias_dict = {c["tag"]: c["canonical"] for clist in self.clusters.values() for c in clist}

    def _patrones_propagacion(self):
        """Regex de cada entidad conocida, para marcarla también donde el modelo no la detectó."""
        apellidos = defaultdict(set)
        for c in self.clusters.get("PERSON", []):
            for s in c["surnames"]:
                apellidos[s].add(c["tag"])

        patrones = []
        for etype, clist in self.clusters.items():
            for c in clist:
                formas = set()
                if etype in NUMERICOS:
                    for v in c["variants"]:
                        digitos = re.sub(r"\D", "", v)
                        if len(digitos) >= 7:
                            formas.add(r"(?<!\d)" + r"[\s.\-—–]?".join(digitos) + r"(?!\d)")
                else:
                    for v in c["variants"]:
                        palabras = v.split()
                        if len(palabras) >= 2 or (len(v) >= 4 and etype != "LOCATION"):
                            formas.add(r"\b" + SEPARADOR_OCR.join(patron_palabra(p) for p in palabras) + r"\b")
                    if etype == "PERSON":
                        # Apellido suelto, solo si no lo comparte con otra persona
                        for s in c["surnames"]:
                            if len(s) >= 4 and s not in TODAS_STOPWORDS and apellidos[s] == {c["tag"]}:
                                formas.add(r"\b" + patron_palabra(s) + r"\b")
                for f in formas:
                    patrones.append((etype, re.compile(f, re.IGNORECASE)))
        return patrones

    def completar(self, texto, resultados):
        """Agrega las apariciones de entidades ya conocidas que el modelo no detectó en este texto."""
        if not getattr(self, "patrones", None):
            return resultados
        candidatos = []
        for etype, patron in self.patrones:
            for m in patron.finditer(texto):
                if etype == "PERSON" and not texto[m.start()].isupper():
                    continue
                candidatos.append(RecognizerResult(etype, m.start(), m.end(), 1.0))

        # Organizaciones: el OCR suele deformar el nombre ("GRUPO RNEAZZO S.R.I."), se compara aproximado
        tokens = [m.span() for m in re.finditer(r"\S+", texto)]
        for c in self.clusters.get("ORGANIZATION", []):
            for v in c["variants"]:
                n = len(v.split())
                if n < 2:
                    continue
                for k in range(len(tokens) - n + 1):
                    ini, fin = tokens[k][0], tokens[k + n - 1][1]
                    ventana = texto[ini:fin].strip("\"'“”.,;:")
                    if sim(ventana, v) >= 0.75:
                        ini, fin = recortar_bordes(texto, ini, fin)
                        candidatos.append(RecognizerResult("ORGANIZATION", ini, fin, 0.9))

        # Una aparición conocida más larga reemplaza a detecciones parciales que contiene
        # (ej. el modelo vio solo "Gomez" en "María Laura Gomez")
        actuales = list(resultados)
        for c in sorted(candidatos, key=lambda r: -(r.end - r.start)):
            solapes = [r for r in actuales if c.start < r.end and c.end > r.start]
            if any(r.start < c.start or r.end > c.end or (r.end - r.start) >= (c.end - c.start) for r in solapes):
                continue
            actuales = [r for r in actuales if r not in solapes] + [c]
        return sorted(actuales, key=lambda r: r.start)

    def etiqueta(self, tipo, valor):
        tipo_es = ETIQUETAS_ES.get(tipo, tipo)
        if not self.seudonimos:
            if self.modo == "mascara":
                return "*" * len(valor)
            if self.modo == "borrar":
                return ""
            if self.modo == "hash":
                return hashlib.sha256(valor.encode()).hexdigest()[:16]
            return f"<{tipo_es}>"

        val_clean = valor.strip().strip(",.;:()[]{}¿?¡!\"'/-")
        n_val = clave(tipo, val_clean)
        if (tipo, n_val) in self.mapa_reemplazos:
            return self.mapa_reemplazos[(tipo, n_val)]
        return self.registrar_cluster(tipo, valor)

    def ocultar(self, tipo, valor):
        if not valor or not valor.strip():
            return valor
        tipo_es = ETIQUETAS_ES.get(tipo, tipo)
        self.conteo[tipo_es] += 1
        return self.etiqueta(tipo, valor)

    def equivalencias(self):
        def orden(etiqueta):
            tipo, _, n = etiqueta.strip("<>").rpartition("_")
            return tipo, int(n) if n.isdigit() else 0
        return {e: self.equivalencias_dict[e] for e in sorted(self.equivalencias_dict, key=orden)}

    def guardar_mapa(self, ruta_mapa, nombre_archivo=""):
        equiv = self.equivalencias()
        if not equiv:
            return
        datos = {
            "archivo": nombre_archivo,
            "total_entidades": len(equiv),
            "equivalencias": equiv,
        }
        with open(ruta_mapa, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
        print(f"  mapa -> {ruta_mapa}")

    def analizar_texto(self, texto):
        raw = self.analyzer.analyze(text=texto, language=self.idioma, score_threshold=self.umbral)
        return depurar_resultados(texto, raw, estricto=True)

    def _aplicar(self, texto, resultados):
        for r in resultados:
            tipo_es = ETIQUETAS_ES.get(r.entity_type, r.entity_type)
            self.conteo[tipo_es] += 1

        # Reemplazar de atrás hacia adelante para mantener los offsets exactos
        texto_anon = texto
        for r in sorted(resultados, key=lambda x: -x.start):
            val = texto[r.start:r.end]
            tag = self.etiqueta(r.entity_type, val)
            texto_anon = texto_anon[:r.start] + tag + texto_anon[r.end:]
        return texto_anon

    def analizar_documento(self, bloques):
        """Pase 1 sobre todos los bloques de un documento: detecta, agrupa entidades
        y completa las apariciones que el modelo no vio. Devuelve los resultados por bloque."""
        resultados = [self.analizar_texto(b) if b and b.strip() else [] for b in bloques]
        self.preparar_documento([(r.entity_type, b[r.start:r.end])
                                 for b, res in zip(bloques, resultados) for r in res])
        return [self.completar(b, res) if b else res for b, res in zip(bloques, resultados)]

    def documento(self, bloques):
        """Anonimiza una lista de textos que forman un mismo documento, con etiquetas homogéneas."""
        resultados = self.analizar_documento(bloques)
        return [self._aplicar(b, res) if res else b for b, res in zip(bloques, resultados)]

    def texto(self, texto):
        if not texto or not texto.strip():
            return texto
        return self.documento([texto])[0]

    def valores(self, valores, contexto, cache):
        pendientes = [v for v in dict.fromkeys(valores) if v not in cache and v and v.strip()]
        if pendientes:
            analisis = self.batch.analyze_iterator(pendientes, language=self.idioma, batch_size=256,
                                                   score_threshold=self.umbral, context=[contexto])
            for v, res in zip(pendientes, analisis):
                dep = depurar_resultados(v, res, estricto=False)
                cache[v] = self._aplicar(v, dep)
        return [cache.get(v, v) for v in valores]


# ==============================================================================
# PROCESAMIENTO DE FORMATOS Y PDF
# ==============================================================================

def lista(texto):
    return {c.strip().lower() for c in texto.split(",") if c.strip()} if texto else set()


def leer_texto(ruta):
    for cod in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return ruta.read_text(encoding=cod)
        except UnicodeDecodeError:
            continue
    return None


def anonimizar_filas(filas, cabecera, anon, columnas, ocultas, caches):
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


def procesar_csv(entrada, salida, anon, columnas, ocultas):
    contenido = leer_texto(entrada)
    try:
        dialecto = csv.Sniffer().sniff(contenido[:20_000], delimiters=",;\t|")
    except csv.Error:
        dialecto = csv.excel_tab if entrada.suffix.lower() == ".tsv" else csv.excel
    lector = csv.reader(contenido.splitlines(), dialecto)
    cabecera = next(lector)
    caches = defaultdict(dict)
    with open(salida, "w", encoding="utf-8-sig", newline="") as f:
        escritor = csv.writer(f, delimiter=dialecto.delimiter, quotechar=dialecto.quotechar or '"')
        escritor.writerow(cabecera)
        lote = []
        for fila in lector:
            lote.append(fila)
            if len(lote) >= FILAS_LOTE:
                anonimizar_filas(lote, cabecera, anon, columnas, ocultas, caches)
                escritor.writerows(lote)
                lote = []
        if lote:
            anonimizar_filas(lote, cabecera, anon, columnas, ocultas, caches)
            escritor.writerows(lote)


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


def procesar_docx(entrada, salida, anon):
    from docx import Document
    doc = Document(entrada)

    def bloques(contenedor):
        yield from contenedor.paragraphs
        for tabla in contenedor.tables:
            for fila in tabla.rows:
                for celda in fila.cells:
                    yield from bloques(celda)

    parrafos = list(bloques(doc))
    for seccion in doc.sections:
        for parte in (seccion.header, seccion.footer):
            parrafos.extend(bloques(parte))

    # Todo el documento a la vez, para que la misma persona tenga la misma etiqueta en cualquier párrafo
    originales = [p.text for p in parrafos]
    for p, original, nuevo in zip(parrafos, originales, anon.documento(originales)):
        if nuevo != original and p.runs:
            p.runs[0].text = nuevo
            for run in p.runs[1:]:
                run.text = ""
    doc.save(salida)


def quitar_subrayado(imagen):
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


# Hoja de protocolo notarial: el sello del escribano va arriba a la derecha, sobre el número de folio
ZONA_SELLO_PROTOCOLO = (0.63, 0.085, 0.84, 0.18)  # fracciones de ancho y alto de la página
# Texto que el OCR alcanza a leer de un sello redondo (en mayúsculas, a diferencia del cuerpo)
TEXTO_SELLO = re.compile(r"(?:ES)?C?RIBAN[OA]S?|MAT\.?")


def zonas_sello(ancho, alto, palabras):
    """Recuadros que tapan los sellos del escribano, cuyo nombre el OCR no puede leer."""
    x0, y0, x1, y1 = ZONA_SELLO_PROTOCOLO
    zonas = [(x0 * ancho, y0 * alto, x1 * ancho, y1 * alto)]
    d = 0.24 * ancho  # diámetro aproximado de un sello redondo
    for p in palabras:
        if not TEXTO_SELLO.fullmatch(p["text"]):
            continue
        b = p["box"]
        cx = b["x"] + b["width"] / 2
        if p["text"].startswith("MAT"):
            cy = b["y"] + b["height"] / 2
        else:  # "ESCRIBANO" va en el borde inferior del sello
            cy = b["y"] + b["height"] - d / 2
        zonas.append((cx - d / 2, cy - d / 2, cx + d / 2, cy + d / 2 + b["height"]))
    return zonas


def es_ruido(linea):
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
    letras_cant = sum(c.isalpha() for c in linea)
    if letras_cant == 0:
        return True
    if letras_cant < 4 and not re.fullmatch(r"[a-záéíóúñ]{1,3}", linea, re.I):
        return True
    compacto = re.sub(r"\s", "", linea)
    if compacto and letras_cant * 2 < len(compacto) and len(linea) < 24:
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
    bloques = secciones(texto_continuo(paginas))
    partes = [
        "---",
        "documento: escritura de constitución de sociedad",
        "uso: lectura del texto anonimizado",
        "anonimizado: true",
        "origen: OCR de un PDF escaneado",
        "etiquetas: <PERSONA_1>, <PERSONA_2> personas diferenciadas; <DIRECCION_1> domicilios; <LUGAR_1> ciudades/países; <DNI_1> documentos; <CUIT_1> CUIT o CUIL; <TELEFONO_1> teléfonos; <ORGANIZACION_1> organizaciones",
        "---",
        "",
        f"# {nombre}",
        "",
        "Texto anonimizado, en el orden de la escritura. Las etiquetas numeradas identifican de forma consistente a la misma entidad a lo largo de todo el documento para su procesamiento por IA. El escaneo dejó errores de lectura; no hay que reconstruir lo tapado.",
        "",
    ]
    for titulo, cuerpo in bloques:
        partes.append(f"## {titulo}")
        partes.append("")
        partes.append(incisos(cuerpo) if titulo == "Objeto social" else cuerpo)
        partes.append("")
    return "\n".join(partes).strip() + "\n"


def procesar_pdf(entrada, salida, anon, sellos=False):
    origen = pymupdf.open(entrada)
    destino = pymupdf.open()
    paginas_info = []

    # Pase 1: lectura de todas las páginas
    for n, pagina in enumerate(origen, start=1):
        imagen, crudo, palabras, uso_ocr = contenido_pagina(pagina)
        if uso_ocr:
            print(f"  página {n}: OCR ({len(crudo)} caracteres)", flush=True)
        paginas_info.append({"num": n, "imagen": imagen, "crudo": crudo, "palabras": palabras})

    # Análisis conjunto: misma entidad, misma etiqueta en todas las páginas
    for info, resultados in zip(paginas_info, anon.analizar_documento([i["crudo"] for i in paginas_info])):
        info["resultados"] = resultados

    # Pase 2: Aplicación y redacción sobre las páginas
    textos = []
    for info in paginas_info:
        crudo = info["crudo"]
        imagen = info["imagen"]
        palabras = info["palabras"]
        resultados = info["resultados"]

        anonimizado = anon._aplicar(crudo, resultados)
        textos.append(anonimizado)

        dibujo = ImageDraw.Draw(imagen)
        papel = color_papel(imagen)
        cache = {}
        for r in resultados:
            valor = crudo[r.start:r.end]
            tapar(dibujo, palabras, r, anon.etiqueta(r.entity_type, valor), papel, cache)

        if sellos:
            for zona in zonas_sello(imagen.width, imagen.height, palabras):
                dibujo.rectangle(zona, fill=papel)

        buffer = io.BytesIO()
        imagen.save(buffer, format="JPEG", quality=80)
        pagina_orig = origen[info["num"] - 1]
        nueva = destino.new_page(width=pagina_orig.rect.width, height=pagina_orig.rect.height)
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


# ==============================================================================
# DESANONIMIZACIÓN Y MAIN
# ==============================================================================

def desanonimizar(archivo_entrada, archivo_mapa, archivo_salida=None):
    if not archivo_entrada.exists():
        sys.exit(f"No existe el archivo de texto: {archivo_entrada}")
    if not archivo_mapa.exists():
        sys.exit(f"No existe el archivo mapa: {archivo_mapa}")

    with open(archivo_mapa, "r", encoding="utf-8") as f:
        mapa_datos = json.load(f)

    equivalencias = mapa_datos.get("equivalencias", {})
    if not equivalencias:
        sys.exit(f"El archivo {archivo_mapa} no contiene equivalencias.")

    texto = leer_texto(archivo_entrada)
    if texto is None:
        sys.exit(f"No se pudo leer el contenido de {archivo_entrada}")

    reemplazos = 0
    for etiqueta, valor_real in sorted(equivalencias.items(), key=lambda x: -len(x[0])):
        if etiqueta in texto:
            cant = texto.count(etiqueta)
            texto = texto.replace(etiqueta, valor_real)
            reemplazos += cant
            print(f"  {etiqueta} -> {valor_real} ({cant} veces)")

    if archivo_salida is None:
        archivo_salida = archivo_entrada.parent / f"{archivo_entrada.stem}.restaurado{archivo_entrada.suffix}"

    archivo_salida.write_text(texto, encoding="utf-8")
    print(f"\nTexto restaurado guardado en: {archivo_salida}")
    print(f"Total de reemplazos: {reemplazos}")


def procesar(entrada, salida, anon, args):
    anon.nuevo_documento()
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
        procesar_pdf(entrada, salida, anon, getattr(args, "sellos", False))
    else:
        print(f"  Formato no soportado, se omite: {entrada.name}")
        return
    print(f"  -> {salida}")
    if anon.seudonimos:
        ruta_mapa = salida.parent / f"{entrada.stem}.mapa.json"
        anon.guardar_mapa(ruta_mapa, entrada.name)


def main():
    ap = argparse.ArgumentParser(description="Anonimiza documentos con Presidio")
    ap.add_argument("archivo", type=Path, nargs="?", help="si se omite, procesa todo input/")
    ap.add_argument("--idioma", choices=list(MODELOS), default="es")
    ap.add_argument("--modo", choices=["reemplazar", "mascara", "hash", "borrar"], default="reemplazar")
    ap.add_argument("--seudonimos", action="store_true", help="diferencia entidades individualmente (<PERSONA_1>, <DIRECCION_1>...) y guarda .mapa.json")
    ap.add_argument("--desanonimizar", type=Path, help="archivo de texto (ej. respuesta de la IA) a restaurar con datos reales")
    ap.add_argument("--mapa", type=Path, help="archivo .mapa.json con equivalencias para desanonimizar")
    ap.add_argument("--sellos", action="store_true", help="(PDF) tapa los sellos del escribano en hojas de protocolo notarial")
    ap.add_argument("--umbral", type=float, default=0.35)
    ap.add_argument("--columnas", type=lista, default=set())
    ap.add_argument("--ocultar-columnas", type=lista, default=set())
    args = ap.parse_args()

    if args.desanonimizar:
        if not args.mapa:
            sys.exit("Para desanonimizar debes indicar el mapa de equivalencias con --mapa ruta.mapa.json")
        desanonimizar(args.desanonimizar, args.mapa)
        return

    if args.archivo:
        if not args.archivo.exists():
            sys.exit(f"No existe: {args.archivo}")
        trabajos = [(args.archivo, OUTPUT / args.archivo.name)]
    else:
        INPUT.mkdir(exist_ok=True)
        trabajos = [(p, OUTPUT / p.relative_to(INPUT)) for p in sorted(INPUT.rglob("*")) if p.is_file() and p.name != ".gitkeep"]
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
