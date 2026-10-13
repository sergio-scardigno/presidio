# Presidio

Anonimiza documentos con [Microsoft Presidio](https://github.com/microsoft/presidio). Detecta personas, lugares, correos, teléfonos y, en español, DNI y CUIT/CUIL. El resultado queda en `output/`.

Los archivos de `input/` y `output/` no se versionan: pueden tener datos personales.

## 1. Requisitos

- Windows 10 o 11
- Python 3.13
- Paquete de OCR en español de Windows (es-MX), solo para PDF escaneados

## 2. Instalación

En PowerShell, desde esta carpeta:

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

La primera instalación descarga los modelos de español e inglés y tarda varios minutos.

## 3. Poner los archivos

Copiá los documentos a `input/`. Formatos admitidos: `.txt`, `.md`, `.csv`, `.tsv`, `.json`, `.log`, `.docx`, `.xlsx` y `.pdf`.

## 4. Anonimizar

Todo lo que haya en `input/`:

```powershell
python anonimizar.py
```

Un solo archivo:

```powershell
python anonimizar.py input\documento.csv
```

La primera ejecución carga los modelos y puede tardar. Al terminar muestra cuántas entidades encontró y la ruta del resultado.

## 5. Elegir cómo se reemplaza

| Opción | Efecto |
|---|---|
| `--idioma es` o `--idioma en` | Idioma del texto. Por defecto, español. |
| `--modo reemplazar` | Sustituye por etiquetas descriptivas en español (`<PERSONA>`, `<DIRECCION>`, `<LUGAR>`, `<DNI>`, etc.). Modo por defecto. |
| `--modo mascara` | Cubre el dato con asteriscos. |
| `--modo hash` | Lo sustituye por un hash SHA-256. |
| `--modo borrar` | Lo elimina. |
| `--seudonimos` | Numera cada entidad (`<PERSONA_1>`, `<DNI_1>`, `<TELEFONO_1>`, `<DIRECCION_1>`) con la misma etiqueta en todo el documento: reconoce el apellido suelto, nombres sin tilde y DNI o teléfonos escritos con otro formato. Guarda `archivo.mapa.json`. |
| `--desanonimizar resp.txt` | Restaura los datos reales en un texto (ej. una respuesta recibida) usando `--mapa`. |
| `--mapa archivo.mapa.json` | Ruta al mapa de equivalencias generado por `--seudonimos`. |
| `--sellos` | En PDF de escrituras, tapa los sellos del escribano (el OCR no lee su nombre). |
| `--umbral 0.35` | Confianza mínima, de 0 a 1. Un valor más alto detecta menos. |
| `--columnas a,b,c` | En CSV o Excel, analiza solo esas columnas. |
| `--ocultar-columnas a,b` | Sustituye la columna entera sin analizarla. |

### Compartir el documento sin exponer los datos

El programa no se conecta a ningún servicio externo: todo corre en tu computadora. Este es el uso previsto para compartir un documento con un servicio o una persona externa sin pasarle datos personales:

1. **Anonimizá el documento con etiquetas numeradas:**
   ```powershell
   python anonimizar.py input\contrato.docx --seudonimos
   ```
   Genera `output\contrato.docx` (con `<PERSONA_1>`, `<DNI_1>`, `<DIRECCION_1>`...) y `output\contrato.mapa.json`, que guarda a qué dato real corresponde cada etiqueta. El mapa no se comparte.

2. **Compartí el texto anonimizado** y pedí lo que necesites (revisión, resumen, redacción). Quien lo recibe solo ve las etiquetas.

3. **Guardá la respuesta en un `.txt` y restaurá los datos reales:**
   ```powershell
   python anonimizar.py --desanonimizar respuesta.txt --mapa output\contrato.mapa.json
   ```
   Crea `respuesta.restaurado.txt` con los nombres, documentos y domicilios reales.

Ejemplo para un CSV con documento y CUIL en columnas propias:

```powershell
python anonimizar.py input\padron.csv --seudonimos --ocultar-columnas nro_documento,cuil
```

## 6. Qué sale de un PDF

| Archivo | Contenido |
|---|---|
| `.txt` | Texto anonimizado, página por página. |
| `.pdf` | El escaneo original, con los datos tapados. |
| `.md` | El mismo texto, ordenado en cláusulas. |

Si la página ya tiene texto, se usa ese texto. Si es un escaneo, se lee con el OCR en español de Windows. Hace falta el paquete de idioma es-MX.
