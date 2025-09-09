import os
import io
import re
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import RGBColor, Inches
from werkzeug.datastructures import FileStorage

# Intentamos usar Pillow para conservar la relación de aspecto al escalar.
# Si no está disponible, trabajaremos con un "best effort" por ancho.
try:
    from PIL import Image
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

app = Flask(__name__)

@app.get("/health")
def health():
    return jsonify({"ok": True})

# ------------ Utilidades Markdown -> DOCX (ahora con imágenes) ----------------

HEADING_RE      = re.compile(r'^(#{1,6})\s+(.*)$')
UL_RE           = re.compile(r'^\s*[-*+]\s+(.*)$')
OL_RE           = re.compile(r'^\s*\d+\.\s+(.*)$')
TABLE_ROW_RE    = re.compile(r'^\s*\|(.+)\|\s*$')
TABLE_ALIGN_RE  = re.compile(r'^\s*\|?\s*(:?-{3,}:?\s*\|)+\s*(:?-{3,}:?)\s*\|?\s*$')
# Imagen en línea tipo: ![alt](fileName.ext)
IMG_LINE_RE     = re.compile(r'^\s*!\[([^\]]*)\]\(([^)]+)\)\s*$')
# Imagen embebida en medio de texto (menos habitual en Mistral OCR, pero soportado)
IMG_INLINE_RE   = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')

def force_styles_black(doc: Document):
    """Pone el color de fuente a negro en estilos usados (encabezados, normal, listas)."""
    target_styles = ["Normal", "List Paragraph", "List Bullet", "List Number"]
    target_styles += [f"Heading {i}" for i in range(1, 10)]
    for name in target_styles:
        try:
            style = doc.styles[name]
            if style and style.font:
                style.font.color.rgb = RGBColor(0, 0, 0)
        except KeyError:
            pass

def add_paragraph(doc, text):
    if not text.strip():
        doc.add_paragraph("")  # línea en blanco
    else:
        doc.add_paragraph(text)

def flush_list(doc, buf, ordered):
    if not buf:
        return
    style = "List Number" if ordered else "List Bullet"
    for item in buf:
        doc.add_paragraph(item, style=style)
    buf.clear()

def flush_table(doc, rows):
    if not rows:
        return
    # Ignora fila de alineación tipo | :-- | --- | :--: |
    filtered = [r for r in rows if not TABLE_ALIGN_RE.match(r)]
    if not filtered:
        return
    matrix = []
    for r in filtered:
        m = TABLE_ROW_RE.match(r)
        if not m:
            continue
        cells = [c.strip() for c in m.group(1).split("|")]
        matrix.append(cells)
    if not matrix:
        return
    cols = max(len(r) for r in matrix)
    table = doc.add_table(rows=len(matrix), cols=cols)
    table.style = "Table Grid"
    for i, row in enumerate(matrix):
        for j in range(cols):
            table.cell(i, j).text = row[j] if j < len(row) else ""

def add_image_paragraph(doc: Document, filename: str, img_bytes: bytes, alt_text: str = ""):
    """
    Inserta una imagen ajustando su ancho al ancho útil de la página.
    Mantiene proporción si PIL está disponible.
    """
    # Calcular ancho útil de la página (en EMUs) -> convertimos luego a Inches
    section = doc.sections[-1]
    page_width = section.page_width
    left_margin = section.left_margin
    right_margin = section.right_margin
    usable_width_emu = page_width - left_margin - right_margin

    # EMU por pulgada en python-docx
    EMUS_PER_INCH = 914400
    usable_width_in = float(usable_width_emu) / EMUS_PER_INCH

    stream = io.BytesIO(img_bytes)

    if PIL_AVAILABLE:
        try:
            with Image.open(io.BytesIO(img_bytes)) as im:
                width_px, height_px = im.size
                dpi = im.info.get("dpi", (96, 96))[0] or 96
                width_in = width_px / float(dpi)
                height_in = height_px / float(dpi)

                # Escalado: si la imagen es más ancha que el ancho útil, reducimos
                scale = 1.0
                if width_in > usable_width_in:
                    scale = usable_width_in / width_in

                new_width_in = width_in * scale
                # Insertamos con ancho en inches; docx mantiene la relación
                p = doc.add_paragraph()
                run = p.add_run()
                run.add_picture(stream, width=Inches(new_width_in))
                if alt_text:
                    # python-docx no soporta alt_text nativo; añadimos como pie/nota ligera
                    p2 = doc.add_paragraph(alt_text)
                    p2.style = "Caption" if "Caption" in [s.name for s in doc.styles] else "Normal"
                return
        except Exception:
            # Si hay cualquier problema con PIL, fallback
            pass

    # Fallback sin PIL: insertar a ancho útil (podría distorsionar en casos raros)
    p = doc.add_paragraph()
    run = p.add_run()
    stream.seek(0)
    run.add_picture(stream, width=Inches(usable_width_in))
    if alt_text:
        p2 = doc.add_paragraph(alt_text)
        p2.style = "Caption" if "Caption" in [s.name for s in doc.styles] else "Normal"

def handle_inline_images(doc: Document, text: str, images: dict):
    """
    Parte una línea con texto + imágenes inline en runs y añade imágenes donde tocan.
    Si falta la imagen, deja un marcador [imagen faltante: filename].
    """
    parts = []
    last_end = 0
    for m in IMG_INLINE_RE.finditer(text):
        if m.start() > last_end:
            parts.append(("text", text[last_end:m.start()]))
        alt = (m.group(1) or "").strip()
        fname = (m.group(2) or "").strip()
        parts.append(("img", (fname, alt)))
        last_end = m.end()
    if last_end < len(text):
        parts.append(("text", text[last_end:]))

    if not parts:
        # no había imagen inline; devolver como párrafo normal
        add_paragraph(doc, text)
        return

    # Construimos: texto previo como párrafo, imagen como bloque independiente para
    # asegurar buen flujo en Word, y continuamos.
    for kind, payload in parts:
        if kind == "text":
            if payload.strip():
                add_paragraph(doc, payload.strip())
            else:
                doc.add_paragraph("")
        else:
            fname, alt = payload
            blob = images.get(fname)
            if blob:
                add_image_paragraph(doc, fname, blob, alt_text=alt)
            else:
                add_paragraph(doc, f"[imagen faltante: {fname}]")

def markdown_to_doc(md_text: str, images: dict, filename: str = "output.docx"):
    """
    Convierte Markdown a DOCX, insertando imágenes encontradas en `images`.
    `images` es un dict: {fileName(str): bytes}
    """
    doc = Document()
    force_styles_black(doc)

    lines = md_text.splitlines()
    ul_buf, ol_buf, tbl_buf = [], [], []
    in_table = False
    para_buf = []

    def flush_para():
        if para_buf:
            # Antes de cerrar párrafo, procesar imágenes inline si existieran:
            text = " ".join(para_buf).strip()
            para_buf.clear()
            if IMG_INLINE_RE.search(text):
                handle_inline_images(doc, text, images)
            else:
                add_paragraph(doc, text)

    for raw in lines:
        line = raw.rstrip("\n")

        # Tablas
        if TABLE_ROW_RE.match(line):
            in_table = True
            tbl_buf.append(line)
            continue
        else:
            if in_table:
                flush_para()
                flush_list(doc, ul_buf, ordered=False)
                flush_list(doc, ol_buf, ordered=True)
                flush_table(doc, tbl_buf)
                tbl_buf = []
                in_table = False

        # Encabezados
        m = HEADING_RE.match(line)
        if m:
            flush_para()
            flush_list(doc, ul_buf, ordered=False)
            flush_list(doc, ol_buf, ordered=True)
            level = len(m.group(1))
            text = m.group(2).strip()
            level = min(max(level, 1), 9)
            doc.add_heading(text, level=level)
            continue

        # Listas
        m_ul = UL_RE.match(line)
        m_ol = OL_RE.match(line)
        if m_ul:
            flush_para()
            flush_list(doc, ol_buf, ordered=True)
            ul_buf.append(m_ul.group(1).strip())
            continue
        if m_ol:
            flush_para()
            flush_list(doc, ul_buf, ordered=False)
            ol_buf.append(m_ol.group(1).strip())
            continue

        # Línea de imagen (bloque)
        m_img = IMG_LINE_RE.match(line)
        if m_img:
            flush_para()
            flush_list(doc, ul_buf, ordered=False)
            flush_list(doc, ol_buf, ordered=True)
            alt = (m_img.group(1) or "").strip()
            fname = (m_img.group(2) or "").strip()
            blob = images.get(fname)
            if blob:
                add_image_paragraph(doc, fname, blob, alt_text=alt)
            else:
                add_paragraph(doc, f"[imagen faltante: {fname}]")
            continue

        # Separador de párrafos
        if not line.strip():
            flush_para()
            flush_list(doc, ul_buf, ordered=False)
            flush_list(doc, ol_buf, ordered=True)
            continue

        # Texto normal (acumulado)
        para_buf.append(line.strip())

    # flush final
    flush_para()
    flush_list(doc, ul_buf, ordered=False)
    flush_list(doc, ol_buf, ordered=True)
    if tbl_buf:
        flush_table(doc, tbl_buf)

    force_styles_black(doc)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, filename

# -------------------- Endpoint principal --------------------

@app.post("/docx")
def make_docx():
    """
    Acepta:
      1) JSON: {"markdown": "...", "filename": "informe.docx"} (compatibilidad)
      2) multipart/form-data:
           - campo de texto: "markdown"
           - n veces: archivo "file" (el nombre del archivo DEBE coincidir con el del Markdown)
    """
    # Intento 1: JSON (como el código original)
    data = request.get_json(silent=True)
    if data and isinstance(data, dict) and ("markdown" in data or "text" in data):
        filename = data.get("filename", "output.docx")
        if data.get("markdown"):
            buf, fname = markdown_to_doc(data["markdown"], images={}, filename=filename)
        else:
            # compat: texto plano en un párrafo
            doc = Document()
            force_styles_black(doc)
            add_paragraph(doc, data.get("text", ""))
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            fname = filename

        return send_file(
            buf,
            as_attachment=True,
            download_name=fname,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    # Intento 2: multipart/form-data (n8n)
    # - markdown en request.form["markdown"]
    # - imágenes en request.files.getlist("file")
    if request.form and ("markdown" in request.form or "text" in request.form):
        md_text = request.form.get("markdown", None)
        plain_text = request.form.get("text", None)
        filename = request.form.get("filename", "output.docx")

        # Recoger todas las imágenes bajo la misma key "file"
        images_map = {}
        file_list = request.files.getlist("file")
        for f in file_list:
            if not isinstance(f, FileStorage):
                continue
            # Guardamos por filename exacto (que debe coincidir con el que aparece en el Markdown)
            if f.filename:
                images_map[f.filename] = f.read()

        if md_text:
            buf, fname = markdown_to_doc(md_text, images_map, filename=filename)
        else:
            # compat: texto plano si solo viene 'text' sin markdown
            doc = Document()
            force_styles_black(doc)
            add_paragraph(doc, plain_text or "")
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            fname = filename

        return send_file(
            buf,
            as_attachment=True,
            download_name=fname,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    # Si no es ni JSON ni multipart esperado:
    return jsonify({"error": "Bad request: envía JSON con 'markdown' o multipart/form-data con 'markdown' y archivos 'file'."}), 400


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
