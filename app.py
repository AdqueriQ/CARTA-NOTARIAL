"""
Backend Flask — Carta Notarial BBVA
Despliega en Render.com como Web Service.

Requirements (requirements.txt):
    flask
    flask-cors

El archivo CARTA_NOTARIAL.docx debe estar en el mismo directorio que app.py.
LibreOffice debe estar instalado en el servidor (en Render: usa el buildpack o Dockerfile).
"""

import os, zipfile, re, tempfile, subprocess
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # permite llamadas desde GitHub Pages

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DOCX_ORIG = os.path.join(BASE_DIR, 'CARTA_NOTARIAL.docx')

# ─── helpers ─────────────────────────────────────────────
def xml_escape(s):
    return (s.replace('&','&amp;')
             .replace('<','&lt;')
             .replace('>','&gt;')
             .replace('"','&quot;')
             .replace("'",'&apos;'))

def make_run(text, bold=False, underline=False):
    bp = '<w:b/><w:bCs/>' if bold else ''
    up = '<w:u w:val="single"/>' if underline else ''
    rpr = (
        '<w:rPr><w:rFonts w:ascii="Lato" w:eastAsia="Lato" w:hAnsi="Lato" w:cs="Lato"/>'
        + bp + up +
        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
    )
    return '<w:r>' + rpr + '<w:t xml:space="preserve">' + xml_escape(text) + '</w:t></w:r>'

def make_cell(text, width, align='both'):
    return (
        '<w:tc><w:tcPr><w:tcW w:w="' + str(width) + '" w:type="dxa"/></w:tcPr>'
        '<w:p w:rsidR="00B136EE" w:rsidRDefault="00B136EE" w:rsidP="00B136EE">'
        '<w:pPr><w:jc w:val="' + align + '"/>'
        '<w:rPr><w:rFonts w:ascii="Lato" w:eastAsia="Lato" w:hAnsi="Lato" w:cs="Lato"/>'
        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
        + make_run(text) +
        '</w:p></w:tc>'
    )

def normalize_products(data):
    productos = data.get('productos')
    if isinstance(productos, list) and productos:
        clean = []
        for p in productos[:6]:
            if not isinstance(p, dict):
                continue
            tipo = str(p.get('tipo_prod', '')).strip().upper()
            contrato = str(p.get('ncontrato', '')).strip().upper()
            monto = str(p.get('monto', '')).strip()
            if not (tipo and contrato and monto):
                continue
            clean.append({'tipo_prod': tipo, 'ncontrato': contrato, 'monto': monto})
        if clean:
            return clean

    # Compatibilidad con el formato antiguo (un solo producto)
    return [{
        'tipo_prod': str(data.get('tipo_prod', '')).strip().upper(),
        'ncontrato': str(data.get('ncontrato', '')).strip().upper(),
        'monto': str(data.get('monto', '')).strip(),
    }]

def fill_docx(data):
    """
    Recibe dict con claves:
        fecha, nombre, representante, direccion,
        cod_ofic, nom_ofic, tipo_prod, ncontrato, monto, telefono
    Devuelve bytes del .docx llenado.
    """
    files = {}
    with zipfile.ZipFile(DOCX_ORIG, 'r') as z:
        for name in z.namelist():
            files[name] = z.read(name)

    doc = files['word/document.xml'].decode('utf-8')
    is_juridica = str(data.get('tipo_persona', 'natural')).strip().lower() == 'juridica'

    # ── 1. Fecha (reemplaza párrafo completo) ──────────────
    old_fecha = doc[
        doc.index('<w:p w:rsidR="008601FC" w:rsidRDefault="00B136EE">'):
        doc.index('</w:p>', doc.index('<w:p w:rsidR="008601FC" w:rsidRDefault="00B136EE">')) + 6
    ]
    new_fecha = (
        '<w:p w:rsidR="008601FC" w:rsidRDefault="00B136EE">'
        '<w:pPr><w:spacing w:line="288" w:lineRule="auto"/>'
        '<w:ind w:left="4248" w:firstLine="708"/>'
        '<w:jc w:val="right"/>'
        '<w:rPr><w:rFonts w:ascii="Lato" w:eastAsia="Lato" w:hAnsi="Lato" w:cs="Lato"/>'
        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
        + make_run(xml_escape(data['fecha'])) +
        '</w:p>'
    )
    doc = doc.replace(old_fecha, new_fecha)

    # ── 2. NOMBRE COMPLETO ─────────────────────────────────
    if is_juridica:
        doc = re.sub(
            r'<w:r[^>]*>.*?<w:t>NOMBRE COMPLETO</w:t>.*?</w:r>',
            make_run(data['nombre'], bold=True),
            doc,
            count=1,
            flags=re.DOTALL,
        )
    else:
        doc = doc.replace(
            '<w:t>NOMBRE COMPLETO</w:t>',
            '<w:t xml:space="preserve">' + xml_escape(data['nombre']) + '</w:t>'
        )

    # ── 3. Representante Legal (insertar párrafo si aplica) ─
    if data.get('representante','').strip():
        rep_para = (
            '<w:p w:rsidR="00B136EE" w:rsidRDefault="00B136EE">'
            '<w:pPr><w:spacing w:line="288" w:lineRule="auto"/>'
            '<w:jc w:val="both"/>'
            '<w:rPr><w:rFonts w:ascii="Lato" w:eastAsia="Lato" w:hAnsi="Lato" w:cs="Lato"/>'
            '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
            + make_run(data['representante'], bold=is_juridica) +
            '</w:p>'
        )
        # insertar después del párrafo NOMBRE
        marker = '<w:t xml:space="preserve">' + xml_escape(data['nombre']) + '</w:t>'
        idx = doc.index(marker)
        end_para = doc.index('</w:p>', idx) + 6
        doc = doc[:end_para] + rep_para + doc[end_para:]

    # ── 4. DIRECCION ───────────────────────────────────────
    doc = doc.replace(
        '<w:t>DIRECCION</w:t>',
        '<w:t xml:space="preserve">' + xml_escape(data['direccion']) + '</w:t>'
    )

    # ── 5. Código + ciudad destino ─────────────────────────
    doc = doc.replace(
        '<w:t>0310 TARAPOTO</w:t>',
        '<w:t xml:space="preserve">' + xml_escape(data['cod_ofic'] + ' ' + data['nom_ofic']) + '</w:t>'
    )

    # ── 6. 0310 OFICINA TARAPOTO en cuerpo ─────────────────
    # Están en runs separados: "0310 " | "OFICINA " | "TARAPOTO"
    doc = doc.replace('<w:t>0310 </w:t>',
                      '<w:t xml:space="preserve">' + xml_escape(data['cod_ofic']) + ' </w:t>', 1)
    # OFICINA es fijo, no cambia
    doc = doc.replace('<w:t>TARAPOTO</w:t>',
                      '<w:t xml:space="preserve">' + xml_escape(data['nom_ofic']) + '</w:t>')

    # ── 7. Teléfono ────────────────────────────────────────
    doc = doc.replace(
        '<w:t xml:space="preserve"> 996293543</w:t>',
        '<w:t xml:space="preserve"> ' + xml_escape(data['telefono']) + '</w:t>'
    )

    # ── 8. Llenar filas de la tabla (hasta 6 productos) ─────
    tbl_start = doc.index('<w:tbl')
    tbl_end   = doc.index('</w:tbl>', tbl_start) + 8
    tbl_xml   = doc[tbl_start:tbl_end]

    productos = normalize_products(data)

    row_matches = re.findall(r'<w:tr\b.*?</w:tr>', tbl_xml, flags=re.DOTALL)
    if len(row_matches) < 2:
        raise ValueError('No se encontró una fila de datos en la tabla del documento')

    # Usar la primera fila sin texto como plantilla; si no existe, usar la segunda.
    old_data_row = None
    for r in row_matches[1:]:
        if '<w:t' not in r:
            old_data_row = r
            break
    if old_data_row is None:
        old_data_row = row_matches[1]

    row_chunks = []
    for p in productos:
        row_chunks.append(
            '<w:tr w:rsidR="00B136EE" w:rsidTr="00B136EE">'
            + make_cell(p['tipo_prod'], 2942, 'left')
            + make_cell(p['ncontrato'], 2943, 'center')
            + make_cell('S/ ' + p['monto'], 2943, 'center')
            + '</w:tr>'
        )
    new_tbl = tbl_xml.replace(old_data_row, ''.join(row_chunks), 1)
    doc = doc[:tbl_start] + new_tbl + doc[tbl_end:]

    files['word/document.xml'] = doc.encode('utf-8')

    # Escribir docx en memoria
    import io
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, content in files.items():
            zout.writestr(name, content)
    buf.seek(0)
    return buf.read()

# ─── rutas ───────────────────────────────────────────────
@app.route('/', methods=['GET'])
def index():
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/generar-pdf', methods=['POST'])
def generar_pdf():
    data = request.get_json(force=True)

    required = ['fecha','nombre','direccion','cod_ofic','nom_ofic',
                'tipo_prod','ncontrato','monto','telefono']
    for k in required:
        if not data.get(k,'').strip():
            return jsonify({'error': f'Campo requerido: {k}'}), 400

    productos = normalize_products(data)
    if not productos:
        return jsonify({'error': 'Debe ingresar al menos 1 producto'}), 400
    if len(productos) > 6:
        return jsonify({'error': 'Solo se permiten hasta 6 productos'}), 400
    data['productos'] = productos

    try:
        docx_bytes = fill_docx(data)
    except Exception as e:
        return jsonify({'error': 'Error llenando el documento: ' + str(e)}), 500

    # Guardar docx temporal y convertir a PDF con LibreOffice
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, 'carta.docx')
        pdf_path  = os.path.join(tmpdir, 'carta.pdf')

        with open(docx_path, 'wb') as f:
            f.write(docx_bytes)

        # LibreOffice headless conversion
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf',
             '--outdir', tmpdir, docx_path],
            capture_output=True, timeout=60
        )

        if result.returncode != 0 or not os.path.exists(pdf_path):
            return jsonify({
                'error': 'Error convirtiendo a PDF',
                'detail': result.stderr.decode()
            }), 500

        nombre_archivo = 'Carta_Notarial_' + re.sub(r'[^a-zA-Z0-9]','_', data['nombre'])[:30] + '.pdf'

        import io
        with open(pdf_path, 'rb') as f:
            pdf_bytes = f.read()

    # Fuera del TemporaryDirectory — enviamos desde memoria
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype='application/pdf',
        as_attachment=True,
        download_name=nombre_archivo
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
