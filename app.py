import os
import uuid
import base64
import subprocess
import tempfile
import time
import zipfile
import requests
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ─────────────────────────────────────────────────────────────
# Sesiones compartidas (ADMIN/CLIENTE) – en memoria
# Nota: si el servicio se reinicia, se pierden. (Sin caducidad por petición)
# ─────────────────────────────────────────────────────────────
SESSIONS = {}  # sid -> {"ver": int, "data": dict, "created_at": int, "updated_at": int}

def _sid_ok(sid: str) -> bool:
    return isinstance(sid, str) and 1 <= len(sid) <= 64


# ─── Carta Notarial — llenado de DOCX ───────────────────────
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

def make_cell(text, width):
    return (
        '<w:tc><w:tcPr><w:tcW w:w="' + str(width) + '" w:type="dxa"/></w:tcPr>'
        '<w:p w:rsidR="00B136EE" w:rsidRDefault="00B136EE" w:rsidP="00B136EE">'
        '<w:pPr><w:jc w:val="both"/>'
        '<w:rPr><w:rFonts w:ascii="Lato" w:eastAsia="Lato" w:hAnsi="Lato" w:cs="Lato"/>'
        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
        + make_run(text) +
        '</w:p></w:tc>'
    )

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
            + make_run(data['representante']) +
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

    # ── 8. Llenar primera fila de datos de la tabla ────────
    tbl_start = doc.index('<w:tbl>')
    tbl_end   = doc.index('</w:tbl>') + 8
    tbl_xml   = doc[tbl_start:tbl_end]

    rows = tbl_xml.split('<w:tr ')
    # rows[2] = primera fila de datos (vacía)
    old_data_row = '<w:tr ' + rows[2]
    new_data_row = (
        '<w:tr w:rsidR="00B136EE" w:rsidTr="00B136EE">'
        + make_cell(data['tipo_prod'], 2942)
        + make_cell(data['ncontrato'], 2943)
        + make_cell('S/ ' + data['monto'], 2943)
        + '</w:tr>'
    )
    new_tbl = tbl_xml.replace(old_data_row, new_data_row, 1)
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


@app.route('/session/save', methods=['POST'])
def session_save():
    try:
        payload = request.get_json(silent=True) or {}
        sid = payload.get('sid')
        data = payload.get('data')
        base_ver = payload.get('base_ver')
        if not _sid_ok(sid):
            return jsonify({'ok': False, 'error': 'sid inválido'}), 400
        if not isinstance(data, dict):
            return jsonify({'ok': False, 'error': 'data inválido'}), 400

        now = int(time.time())
        cur = SESSIONS.get(sid, {'ver': 0, 'data': {}, 'created_at': now, 'updated_at': now})
        cur_ver = int(cur.get('ver', 0))
        # Si el cliente intenta guardar sobre una versión vieja, igual aceptamos (last-write-wins)
        new_ver = cur_ver + 1
        SESSIONS[sid] = {
            'ver': new_ver,
            'data': data,
            'created_at': int(cur.get('created_at', now)),
            'updated_at': now,
        }
        return jsonify({'ok': True, 'sid': sid, 'ver': new_ver, 'updated_at': now})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/session/get', methods=['GET'])
def session_get():
    try:
        sid = request.args.get('sid', '')
        if not _sid_ok(sid):
            return jsonify({'ok': False, 'error': 'sid inválido'}), 400
        cur = SESSIONS.get(sid)
        if not cur:
            return jsonify({'ok': True, 'sid': sid, 'ver': 0, 'data': None})
        return jsonify({
            'ok': True,
            'sid': sid,
            'ver': int(cur.get('ver', 0)),
            'updated_at': int(cur.get('updated_at', 0)),
            'data': cur.get('data'),
        })
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/session/list', methods=['GET'])
def session_list():
    try:
        try:
            limit = int(request.args.get('limit', '15'))
        except Exception:
            limit = 15
        if limit < 1:
            limit = 1
        if limit > 50:
            limit = 50

        items = []
        for sid, cur in SESSIONS.items():
            try:
                data = cur.get('data') or {}
                cn = (data.get('cn') or '').strip()
                # Solo listar sesiones que ya tienen nombre (para que el cliente vea "Nombre", no códigos)
                if not cn:
                    continue
                items.append({
                    'sid': sid,
                    'cn': cn,
                    'ver': int(cur.get('ver', 0)),
                    'updated_at': int(cur.get('updated_at', 0)),
                    'created_at': int(cur.get('created_at', 0)),
                })
            except Exception:
                continue
        items.sort(key=lambda x: x.get('updated_at', 0), reverse=True)
        return jsonify({'ok': True, 'sessions': items[:limit]})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/session/clear', methods=['POST'])
def session_clear():
    try:
        payload = request.get_json(silent=True) or {}
        sid = payload.get('sid')
        if not _sid_ok(sid):
            return jsonify({'ok': False, 'error': 'sid inválido'}), 400
        SESSIONS.pop(sid, None)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route("/", methods=["GET","HEAD"])
def root():
    return jsonify({"status": "ok"})

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/convert', methods=['POST'])
def convert():
    try:
        data = request.get_json()
        if not data or 'xlsx_b64' not in data:
            return jsonify({'error': 'Falta el campo xlsx_b64'}), 400

        xlsx_bytes = base64.b64decode(data['xlsx_b64'])

        tmp_id = str(uuid.uuid4())
        xlsx_path = f'/tmp/{tmp_id}.xlsx'
        pdf_path  = f'/tmp/{tmp_id}.pdf'

        with open(xlsx_path, 'wb') as f:
            f.write(xlsx_bytes)

        result = subprocess.run(
            ['libreoffice', '--headless', '--norestore',
             '--convert-to', 'pdf',
             '--outdir', '/tmp',
             xlsx_path],
            capture_output=True, text=True, timeout=60
        )

        if result.returncode != 0 or not os.path.exists(pdf_path):
            os.remove(xlsx_path)
            return jsonify({
                'error': 'Error al convertir',
                'detalle': result.stderr
            }), 500

        with open(pdf_path, 'rb') as f:
            pdf_b64 = base64.b64encode(f.read()).decode('ascii')

        os.remove(xlsx_path)
        os.remove(pdf_path)

        return jsonify({'pdf_b64': pdf_b64})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/telegram', methods=['POST'])
def telegram():
    try:
        data = request.get_json()
        if not data or not data.get('token') or not data.get('chat_id'):
            return jsonify({'ok': False, 'error': 'Faltan token o chat_id'}), 400

        token   = data['token']
        chat_id = data['chat_id']
        text    = data.get('text', '')
        pdf_b64 = data.get('pdf_b64')
        filename = data.get('filename', 'informe.pdf')

        # 1) Enviar mensaje de texto
        requests.post(
            f'https://api.telegram.org/bot{token}/sendMessage',
            json={'chat_id': chat_id, 'text': text},
            timeout=10
        )

        # 2) Enviar PDF si viene
        if pdf_b64 and filename:
            pdf_bytes = base64.b64decode(pdf_b64)
            files = {'document': (filename, pdf_bytes, 'application/pdf')}
            requests.post(
                f'https://api.telegram.org/bot{token}/sendDocument',
                data={'chat_id': chat_id, 'caption': text[:200]},
                files=files,
                timeout=30
            )

        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500



@app.route('/generar-pdf', methods=['POST'])
def generar_pdf():
    data = request.get_json(force=True)

    required = ['fecha','nombre','direccion','cod_ofic','nom_ofic',
                'tipo_prod','ncontrato','monto','telefono']
    for k in required:
        if not data.get(k,'').strip():
            return jsonify({'error': f'Campo requerido: {k}'}), 400

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

        # LibreOffice headless conversion (mismo método que /convert)
        result = subprocess.run(
            ['libreoffice', '--headless', '--norestore',
             '--convert-to', 'pdf',
             '--outdir', tmpdir, docx_path],
            capture_output=True, text=True, timeout=60
        )

        if result.returncode != 0 or not os.path.exists(pdf_path):
            return jsonify({
                'error': 'Error convirtiendo a PDF',
                'detail': result.stderr
            }), 500

        nombre_archivo = 'Carta_Notarial_' + re.sub(r'[^a-zA-Z0-9]','_', data['nombre'])[:30] + '.pdf'

        return send_file(
            pdf_path,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=nombre_archivo
        )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
