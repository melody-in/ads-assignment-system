import os
import sys
import json
import uuid
from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
from docx import Document
from fpdf import FPDF
import io
import time

app = Flask(__name__)
CORS(app)

# ─── Configuration ───────────────────────────────────────────────────────────

BASE_DIR = os.path.dirname(__file__)
DOCUMENTS_DIR = os.path.join(BASE_DIR, 'documents')
SETTINGS_PATH = os.path.join(DOCUMENTS_DIR, 'settings.json')
OUTPUT_DIR = os.path.join(BASE_DIR, 'generated')
_cleanup_done = False

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(DOCUMENTS_DIR, exist_ok=True)


def load_settings():
    """Load settings from documents/settings.json, creating defaults if missing."""
    defaults = {
        "template": "",
        "findName": "N. Akshit Vinay",
        "findPid": "25MSRSGIS001",
        "findCourse": "M.Sc. GIS & Remote Sensing"
    }

    if not os.path.exists(SETTINGS_PATH):
        with open(SETTINGS_PATH, 'w') as f:
            json.dump(defaults, f, indent=2)
        print(f"[INFO] Created default settings at {SETTINGS_PATH}", file=sys.stderr)
        return defaults

    try:
        with open(SETTINGS_PATH, 'r') as f:
            settings = json.load(f)
            for key in defaults:
                if key not in settings:
                    settings[key] = defaults[key]
            return settings
    except Exception as e:
        print(f"[WARN] Could not read settings: {e}", file=sys.stderr)
        return defaults


def find_template():
    """Find the DOCX template in the documents/ folder."""
    settings = load_settings()
    template_name = settings.get('template', '')

    # If template is specified in settings, use it
    if template_name:
        path = os.path.join(DOCUMENTS_DIR, template_name)
        if os.path.exists(path):
            return path

    # Otherwise, look for any .docx file in the documents/ folder
    docx_files = [f for f in os.listdir(DOCUMENTS_DIR)
                  if f.lower().endswith('.docx') and f != 'settings.json']
    if docx_files:
        # Use the most recently modified one
        docx_files.sort(key=lambda f: os.path.getmtime(os.path.join(DOCUMENTS_DIR, f)), reverse=True)
        chosen = docx_files[0]
        # Auto-update settings
        settings['template'] = chosen
        try:
            with open(SETTINGS_PATH, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception:
            pass
        return os.path.join(DOCUMENTS_DIR, chosen)

    # Fallback: look in project root
    root_docx = [f for f in os.listdir(BASE_DIR) if f.lower().endswith('.docx')]
    if root_docx:
        return os.path.join(BASE_DIR, root_docx[0])

    return None


# ─── DOCX Modification (Pure Find-and-Replace) ───────────────────────────────

def modify_docx_template(student_name, student_pid, student_course, output_path=None):
    """
    Modify the DOCX template by performing a pure find-and-replace.
    Only the EXACT strings specified in settings are replaced — nothing else changes.
    """
    template_path = find_template()
    if not template_path:
        raise FileNotFoundError("No DOCX template found. Place a .docx file in the 'documents/' folder.")

    settings = load_settings()
    find_name = settings.get('findName', 'N. Akshit Vinay')
    find_pid = settings.get('findPid', '25MSRSGIS001')
    find_course = settings.get('findCourse', 'M.Sc. GIS & Remote Sensing')

    doc = Document(template_path)

    # ── Replace in all paragraphs ──
    for para in doc.paragraphs:
        for run in para.runs:
            if find_name in run.text:
                run.text = run.text.replace(find_name, student_name)
            if find_pid in run.text:
                run.text = run.text.replace(find_pid, student_pid)
            if find_course in run.text:
                run.text = run.text.replace(find_course, student_course)

    # ── Replace in all tables ──
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if find_name in run.text:
                            run.text = run.text.replace(find_name, student_name)
                        if find_pid in run.text:
                            run.text = run.text.replace(find_pid, student_pid)
                        if find_course in run.text:
                            run.text = run.text.replace(find_course, student_course)

    if output_path:
        doc.save(output_path)
        return output_path

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ─── HTML Preview ────────────────────────────────────────────────────────────

def escape_html(text):
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;'))


def docx_to_html_preview(docx_path):
    """Extract DOCX content as HTML preview (preserves all original content)."""
    doc = Document(docx_path)
    html_parts = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style_name = para.style.name if para.style else ''
        safe_text = escape_html(text)

        if 'Heading 1' in style_name or 'Title' in style_name:
            html_parts.append(f'<h1>{safe_text}</h1>')
        elif 'Heading 2' in style_name:
            html_parts.append(f'<h2>{safe_text}</h2>')
        elif 'Heading 3' in style_name:
            html_parts.append(f'<h3>{safe_text}</h3>')
        else:
            is_bold = any(run.bold for run in para.runs if run.text.strip())
            if is_bold and len(text) < 100:
                html_parts.append(f'<p class="strong">{safe_text}</p>')
            else:
                html_parts.append(f'<p>{safe_text}</p>')

    for table in doc.tables:
        html_parts.append('<table>')
        for row in table.rows:
            html_parts.append('<tr>')
            for cell in row.cells:
                cell_text = cell.text.strip()
                html_parts.append(f'<td>{escape_html(cell_text) if cell_text else ""}</td>')
            html_parts.append('</tr>')
        html_parts.append('</table>')

    return '\n'.join(html_parts)


# ─── PDF Generation ──────────────────────────────────────────────────────────

def find_unicode_font():
    """Find a Unicode-capable font (prefer Times New Roman for academic docs)."""
    win_fonts = r'C:\Windows\Fonts'
    linux_fonts = '/usr/share/fonts'

    candidates = [
        # Windows - Times New Roman preferred
        (os.path.join(win_fonts, 'times.ttf'), 'TimesNewRoman'),
        (os.path.join(win_fonts, 'timesbd.ttf'), 'TimesNewRoman-Bold'),
        (os.path.join(win_fonts, 'timesi.ttf'), 'TimesNewRoman-Italic'),
        # Windows fallbacks
        (os.path.join(win_fonts, 'arial.ttf'), 'Arial'),
        (os.path.join(win_fonts, 'Calibri.ttf'), 'Calibri'),
        # Linux fallbacks
        (os.path.join(linux_fonts, 'truetype/dejavu/DejaVuSans.ttf'), 'DejaVuSans'),
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    ]

    for candidate in candidates:
        if isinstance(candidate, tuple):
            fpath, fname = candidate
        else:
            fpath = candidate
            fname = None
        if os.path.exists(fpath):
            return fpath, fname or os.path.splitext(os.path.basename(fpath))[0]

    return None, None


def find_bold_font(font_dir, base_name):
    """Try to find a bold variant of a font file."""
    bold_variants = [
        base_name.replace('.ttf', 'bd.ttf'),
        base_name.replace('.ttf', 'b.ttf'),
        base_name.replace('.ttf', '-Bold.ttf'),
        base_name.replace('Sans.ttf', 'Sans-Bold.ttf'),
        base_name.replace('SansCondensed.ttf', 'SansCondensed-Bold.ttf'),
        base_name.replace('times.ttf', 'timesbd.ttf'),
        base_name.replace('times.ttf', 'Times_New_Roman_Bold.ttf'),
        base_name.replace('Regular.ttf', 'Bold.ttf'),
    ]
    for bv in bold_variants:
        bp = os.path.join(font_dir, bv)
        if os.path.exists(bp):
            return bp
    return None


def generate_pdf_from_docx(docx_path, output_path):
    """Generate a PDF from the DOCX file using Times New Roman."""
    doc = Document(docx_path)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)

    # Find a Unicode font (prefer Times New Roman)
    font_path, family_name = find_unicode_font()
    font_name = 'Helvetica'

    if font_path:
        try:
            pdf.add_font('DocFont', '', font_path, uni=True)

            # Try to find bold variant
            bold_path = find_bold_font(os.path.dirname(font_path), os.path.basename(font_path))
            if bold_path:
                try:
                    pdf.add_font('DocFont', 'B', bold_path, uni=True)
                    font_name = 'DocFont'
                except Exception:
                    font_name = 'DocFont'
            else:
                font_name = 'DocFont'
        except Exception as e:
            print(f"[PDF] Font load error: {e}", file=sys.stderr)

    pdf.add_page()

    # Process paragraphs
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style_name = para.style.name if para.style else ''
        is_bold = any(run.bold for run in para.runs if run.text.strip())

        if 'Title' in style_name:
            pdf.set_font(font_name, 'B' if font_name == 'DocFont' else '', 16)
        elif 'Heading 1' in style_name:
            pdf.set_font(font_name, 'B' if font_name == 'DocFont' else '', 14)
        elif 'Heading 2' in style_name:
            pdf.set_font(font_name, 'B' if font_name == 'DocFont' else '', 12)
        elif 'Heading 3' in style_name:
            pdf.set_font(font_name, 'B' if font_name == 'DocFont' else '', 11)
        elif is_bold:
            pdf.set_font(font_name, 'B' if font_name == 'DocFont' else '', 10)
        else:
            pdf.set_font(font_name, '', 10)

        try:
            pdf.multi_cell(0, 5.5, text)
            pdf.ln(1.5)
        except Exception:
            try:
                safe_text = text.encode('utf-8', errors='replace').decode('utf-8')
                pdf.multi_cell(0, 5.5, safe_text)
                pdf.ln(1.5)
            except Exception:
                pass

    # Process tables
    for table in doc.tables:
        if not table.rows:
            continue
        if pdf.get_y() > pdf.h - 40:
            pdf.add_page()

        col_count = max(len(table.columns), 1)
        col_width = 190 / col_count

        for ri, row in enumerate(table.rows):
            is_header = (ri == 0)
            bold_style = 'B' if font_name == 'DocFont' and is_header else ''

            for cell in row.cells:
                text = cell.text.strip()[:200]
                try:
                    pdf.set_font(font_name, bold_style, 7.5)
                    pdf.cell(col_width, 5, text, border=1)
                except Exception:
                    try:
                        safe_text = text.encode('utf-8', errors='replace').decode('utf-8')
                        pdf.cell(col_width, 5, safe_text, border=1)
                    except Exception:
                        pass
            pdf.ln()

    pdf.output(output_path)
    return output_path


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    global _cleanup_done
    if not _cleanup_done:
        _cleanup_old_files()

        # Check template exists
        tmpl = find_template()
        if tmpl:
            print(f"[OK] Template found: {tmpl}", file=sys.stderr)
        else:
            print(f"[WARN] No DOCX template found. Place a .docx file in the 'documents/' folder.", file=sys.stderr)

        _cleanup_done = True
    return render_template('index.html')


def _cleanup_old_files():
    now = time.time()
    for f in os.listdir(OUTPUT_DIR):
        fpath = os.path.join(OUTPUT_DIR, f)
        if os.path.isfile(fpath) and now - os.path.getmtime(fpath) > 1800:
            try:
                os.remove(fpath)
            except Exception:
                pass


@app.route('/generate', methods=['POST'])
def generate_assignment():
    data = request.get_json()
    student_name = data.get('studentName', '').strip()
    student_pid = data.get('studentPid', '').strip()
    student_course = data.get('studentCourse', '').strip()

    if not student_name or not student_pid or not student_course:
        return jsonify({'error': 'Please provide student name, PID, and course name.'}), 400

    file_id = uuid.uuid4().hex[:12]
    docx_path = os.path.join(OUTPUT_DIR, f'assignment_{file_id}.docx')
    pdf_path = os.path.join(OUTPUT_DIR, f'assignment_{file_id}.pdf')

    try:
        # Check template
        tmpl = find_template()
        if not tmpl:
            return jsonify({'error': 'No DOCX template found. Place a .docx file in the "documents/" folder.'}), 400

        modify_docx_template(student_name, student_pid, student_course, docx_path)

        pdf_ready = False
        try:
            generate_pdf_from_docx(docx_path, pdf_path)
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                pdf_ready = True
        except Exception as e:
            print(f"[PDF] Generation error: {e}", file=sys.stderr)

        preview_html = docx_to_html_preview(docx_path)

        return jsonify({
            'success': True,
            'previewHtml': preview_html,
            'docxUrl': f'/download-docx/{file_id}',
            'pdfUrl': f'/download-pdf/{file_id}',
            'pdfReady': pdf_ready,
            'studentName': student_name,
            'studentPid': student_pid,
            'studentCourse': student_course
        })

    except Exception as e:
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500


@app.route('/download-docx/<file_id>')
def download_docx(file_id):
    docx_path = os.path.join(OUTPUT_DIR, f'assignment_{file_id}.docx')
    if not os.path.exists(docx_path):
        return jsonify({'error': 'File not found'}), 404

    return send_file(
        docx_path,
        as_attachment=True,
        download_name='Assignment.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


@app.route('/download-pdf/<file_id>')
def download_pdf(file_id):
    pdf_path = os.path.join(OUTPUT_DIR, f'assignment_{file_id}.pdf')
    if not os.path.exists(pdf_path):
        return jsonify({'error': 'PDF not generated yet. Please regenerate the assignment.'}), 404

    return send_file(
        pdf_path,
        as_attachment=True,
        download_name='Assignment.pdf',
        mimetype='application/pdf'
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', debug=False, port=port)
