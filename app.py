import os
import sys
import uuid
from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
from docx import Document
from fpdf import FPDF
import io
import time

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'FINAL_SATS_ASS.docx')

# Check template exists at startup
if not os.path.exists(TEMPLATE_PATH):
    print(f"ERROR: Template file FINAL_SATS_ASS.docx not found at {TEMPLATE_PATH}", file=sys.stderr)
    print("Make sure the file is committed to git for deployment.", file=sys.stderr)
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'generated')
_cleanup_done = False

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── DOCX Modification ───────────────────────────────────────────────────────

def modify_docx_template(student_name, student_pid, student_course, output_path=None):
    """Modify the DOCX template by replacing student name, PID, and adding course."""
    doc = Document(TEMPLATE_PATH)

    if doc.tables and len(doc.tables) > 0:
        table = doc.tables[0]
        cell = table.rows[0].cells[0]

        pid_para_index = None
        for pi, para in enumerate(cell.paragraphs):
            for run in para.runs:
                if 'N. Akshit Vinay' in run.text:
                    run.text = run.text.replace('N. Akshit Vinay', student_name)
                if '25MSRSGIS001' in run.text:
                    run.text = run.text.replace('25MSRSGIS001', student_pid)
                    pid_para_index = pi

        # Add student course as a new paragraph at the end of the cell
        if student_course and pid_para_index is not None:
            cell.add_paragraph(f'Course: {student_course}')

    if output_path:
        doc.save(output_path)
        return output_path

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def escape_html(text):
    """Escape HTML special characters."""
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;'))


def docx_to_html_preview(docx_path):
    """Extract DOCX content as HTML preview."""
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
                if cell_text:
                    html_parts.append(f'<td>{escape_html(cell_text)}</td>')
                else:
                    html_parts.append('<td></td>')
            html_parts.append('</tr>')
        html_parts.append('</table>')

    return '\n'.join(html_parts)


def find_unicode_font():
    """Find a Unicode-capable font on the system (Windows or Linux)."""
    # Windows font paths
    win_fonts = r'C:\Windows\Fonts'
    # Linux font paths
    linux_fonts = '/usr/share/fonts'

    candidates = [
        # Windows
        (os.path.join(win_fonts, 'arial.ttf'), 'Arial'),
        (os.path.join(win_fonts, 'Calibri.ttf'), 'Calibri'),
        (os.path.join(win_fonts, 'segoeui.ttf'), 'Segoe UI'),
        (os.path.join(win_fonts, 'times.ttf'), 'Times New Roman'),
        (os.path.join(win_fonts, 'tahoma.ttf'), 'Tahoma'),
        # Linux - DejaVu (commonly available)
        (os.path.join(linux_fonts, 'truetype/dejavu/DejaVuSans.ttf'), 'DejaVuSans'),
        (os.path.join(linux_fonts, 'truetype/dejavu/DejaVuSansCondensed.ttf'), 'DejaVuSansCondensed'),
        # Linux - Ubuntu
        (os.path.join(linux_fonts, 'truetype/ubuntu/Ubuntu.ttf'), 'Ubuntu'),
        # Linux - Liberation (commonly installed)
        (os.path.join(linux_fonts, 'truetype/liberation/LiberationSans-Regular.ttf'), 'LiberationSans'),
        # Generic fallback paths
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


def generate_pdf_from_docx(docx_path, output_path):
    """Generate a PDF from the DOCX file preserving text content."""
    doc = Document(docx_path)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)

    # Find a Unicode-capable font
    font_path, family_name = find_unicode_font()
    font_name = 'Helvetica'
    has_bold = True

    if font_path:
        try:
            pdf.add_font('DocFont', '', font_path)
            # Try to find a bold variant
            dirname = os.path.dirname(font_path)
            basename = os.path.basename(font_path)
            bold_variants = [
                basename.replace('.ttf', 'bd.ttf'),
                basename.replace('.ttf', 'b.ttf'),
                basename.replace('Sans.ttf', 'Sans-Bold.ttf'),
                basename.replace('SansCondensed.ttf', 'SansCondensed-Bold.ttf'),
                basename.replace('arial.ttf', 'arialbd.ttf'),
                basename.replace('Regular.ttf', 'Bold.ttf'),
            ]
            bold_found = False
            for bv in bold_variants:
                bp = os.path.join(dirname, bv)
                if os.path.exists(bp):
                    try:
                        pdf.add_font('DocFont', 'B', bp)
                        bold_found = True
                        break
                    except Exception:
                        continue

            has_bold = bold_found
            font_name = 'DocFont'
        except Exception as e:
            print(f"Font add error: {e}", file=sys.stderr)
            font_name = 'Helvetica'
            has_bold = True

    pdf.add_page()

    # Process paragraphs
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style_name = para.style.name if para.style else ''
        is_bold = any(run.bold for run in para.runs if run.text.strip())

        if 'Title' in style_name:
            pdf.set_font(font_name, 'B' if has_bold else '', 18)
        elif 'Heading 1' in style_name:
            pdf.set_font(font_name, 'B' if has_bold else '', 16)
        elif 'Heading 2' in style_name:
            pdf.set_font(font_name, 'B' if has_bold else '', 14)
        elif 'Heading 3' in style_name:
            pdf.set_font(font_name, 'B' if has_bold else '', 12)
        elif is_bold:
            pdf.set_font(font_name, 'B' if has_bold else '', 11)
        else:
            pdf.set_font(font_name, '', 11)

        try:
            pdf.multi_cell(0, 6, text)
            pdf.ln(2)
        except Exception:
            try:
                safe_text = text.encode('utf-8', errors='replace').decode('utf-8')
                pdf.multi_cell(0, 6, safe_text)
                pdf.ln(2)
            except Exception:
                pass

    # Process tables
    for table in doc.tables:
        if not table.rows:
            continue
        if pdf.get_y() > pdf.h - 40:
            pdf.add_page()

        col_width = 190 / max(len(table.columns), 1)

        for ri, row in enumerate(table.rows):
            is_header = (ri == 0)
            pdf.set_font(font_name, 'B' if (is_header and has_bold) else '', 8)

            for cell in row.cells:
                text = cell.text.strip()[:200]
                try:
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
        _cleanup_done = True
    return render_template('index.html')


def _cleanup_old_files():
    """Clean up files older than 30 minutes."""
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
    """Generate the assignment with user details."""
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
        modify_docx_template(student_name, student_pid, student_course, docx_path)

        pdf_ready = False
        try:
            generate_pdf_from_docx(docx_path, pdf_path)
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                pdf_ready = True
        except Exception as e:
            print(f"PDF generation error: {e}", file=sys.stderr)

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
    """Download the generated DOCX file."""
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
    """Download the generated PDF file."""
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
