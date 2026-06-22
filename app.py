import re
import os
import uuid  # <-- UPGRADE: Untuk membuat ID unik per user luar
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from docx import Document
from docx.shared import Pt

app = Flask(__name__, template_folder='.')
CORS(app) # Mengizinkan akses dari domain luar (GitHub Pages)

# Buat folder penyimpanan sementara untuk user online
OUTPUT_DIR = "outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def bersihkan_dan_urai_kresna(text):
    matches = list(re.finditer(r"\bBidang\s*:", text))
    text_clean = text[matches[-1].start():] if matches else text
    text_clean = re.sub(r'\[[^\]]+\]\s*KRESNA AI:\s*|KRESNA AI:\s*', '', text_clean)

    data_terurai = {}
    def ambil_pola(pattern, src, flags=0):
        match = re.search(pattern, src, flags)
        return match.group(1).strip() if match else ""

    now = datetime.now()
    hari_list = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    bulan_list = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    
    data_terurai['{{WAKTU}}'] = f"{hari_list[now.weekday()]}, {now.day} {bulan_list[now.month]} {now.year} Pukul {now.strftime('%H.%M')} WIB"
    data_terurai['{{TANGGAL_TTD}}'] = f"{now.day} {bulan_list[now.month]} {now.year}"
    data_terurai['{{BIDANG}}'] = ambil_pola(r"Bidang\s*:\s*(.*)", text_clean)
    data_terurai['{{PERIHAL}}'] = ambil_pola(r"Perihal\s*:\s*(.*)", text_clean)
    data_terurai['{{SUMBER}}'] = ambil_pola(r"A\.\s*Sumber Informasi\s*:\s*(.*)", text_clean)
    data_terurai['{{HUBUNGAN}}'] = ambil_pola(r"B\.\s*Hubungan dengan Sumber\s*:\s*(.*)", text_clean)
    data_terurai['{{CARA}}'] = ambil_pola(r"C\.\s*Cara mendapatkan Informasi\s*:\s*(.*)", text_clean)
    data_terurai['{{NILAI}}'] = ambil_pola(r"E\.\s*Nilai Informasi\s*:\s*(.*)", text_clean)

    fakta_block = ambil_pola(r"II\.\s*FAKTA-FAKTA(.*?)III\.\s*PENDAPAT", text_clean, re.DOTALL)
    if fakta_block:
        data_terurai['{{FAKTA_A}}'] = ambil_pola(r"A\.\s*(.*?)(?=B\.\s*|$)", fakta_block, re.DOTALL)
        data_terurai['{{FAKTA_B}}'] = ambil_pola(r"B\.\s*(.*?)(?=C\.\s*|$)", fakta_block, re.DOTALL)
        data_terurai['{{FAKTA_C}}'] = ambil_pola(r"C\.\s*(.*)", fakta_block, re.DOTALL)
    else:
        data_terurai['{{FAKTA_A}}'] = data_terurai['{{FAKTA_B}}'] = data_terurai['{{FAKTA_C}}'] = ""

    pendapat_block = ambil_pola(r"III\.\s*PENDAPAT PELAPOR(.*)", text_clean, re.DOTALL)
    if pendapat_block:
        abjad = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        for idx, char in enumerate(abjad):
            next_char = abjad[idx+1] if idx+1 < len(abjad) else None
            pattern = rf"{char}\.\s*(.*?)(?={next_char}\.\s*|$)" if next_char else rf"{char}\.\s*(.*)"
            data_terurai[f'{{{{PENDAPAT_{char}}}}}'] = ambil_pola(pattern, pendapat_block, re.DOTALL)
    else:
        for char in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
            data_terurai[f'{{{{PENDAPAT_{char}}}}}'] = ""

    return data_terurai

def suntik_dokumen_docx(p, data_komponen):
    p_text = p.text
    ada_perubahan = False
    if "Waktu mendapatkan Informasi" in p_text and ":" in p_text:
        p_text = f"{p_text.split(':')[0]}: {data_komponen['{{WAKTU}}']}"
        ada_perubahan = True
    elif "Bandar Lampung" in p_text:
        p_text = f"Bandar Lampung, {data_komponen['{{TANGGAL_TTD}}']}"
        ada_perubahan = True
    else:
        for key, val in data_komponen.items():
            if key in p_text:
                p_text = p_text.replace(key, val)
                ada_perubahan = True
    if ada_perubahan:
        is_bold = p.runs[0].bold if p.runs else False
        p.text = "" 
        run = p.add_run(p_text)
        run.font.name = 'Tahoma'
        run.bold = is_bold
        run.font.size = Pt(11)

@app.route('/')
def index():
    return render_template('dashboard.html')

@app.route('/api/generate', methods=['POST'])
def generate_doc():
    data = request.json
    narasi = data.get('text', '')
    storage_path = data.get('storage_path', '').strip()

    if not narasi:
        return jsonify({"status": "error", "message": "Input kosong!"}), 400
        
    komponen_data = bersihkan_dan_urai_kresna(narasi)
    
    # UPGRADE: Generate ID acak unik untuk sesi user ini
    file_id = str(uuid.uuid4())
    output_lokal = os.path.join(OUTPUT_DIR, f"output_{file_id}.docx")
    
    try:
        doc = Document("format_original.docx")
        for p in doc.paragraphs: suntik_dokumen_docx(p, komponen_data)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: suntik_dokumen_docx(p, komponen_data)
                        
        # Simpan ke cloud folder lokal server Anda
        doc.save(output_lokal)
        
        # Simpan ke folder fisik PC Anda jika diisi di dashboard
        if storage_path:
            storage_path = os.path.normpath(storage_path) if os.name == 'nt' else storage_path.replace('\\', '/')
            os.makedirs(storage_path, exist_ok=True)
            doc.save(os.path.join(storage_path, "Laporan_Informasi_Kresna.docx"))
            
        # Kembalikan link unduhan absolut berbasiskan URL Ngrok secara dinamis
        return jsonify({
            "status": "success", 
            "docx_url": f"{request.host_url}download/docx/{file_id}"
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/download/docx/<file_id>')
def download_docx(file_id):
    file_path = os.path.join(OUTPUT_DIR, f"output_{file_id}.docx")
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name="Laporan_Informasi_Kresna.docx")
    return "File kadaluarsa atau tidak ditemukan.", 404

if __name__ == '__main__':
    app.run(debug=True, port=5000)
