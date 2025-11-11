from flask import Flask, request, render_template_string, send_file
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO

app = Flask(__name__)

HTML = r"""
<!doctype html>
<html>
<head>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>Laporan DKM → .docx</title>
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:800px;margin:24px auto;padding:0 12px}
    fieldset{border:1px solid #ccc;border-radius:8px;margin:16px 0;padding:12px}
    legend{font-weight:600}
    label{display:block;margin:8px 0 4px}
    input[type=text],input[type=number]{width:100%;padding:8px;border:1px solid #bbb;border-radius:6px}
    .row{display:flex;gap:8px}
    .row>div{flex:1}
    button{padding:8px 12px;border:1px solid #555;border-radius:8px;background:#111;color:#fff}
    .ghost{background:#eee;color:#111;border-color:#ddd}
    small{color:#666}
  </style>
  <script>
    function addPair(sectionId, prefix){
      const box = document.getElementById(sectionId);
      const idx = box.querySelectorAll('.pair').length + 1;
      const wrap = document.createElement('div');
      wrap.className = 'pair';
      wrap.innerHTML = `
        <div class="row">
          <div><label>Deskripsi</label><input name="${prefix}_desc_${idx}" type="text" placeholder="mis. Infaq Bu Ninin"></div>
          <div><label>Jumlah (rupiah, titik sebagai pemisah)</label><input name="${prefix}_amt_${idx}" type="text" placeholder="100.000"></div>
        </div>`;
      box.appendChild(wrap);
    }
  </script>
</head>
<body>
  <h2>Laporan Keuangan DKM → unduh Word (.docx)</h2>
  <form method="post" action="/build">
    <fieldset>
      <legend>Header & Saldo</legend>
      <div class="row">
        <div><label>Bulan & Tahun (contoh: Oktober 2025)</label><input name="bulan_tahun" required></div>
        <div><label>Nama Bulan Sebelumnya (contoh: September)</label><input name="bulan_sebelumnya" required></div>
      </div>
      <label>Saldo awal (rupiah)</label>
      <input name="saldo_awal" placeholder="4.113.000" required>
    </fieldset>

    <fieldset>
      <legend>Pemasukan</legend>
      <div class="row">
        <div><label>Kencleng Jumat: berapa kali?</label><input name="kencleng_kali" type="number" min="0" value="0"></div>
        <div><label>Kencleng Jumat: total (rupiah)</label><input name="kencleng_total" placeholder="1.316.000"></div>
      </div>

      <label>Infaq Warga RW 07 (Total)</label>
      <input name="rw07_total" placeholder="1.740.000">

      <div class="row">
        <div><label>RT 01</label><input name="rt_01" placeholder="290.000"></div>
        <div><label>RT 02</label><input name="rt_02" placeholder="340.000"></div>
        <div><label>RT 03</label><input name="rt_03" placeholder="340.000"></div>
      </div>
      <div class="row">
        <div><label>RT 04</label><input name="rt_04" placeholder="370.000"></div>
        <div><label>RT 05</label><input name="rt_05" placeholder="400.000"></div>
      </div>

      <hr>
      <div id="incomeBox"></div>
      <button type="button" class="ghost" onclick="addPair('incomeBox','income')">+ Tambah baris pemasukan custom</button>
      <div><small>Baris yang dikosongkan tidak akan dimasukkan.</small></div>
    </fieldset>

    <fieldset>
      <legend>Pengeluaran</legend>
      <div class="row">
        <div><label>Honor Khotib: berapa kali Jumat?</label><input name="jumat_khotib" type="number" min="0" value="0"></div>
        <div><label>Honor Khotib (total rupiah)</label><input name="honor_khotib" placeholder="1.000.000"></div>
      </div>
      <label>Honor Marbot + Uang Saku (rupiah)</label>
      <input name="honor_marbot_uangsaku" placeholder="1.250.000">
      <div class="row">
        <div><label>Listrik: berapa kali?</label><input name="listrik_kali" type="number" min="0" value="0"></div>
        <div><label>Bayar Listrik (rupiah)</label><input name="bayar_listrik" placeholder="176.000"></div>
      </div>

      <hr>
      <div id="expenseBox"></div>
      <button type="button" class="ghost" onclick="addPair('expenseBox','expense')">+ Tambah baris pengeluaran custom</button>
      <div><small>Kosong = di-skip. Rapih tanpa baris kosong.</small></div>
    </fieldset>

    <fieldset>
      <legend>Tanda Tangan</legend>
      <div class="row">
        <div><label>Tanggal TTD (contoh: 31 Oktober 2025)</label><input name="tanggal_ttd"></div>
        <div><label>Nama Ketua DKM</label><input name="ttd_ketua" value="Ali Marga"></div>
        <div><label>Nama Bendahara</label><input name="ttd_bendahara" value="Eneng Nariah"></div>
      </div>
    </fieldset>

    <button type="submit">Buat & Unduh .docx</button>
  </form>
</body>
</html>
"""

def to_int(val: str) -> int:
    if not val: return 0
    s = val.replace(".", "").replace(",", "").strip()
    return int(s) if s.isdigit() else 0

def rupiah(n: int) -> str:
    return f"{n:,}".replace(",", ".")

def add_row(tbl, left, right):
    row = tbl.add_row().cells
    row[0].text = left
    row[1].text = right

def base_doc():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    style.font.size = Pt(11)
    return doc

@app.get("/")
def form():
    return render_template_string(HTML)

@app.post("/build")
def build():
    f = request.form

    # Header
    bulan_tahun = f.get("bulan_tahun","")
    bulan_sebelum = f.get("bulan_sebelumnya","")
    saldo_awal_str = f.get("saldo_awal","")
    saldo_awal = to_int(saldo_awal_str)

    # Income fixed
    kencleng_kali = f.get("kencleng_kali","")
    kencleng_total = to_int(f.get("kencleng_total",""))
    rw07_total = to_int(f.get("rw07_total",""))
    rt_vals = [f.get("rt_01",""), f.get("rt_02",""), f.get("rt_03",""),
               f.get("rt_04",""), f.get("rt_05","")]

    # Income custom (pair fields)
    income_pairs = []
    for k,v in f.items():
        # names like income_desc_1 / income_amt_1
        pass
    # Build pairs in order:
    i = 1
    while True:
        desc = f.get(f"income_desc_{i}","").strip()
        amt  = f.get(f"income_amt_{i}","").strip()
        if not desc and not amt:
            # stop when neither provided for a while
            if i>50: break
            i += 1
            continue
        if desc and amt:
            income_pairs.append((desc, to_int(amt)))
        i += 1

    # Expense fixed
    jumat_khotib = f.get("jumat_khotib","")
    honor_khotib = to_int(f.get("honor_khotib",""))
    honor_marbot = to_int(f.get("honor_marbot_uangsaku",""))
    listrik_kali = f.get("listrik_kali","")
    bayar_listrik = to_int(f.get("bayar_listrik",""))

    # Expense custom
    expense_pairs = []
    j = 1
    while True:
        desc = f.get(f"expense_desc_{j}","").strip()
        amt  = f.get(f"expense_amt_{j}","").strip()
        if not desc and not amt:
            if j>80: break
            j += 1
            continue
        if desc and amt:
            expense_pairs.append((desc, to_int(amt)))
        j += 1

    # Build DOCX
    doc = base_doc()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Laporan Keuangan DKM Sirojul Huda"); r.bold=True; r.font.size=Pt(12)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.add_run("Aspol Sukamiskin Bandung")
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.add_run(f"Bulan {bulan_tahun}")

    doc.add_paragraph(f"Saldo {bulan_sebelum}").runs[0].bold = True
    p = doc.add_paragraph(f"{saldo_awal_str}"); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Pemasukan
    doc.add_paragraph(); doc.add_paragraph("Pemasukan").runs[0].bold = True
    t_in = doc.add_table(rows=1, cols=2); t_in.autofit=True; t_in.rows[0].cells[0].text=""; t_in.rows[0].cells[1].text=""

    total_in = 0
    if kencleng_kali or kencleng_total:
        add_row(t_in, f"Kencleng Jumat ({kencleng_kali}x)" if kencleng_kali else "Kencleng Jumat", rupiah(kencleng_total))
        total_in += kencleng_total

    if rw07_total:
        add_row(t_in, "Infaq Warga RW 07 (total)", rupiah(rw07_total))
        total_in += rw07_total

    for desc, amt in income_pairs:
        add_row(t_in, desc, rupiah(amt))
        total_in += amt

    add_row(t_in, "Total Pemasukan", rupiah(total_in))

    # Rincian RT (hanya yang diisi)
    doc.add_paragraph(); doc.add_paragraph("Rincian Infaq Warga RW 07 (berdasarkan RT)").runs[0].bold = True
    t_rt = doc.add_table(rows=1, cols=2); t_rt.autofit=True; t_rt.rows[0].cells[0].text=""; t_rt.rows[0].cells[1].text=""
    for i, val in enumerate(rt_vals, start=1):
        val = val.strip()
        if val:
            add_row(t_rt, f"RT {i:02d}", val)

    # Pemasukan kotor
    doc.add_paragraph(); doc.add_paragraph("Pemasukan Kotor").runs[0].bold = True
    gross = saldo_awal + total_in
    p = doc.add_paragraph(rupiah(gross)); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Pengeluaran
    doc.add_paragraph(); doc.add_paragraph("Pengeluaran").runs[0].bold = True
    t_out = doc.add_table(rows=1, cols=2); t_out.autofit=True; t_out.rows[0].cells[0].text=""; t_out.rows[0].cells[1].text=""
    total_out = 0

    if jumat_khotib or honor_khotib:
        add_row(t_out, f"Honor Khotib ({jumat_khotib}x)".strip(), rupiah(honor_khotib))
        total_out += honor_khotib
    if honor_marbot:
        add_row(t_out, "Honor Marbot + Uang Saku", rupiah(honor_marbot))
        total_out += honor_marbot
    if listrik_kali or bayar_listrik:
        add_row(t_out, f"Bayar Listrik ({listrik_kali}x)".strip(), rupiah(bayar_listrik))
        total_out += bayar_listrik

    for desc, amt in expense_pairs:
        add_row(t_out, desc, rupiah(amt))
        total_out += amt

    add_row(t_out, "Total Pengeluaran", rupiah(total_out))

    # Saldo akhir
    doc.add_paragraph(); doc.add_paragraph(f"Saldo {bulan_tahun}").runs[0].bold = True
    end_bal = gross - total_out
    p = doc.add_paragraph(rupiah(end_bal)); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Footer / TTD
    doc.add_paragraph()
    row = doc.add_table(rows=1, cols=2).rows[0].cells
    row[0].paragraphs[0].add_run("Ketua DKM Sirojul Huda")
    right = row[1].paragraphs[0]; right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right.add_run(f"Bandung, {f.get('tanggal_ttd','')}")

    row2 = doc.add_table(rows=1, cols=2).rows[0].cells
    row2[0].paragraphs[0].add_run(f.get("ttd_ketua",""))
    r2 = row2[1].paragraphs[0]; r2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r2.add_run(f.get("ttd_bendahara",""))

    # Return .docx
    buf = BytesIO()
    doc.save(buf); buf.seek(0)
    filename = f"Laporan_DKM_{bulan_tahun.replace(' ','_')}.docx" if bulan_tahun else "Laporan_DKM.docx"
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
