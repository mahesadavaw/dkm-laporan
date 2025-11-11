import io
import re
import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

st.set_page_config(page_title="Laporan DKM – DOCX", layout="wide")

# ---------- Helpers ----------
def only_digits(s: str) -> int:
    if not s:
        return 0
    return int(re.sub(r"[^0-9]", "", s))

def rupiah(n: int) -> str:
    # format 1.000.000
    s = f"{n:,}".replace(",", ".")
    return s

def add_row(tbl, left, right):
    r = tbl.add_row().cells
    r[0].text = left
    r[1].text = right

def docx_build(
    bulan_tahun, bulan_sebelumnya, saldo_awal_rp,
    kencleng_kali, kencleng_total_rp,
    pemasukan_custom, rt_breakdown,
    khotib_kali, khotib_rp, marbot_rp, listrik_kali, listrik_rp,
    pengeluaran_custom,
    tgl_ttd, ketua, bendahara
):
    doc = Document()

    # Font & spacing padat
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    style.font.size = Pt(11)
    pf = style.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.line_spacing = 1

    # Margin kecil biar 1 lembar
    sec = doc.sections[0]
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.8)
    sec.left_margin = Cm(1.8)
    sec.right_margin = Cm(1.8)

    # Judul
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Laporan Keuangan DKM Sirojul Huda")
    run.bold = True
    run.font.size = Pt(12)

    p = doc.add_paragraph("Aspol Sukamiskin Bandung")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Bulan {bulan_tahun}").alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Saldo awal
    doc.add_paragraph()
    doc.add_paragraph(f"Saldo {bulan_sebelumnya}").runs[0].bold = True
    p = doc.add_paragraph(rupiah(saldo_awal_rp))
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # --- Pemasukan
    doc.add_paragraph()
    doc.add_paragraph("Pemasukan").runs[0].bold = True
    tbl_in = doc.add_table(rows=1, cols=2)
    tbl_in.rows[0].cells[0].text = ""
    tbl_in.rows[0].cells[1].text = ""

    total_pemasukan = 0

    # Kencleng (fixed)
    if kencleng_kali or kencleng_total_rp:
        add_row(tbl_in, f"Kencleng Jumat ({kencleng_kali}x)" if kencleng_kali else "Kencleng Jumat", rupiah(kencleng_total_rp))
        total_pemasukan += kencleng_total_rp

    # Custom pemasukan (skip kosong)
    for desc, amt in pemasukan_custom:
        if desc.strip() and amt > 0:
            add_row(tbl_in, desc.strip(), rupiah(amt))
            total_pemasukan += amt

    add_row(tbl_in, "Total Pemasukan", rupiah(total_pemasukan))

    # Rincian RW 07 (informasi)
    doc.add_paragraph()
    doc.add_paragraph("Rincian Infaq Warga RW 07 (berdasarkan RT)").runs[0].bold = True
    tbl_rt = doc.add_table(rows=1, cols=2)
    tbl_rt.rows[0].cells[0].text = ""
    tbl_rt.rows[0].cells[1].text = ""
    for label, val in rt_breakdown:
        if val > 0:
            add_row(tbl_rt, label, rupiah(val))

    # Pemasukan kotor
    doc.add_paragraph()
    doc.add_paragraph("Pemasukan Kotor").runs[0].bold = True
    p = doc.add_paragraph(rupiah(saldo_awal_rp + total_pemasukan))
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # --- Pengeluaran
    doc.add_paragraph()
    doc.add_paragraph("Pengeluaran").runs[0].bold = True
    tbl_out = doc.add_table(rows=1, cols=2)
    tbl_out.rows[0].cells[0].text = ""
    tbl_out.rows[0].cells[1].text = ""

    total_pengeluaran = 0

    if khotib_kali or khotib_rp:
        add_row(tbl_out, f"Honor Khotib ({khotib_kali}x)" if khotib_kali else "Honor Khotib", rupiah(khotib_rp))
        total_pengeluaran += khotib_rp
    if marbot_rp:
        add_row(tbl_out, "Honor Marbot + Uang Saku", rupiah(marbot_rp))
        total_pengeluaran += marbot_rp
    if listrik_kali or listrik_rp:
        add_row(tbl_out, f"Bayar Listrik ({listrik_kali}x)" if listrik_kali else "Bayar Listrik", rupiah(listrik_rp))
        total_pengeluaran += listrik_rp

    for desc, amt in pengeluaran_custom:
        if desc.strip() and amt > 0:
            add_row(tbl_out, desc.strip(), rupiah(amt))
            total_pengeluaran += amt

    add_row(tbl_out, "Total Pengeluaran", rupiah(total_pengeluaran))

    # Saldo akhir
    doc.add_paragraph()
    doc.add_paragraph(f"Saldo {bulan_tahun}").runs[0].bold = True
    saldo_akhir = saldo_awal_rp + total_pemasukan - total_pengeluaran
    p = doc.add_paragraph(rupiah(saldo_akhir))
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # TTD
    doc.add_paragraph()
    t1 = doc.add_table(rows=1, cols=2).rows[0].cells
    left = t1[0].paragraphs[0]
    right = t1[1].paragraphs[0]; right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right.add_run(f"Bandung, {tgl_ttd}")

    t2 = doc.add_table(rows=1, cols=2).rows[0].cells
    l = t2[0].paragraphs[0]; r = t2[1].paragraphs[0]
    l.add_run("Ketua DKM Sirojul Huda")
    r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r.add_run("Bendahara DKM")

    # ruang tanda tangan
    doc.add_paragraph(); doc.add_paragraph(); doc.add_paragraph()

    t3 = doc.add_table(rows=1, cols=2).rows[0].cells
    l = t3[0].paragraphs[0]; r = t3[1].paragraphs[0]
    l.add_run(ketua)
    r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r.add_run(bendahara)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ---------- UI ----------
st.markdown("### Laporan Keuangan DKM — Generator DOCX (rapi 1 lembar)")

colA, colB, colC = st.columns([1.1, 1.2, 1.0])

with colA:
    bulan_tahun = st.text_input("Bulan & Tahun", placeholder="Oktober 2025")
    bulan_sebelumnya = st.text_input("Nama Bulan Sebelumnya", placeholder="September")
    saldo_awal = only_digits(st.text_input("Saldo September (Rp)", placeholder="4.113.000"))

with colB:
    st.markdown("#### Pemasukan (fixed & custom)")
    kencleng_kali = only_digits(st.text_input("Kencleng Jumat — jumlah kali", placeholder="5"))
    kencleng_total = only_digits(st.text_input("Kencleng Jumat — total (Rp)", placeholder="1.316.000"))

    st.caption("Custom Pemasukan (kosongkan baris yang tidak dipakai)")
    pemasukan_custom = []
    for i in range(1, 11):
        c1, c2 = st.columns([1.4, 0.6])
        desc = c1.text_input(f"Uraian pemasukan {i}", placeholder="Mis. Infaq Warga RW 07 (total)" if i==1 else "", key=f"in_desc_{i}")
        amt = only_digits(c2.text_input(f"Nominal pemasukan {i}", placeholder="1.740.000" if i==1 else "", key=f"in_amt_{i}"))
        pemasukan_custom.append((desc, amt))

with colC:
    st.markdown("#### Rincian RW 07 (RT 1–5)")
    rt_breakdown = []
    for rt, ph in zip(["RT 01","RT 02","RT 03","RT 04","RT 05"], ["290.000","340.000","340.000","370.000","400.000"]):
        v = only_digits(st.text_input(rt, placeholder=ph, key=f"rt_{rt[-2:]}"))
        rt_breakdown.append((rt, v))

st.divider()
colD, colE = st.columns([1.2, 1.2])

with colD:
    st.markdown("#### Pengeluaran (fixed & custom)")
    c1, c2 = st.columns([1.2, 0.8])
    khotib_kali = only_digits(c1.text_input("Jumlah khotib (x Jumat)", placeholder="5"))
    khotib_rp = only_digits(c2.text_input("Honor Khotib (Rp)", placeholder="1.000.000"))

    c3, c4 = st.columns([1.2, 0.8])
    _ = c3.text_input("Marbot (keterangan)", value="Honor Marbot + Uang Saku", disabled=True)
    marbot_rp = only_digits(c4.text_input("Marbot (Rp)", placeholder="1.250.000"))

    c5, c6 = st.columns([1.2, 0.8])
    listrik_kali = only_digits(c5.text_input("Bayar Listrik — jumlah kali", placeholder="2"))
    listrik_rp = only_digits(c6.text_input("Listrik (Rp)", placeholder="176.000"))

with colE:
    pengeluaran_custom = []
    st.caption("Custom Pengeluaran (kosongkan baris yang tidak dipakai)")
    for i in range(1, 16):
        c1, c2 = st.columns([1.4, 0.6])
        desc = c1.text_input(f"Uraian pengeluaran {i}", placeholder="", key=f"ex_desc_{i}")
        amt = only_digits(c2.text_input(f"Nominal pengeluaran {i}", placeholder="", key=f"ex_amt_{i}"))
        pengeluaran_custom.append((desc, amt))

st.divider()
cL, cM, cR = st.columns([1, 1, 1])
with cL:
    tgl_ttd = st.text_input("Tanggal untuk tanda tangan", placeholder="31 Oktober 2025")
with cM:
    ketua = st.text_input("Nama Ketua DKM", placeholder="Ali Marga")
with cR:
    bendahara = st.text_input("Nama Bendahara", placeholder="Eneng Nariah")

if st.button("Unduh DOCX"):
    buf = docx_build(
        bulan_tahun, bulan_sebelumnya, saldo_awal,
        kencleng_kali, kencleng_total,
        pemasukan_custom, rt_breakdown,
        khotib_kali, khotib_rp, marbot_rp, listrik_kali, listrik_rp,
        pengeluaran_custom,
        tgl_ttd, ketua, bendahara
    )
    st.download_button(
        "Download Laporan (DOCX)",
        data=buf.getvalue(),
        file_name=f"Laporan_DKM_{bulan_tahun or 'Bulan'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
