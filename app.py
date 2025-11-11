import streamlit as st
import re

def to_int_from_rp(s: str) -> int:
    if not s:
        return 0
    s = s.strip().replace(".", "")
    return int(s) if s.isdigit() else 0

def rp(n: int) -> str:
    return f"{n:,}".replace(",", ".")

# ====== HEADER ======
bulan_tahun = st.text_input(
    "Bulan & Tahun",
    value="", 
    placeholder="contoh: Oktober 2025",
    key="bulan_tahun",
)

bulan_sebelumnya = st.text_input(
    "Nama Bulan Sebelumnya",
    value="", 
    placeholder="contoh: September",
    key="bulan_sebelumnya",
)

saldo_awal_str = st.text_input(
    "Saldo bulan sebelumnya",
    value="", 
    placeholder="contoh: 4.113.000",
    key="saldo_awal",
)
saldo_awal = to_int_from_rp(saldo_awal_str)

# ====== PEMASUKAN ======
st.subheader("Pemasukan (fixed & custom)")

kencleng_kali_str = st.text_input(
    "Kencleng Jumat — jumlah kali",
    value="",
    placeholder="contoh: 5",
    key="kencleng_kali",
)

kencleng_total_str = st.text_input(
    "Kencleng Jumat — total",
    value="",
    placeholder="contoh: 1.316.000",
    key="kencleng_total",
)
kencleng_total = to_int_from_rp(kencleng_total_str)

# Custom pemasukan (10 slot, baris kosong otomatis di-skip)
custom_income = []
for i in range(1, 11):
    c1, c2 = st.columns([2,1])
    desc = c1.text_input(
        f"Uraian pemasukan {i}",
        value="",
        placeholder="contoh: Infaq Warga RW 07 (total) / Infaq Bu Ninin",
        key=f"in_desc_{i}",
    )
    amt_str = c2.text_input(
        f"Nominal pemasukan {i}",
        value="",
        placeholder="contoh: 1.740.000",
        key=f"in_amt_{i}",
    )
    if desc.strip() and amt_str.strip():
        custom_income.append((desc.strip(), to_int_from_rp(amt_str)))

# ====== RINCIAN RW 07 (RT 1-5) ======
st.subheader("Rincian Infaq Warga RW 07 (RT 1–5)")
rt_values = []
for i, hint in enumerate(["290.000","340.000","340.000","370.000","400.000"], start=1):
    rt_str = st.text_input(
        f"RT {i:02d}",
        value="",
        placeholder=f"contoh: {hint}",
        key=f"rt_{i:02d}",
    )
    if rt_str.strip():
        rt_values.append(to_int_from_rp(rt_str))

# ====== PENGELUARAN ======
st.subheader("Pengeluaran")

khotib_kali = st.text_input(
    "Honor Khotib — jumlah Jumat",
    value="",
    placeholder="contoh: 5",
    key="khotib_kali",
)
khotib_total = to_int_from_rp(st.text_input(
    "Honor Khotib — total",
    value="",
    placeholder="contoh: 1.000.000",
    key="honor_khotib",
))

marbot_total = to_int_from_rp(st.text_input(
    "Honor Marbot + Uang Saku",
    value="",
    placeholder="contoh: 1.250.000",
    key="honor_marbot",
))

listrik_kali = st.text_input(
    "Bayar Listrik — jumlah kali",
    value="",
    placeholder="contoh: 2",
    key="listrik_kali",
)
listrik_total = to_int_from_rp(st.text_input(
    "Bayar Listrik — total",
    value="",
    placeholder="contoh: 176.000",
    key="listrik_total",
))

# Custom pengeluaran (15 slot, kosong otomatis di-skip)
custom_expenses = []
for i in range(1, 16):
    c1, c2 = st.columns([2,1])
    desc = c1.text_input(
        f"Uraian pengeluaran {i}",
        value="",
        placeholder="contoh: Fotocopy / Alat perbaikan kipas / Ongkos pembuatan bedug",
        key=f"ex_desc_{i}",
    )
    amt_str = c2.text_input(
        f"Nominal pengeluaran {i}",
        value="",
        placeholder="contoh: 27.500",
        key=f"ex_amt_{i}",
    )
    if desc.strip() and amt_str.strip():
        custom_expenses.append((desc.strip(), to_int_from_rp(amt_str)))
