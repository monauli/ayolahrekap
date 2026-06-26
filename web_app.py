"""Versi web (online) dari aplikasi Rekap Excel Booking.

UI berbasis Streamlit yang memanggil logika inti di ``app.py``. Berbeda dari
``desktop_app.py`` (tkinter), versi ini TANPA database: alurnya murni
upload -> proses -> download. Versi desktop/.exe tetap utuh dan tidak terpengaruh.

Jalankan lokal:  streamlit run web_app.py
Deploy:          Streamlit Community Cloud (https://share.streamlit.io)
"""

from __future__ import annotations

import hmac
import os
import tempfile
from datetime import date
from io import BytesIO
from pathlib import Path

import streamlit as st

import app as core

# Konstanta alur "Sebulan otomatis" (disalin dari desktop_app.AUTO_MONTHLY_MIN_DAYS).
AUTO_MONTHLY_MIN_DAYS = 20


# ── Util ────────────────────────────────────────────────────────────────────
def format_rupiah(value) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    return "Rp " + f"{number:,.0f}".replace(",", ".")


def infer_report_period(records, mode_choice: str = "Otomatis"):
    """Tentukan tanggal/bulan laporan dari data (port dari desktop_app, tanpa DB)."""
    detected_dates = core.detect_dates_in_records(records)
    if not detected_dates:
        raise ValueError("Tanggal booking tidak terbaca dari kolom Date of Booking.")

    detected_months = core.detect_months_in_records(records)
    if len(detected_months) > 1:
        months = ", ".join(month.strftime("%Y-%m") for month in detected_months)
        raise ValueError(f"File berisi lebih dari satu bulan ({months}). Pisahkan file per bulan dulu.")

    if len(detected_dates) == 1:
        selected = detected_dates[0]
        return selected, None, None, "Harian", f"Harian otomatis: {selected:%Y-%m-%d}"

    selected_month = detected_months[0]
    if mode_choice == "Sebulan":
        return None, selected_month, None, "Sebulan", f"Sebulan dipilih: {selected_month:%Y-%m}"
    if mode_choice == "Multi Hari":
        return None, selected_month, detected_dates, "Multi Hari", (
            f"Multi Hari: {detected_dates[0]:%Y-%m-%d} s/d {detected_dates[-1]:%Y-%m-%d}"
        )
    if len(detected_dates) >= AUTO_MONTHLY_MIN_DAYS:
        return None, selected_month, None, "Sebulan", f"Sebulan otomatis: {selected_month:%Y-%m}"
    return None, selected_month, detected_dates, "Multi Hari", (
        f"Multi Hari otomatis: {detected_dates[0]:%Y-%m-%d} s/d {detected_dates[-1]:%Y-%m-%d}"
    )


def read_bytes(path: Path) -> bytes:
    return Path(path).read_bytes()


def temp_output(name: str) -> Path:
    return Path(tempfile.gettempdir()) / name


# ── Gerbang password ────────────────────────────────────────────────────────
def get_configured_password() -> tuple[str, bool]:
    """Ambil password dari st.secrets atau env. Bool kedua = pakai default (belum diatur)."""
    try:
        secret = st.secrets.get("app_password")
    except Exception:
        secret = None
    secret = secret or os.environ.get("APP_PASSWORD")
    if secret:
        return secret, False
    return "rekap2026", True


def check_password() -> bool:
    configured, is_default = get_configured_password()

    if st.session_state.get("auth_ok"):
        return True

    st.markdown("### 🔒 Masuk")
    if is_default:
        st.warning(
            "Password masih default (`rekap2026`). Atur `app_password` di Streamlit "
            "Secrets atau env `APP_PASSWORD` sebelum dipakai serius."
        )

    def _submit():
        if hmac.compare_digest(st.session_state.get("pw_input", ""), configured):
            st.session_state["auth_ok"] = True
            st.session_state.pop("pw_input", None)
        else:
            st.session_state["auth_ok"] = False

    st.text_input("Password", type="password", key="pw_input", on_change=_submit)
    if st.session_state.get("auth_ok") is False:
        st.error("Password salah.")
    return bool(st.session_state.get("auth_ok"))


# ── Tab: Rekap Omset ────────────────────────────────────────────────────────
def tab_rekap_omset():
    st.subheader("Rekap Omset Lapangan")
    st.caption("Upload file Excel booking mentah → download hasil rekap Omset Lapangan.")

    uploaded = st.file_uploader("File Excel booking mentah", type=["xlsx"], key="rekap_file")
    mode = st.selectbox(
        "Mode laporan",
        ["Otomatis", "Sebulan", "Multi Hari"],
        help="Otomatis: tebak dari data. Sebulan: rekap satu bulan penuh. Multi Hari: rekap rentang tanggal.",
    )

    if not uploaded:
        return

    try:
        source = core.extract_records(BytesIO(uploaded.getvalue()))
    except Exception as exc:  # noqa: BLE001
        st.error(f"Gagal membaca file: {exc}")
        return

    records = source["records"]
    try:
        report_date, report_month, report_dates, feature, period_message = infer_report_period(records, mode)
    except ValueError as exc:
        st.error(str(exc))
        return

    bk = sum(1 for r in records if str(r.get("Booking ID") or "").upper().startswith("BK/"))
    mn = sum(1 for r in records if str(r.get("Booking ID") or "").upper().startswith("MN/"))
    c1, c2, c3 = st.columns(3)
    c1.metric("Total baris", len(records))
    c2.metric("AYO / Walk In", f"{bk} / {mn}")
    c3.metric("Mode", feature)
    st.info(period_message)

    if not st.button("Buat Rekap Excel", type="primary"):
        return

    with st.spinner("Membuat workbook rekap..."):
        try:
            output_path, stats = core.build_workbook(
                {"headers": core.RAW_HEADERS, "records": records},
                report_date=report_date,
                report_month=report_month,
                report_dates=report_dates,
            )
        except Exception as exc:  # noqa: BLE001
            st.error(f"Gagal memproses: {exc}")
            return

    st.success(f"Selesai — {stats['filename']}")
    m1, m2, m3 = st.columns(3)
    m1.metric("Baris dipakai", stats["included_rows"])
    m2.metric("AYO / Walk In", f"{stats['ayo_rows']} / {stats['walk_in_rows']}")
    m3.metric("Total omset", format_rupiah(stats["total_revenue"]))

    duplicates = stats.get("duplicate_booking_ids") or {}
    if duplicates:
        st.warning(f"{len(duplicates)} Booking ID duplikat tetap dihitung. Cek file sumber bila perlu.")

    st.download_button(
        "⬇️ Download hasil rekap",
        data=read_bytes(output_path),
        file_name=stats["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


# ── Tab: Omset per-Kategori ─────────────────────────────────────────────────
def tab_kategori():
    st.subheader("Omset per-Kategori (Olsera)")
    st.caption("Upload file-file Excel kategori → cek group → download rekap atau olah data gabungan.")

    files = st.file_uploader(
        "File Excel kategori (boleh banyak)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="kategori_files",
    )
    if not files:
        return

    with st.spinner("Mengecek group setiap file..."):
        try:
            results = [core.extract_categories(BytesIO(f.getvalue()), f.name) for f in files]
        except Exception as exc:  # noqa: BLE001
            st.error(f"Gagal membaca: {exc}")
            return

    invalid = {
        str(r["filename"]): sorted(r.get("source_groups") or r.get("categories") or {})
        for r in results
        if len(r.get("source_groups") or r.get("categories") or {}) > 1
    }
    duplicates = core.find_duplicate_categories(results)

    total_rows = sum(r.get("total_rows", 0) for r in results)
    c1, c2 = st.columns(2)
    c1.metric("Jumlah file", len(results))
    c2.metric("Total baris", total_rows)

    if invalid:
        lines = "\n".join(f"- **{name}**: {', '.join(groups)}" for name, groups in invalid.items())
        st.error(
            "File berikut punya lebih dari satu group, jadi tidak bisa diolah. "
            "Pisahkan jadi satu group per file dulu:\n\n" + lines
        )
        return

    if duplicates:
        lines = "\n".join(f"- **{cat}**: {', '.join(fs)}" for cat, fs in duplicates.items())
        st.warning("Group berikut muncul di lebih dari satu file (boleh lanjut, tapi cek lagi):\n\n" + lines)
    else:
        st.success("Cek OK. Semua file 1 group. Siap diolah.")

    col_a, col_b = st.columns(2)

    with col_a:
        if st.button("Buat Rekap Per Kategori", key="btn_kategori"):
            with st.spinner("Menyusun rekap per kategori..."):
                out = core.build_category_workbook(results, duplicates, temp_output("Omset Keseluruhan PERKATEGORI.xlsx"))
            st.session_state["kategori_rekap_bytes"] = read_bytes(out)
        if st.session_state.get("kategori_rekap_bytes"):
            st.download_button(
                "⬇️ Download Rekap Per Kategori",
                data=st.session_state["kategori_rekap_bytes"],
                file_name="Omset Keseluruhan PERKATEGORI.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_kategori",
            )

    with col_b:
        if st.button("Olah Data (gabung semua)", key="btn_olah"):
            with st.spinner("Menggabungkan semua data jadi 1 workbook..."):
                details = [core.read_category_detail(BytesIO(f.getvalue()), f.name) for f in files]
                out = core.build_olah_data_workbook(details, temp_output("Olah Data PERKATEGORI.xlsx"))
            st.session_state["kategori_olah_bytes"] = read_bytes(out)
        if st.session_state.get("kategori_olah_bytes"):
            st.download_button(
                "⬇️ Download Olah Data",
                data=st.session_state["kategori_olah_bytes"],
                file_name="Olah Data PERKATEGORI.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_olah",
            )


# ── Tab: PDF → Excel ────────────────────────────────────────────────────────
def tab_pdf():
    st.subheader("PDF → Excel")
    st.caption("Upload satu atau beberapa PDF → download hasil konversi ke Excel.")

    files = st.file_uploader("File PDF (boleh banyak)", type=["pdf"], accept_multiple_files=True, key="pdf_files")
    if not files:
        return

    if not st.button("Konversi ke Excel", type="primary", key="btn_pdf"):
        return

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        paths = []
        for f in files:
            p = tmp_dir / f.name
            p.write_bytes(f.getvalue())
            paths.append(p)
        out_path = tmp_dir / "Hasil PDF ke Excel.xlsx"
        with st.spinner("Mengonversi PDF..."):
            try:
                output_path, stats = core.convert_pdfs_to_excel(paths, out_path)
            except Exception as exc:  # noqa: BLE001
                st.error(f"Konversi gagal: {exc}")
                return
        data = read_bytes(output_path)

    c1, c2, c3 = st.columns(3)
    c1.metric("Halaman", stats.get("page_count", "-"))
    c2.metric("Tabel", stats.get("table_count", "-"))
    c3.metric("Baris data", stats.get("row_count", "-"))
    st.success("Konversi selesai.")
    st.download_button(
        "⬇️ Download Excel",
        data=data,
        file_name="Hasil PDF ke Excel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


# ── Main ────────────────────────────────────────────────────────────────────
def main():
    st.set_page_config(page_title="Rekap Omset Lapangan", page_icon="🏟️", layout="centered")
    st.title("🏟️ Rekap Omset Lapangan")

    if not check_password():
        st.stop()

    tab1, tab2, tab3 = st.tabs(["Rekap Omset", "Omset per-Kategori", "PDF → Excel"])
    with tab1:
        tab_rekap_omset()
    with tab2:
        tab_kategori()
    with tab3:
        tab_pdf()


if __name__ == "__main__":
    main()
