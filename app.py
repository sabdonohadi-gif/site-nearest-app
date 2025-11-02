import streamlit as st
import pandas as pd
from geopy.distance import geodesic
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="📍 Site Nearest Finder", layout="centered")
st.title("📍 Site Nearest Finder")
st.write("Upload file **Mapinfo.xls** sekali saja. Lalu kamu bisa pilih: upload file **ATND.xls** atau input site manual.")
st.write("Mapinfo akan tersimpan di memori, jadi kamu tidak perlu upload ulang setiap kali — hanya saat ingin memperbarui datanya.")

# =========================================================
# 1️⃣ Upload & Simpan Mapinfo.xls
# =========================================================
if "mapinfo_df" not in st.session_state:
    st.session_state.mapinfo_df = None

uploaded_mapinfo = st.file_uploader("📂 Upload / Update Mapinfo.xls", type=["xls", "xlsx"])

if uploaded_mapinfo:
    df_map = pd.read_excel(uploaded_mapinfo)
    required_cols_map = {"Site ID", "Longitude", "Latitude", "BSC"}
    if not required_cols_map.issubset(df_map.columns):
        st.error(f"Kolom wajib hilang di Mapinfo.xls: {required_cols_map - set(df_map.columns)}")
        st.stop()
    st.session_state.mapinfo_df = df_map.copy()
    st.success(f"✅ Mapinfo berhasil dimuat ({len(df_map)} site). Data tersimpan di memori.")
elif st.session_state.mapinfo_df is not None:
    df_map = st.session_state.mapinfo_df
    st.info(f"📂 Menggunakan Mapinfo tersimpan di memori ({len(df_map)} site).")
else:
    st.warning("⚠️ Harap upload file Mapinfo.xls terlebih dahulu sebelum memproses.")
    st.stop()

# =========================================================
# 2️⃣ Pilihan input: Upload atau Manual
# =========================================================
st.markdown("---")
st.subheader("🗂️ Pilih Cara Input Site yang Akan Dicek")
input_mode = st.radio("Pilih metode input:", ["Upload ATND.xls", "Input Manual"])

df_sites = None

if input_mode == "Upload ATND.xls":
    uploaded_sites = st.file_uploader("📂 Upload ATND.xls", type=["xls", "xlsx"])
    if uploaded_sites:
        df_sites = pd.read_excel(uploaded_sites)
        st.success(f"✅ File ATND dimuat ({len(df_sites)} site).")

else:
    st.markdown("Masukkan data site secara manual:")
    if "manual_sites" not in st.session_state:
        st.session_state.manual_sites = pd.DataFrame(
            columns=["No", "Site ID", "Longitude", "Latitude"]
        )

    manual_df = st.data_editor(
        st.session_state.manual_sites,
        num_rows="dynamic",
        key="manual_editor",
        use_container_width=True
    )

    # Update nomor urut otomatis
    if not manual_df.empty:
        manual_df["No"] = range(1, len(manual_df) + 1)

    df_sites = manual_df.dropna(subset=["Site ID", "Longitude", "Latitude"], how="any")

# =========================================================
# 3️⃣ Tombol Proses
# =========================================================
if df_sites is not None and not df_sites.empty:
    if st.button("🚀 Proses Sekarang"):
        with st.spinner("Menghitung jarak site terdekat..."):
            df_map = st.session_state.mapinfo_df.copy()
            required_cols_sites = {"Site ID", "Longitude", "Latitude"}

            if not required_cols_sites.issubset(df_sites.columns):
                st.error(f"Kolom wajib hilang di input site: {required_cols_sites - set(df_sites.columns)}")
                st.stop()

            def get_top3_nearest(lat, lon):
                target = (lat, lon)
                df_map["Distance_km"] = df_map.apply(
                    lambda row: geodesic(target, (row["Latitude"], row["Longitude"])).km, axis=1
                )
                nearest = df_map.nsmallest(3, "Distance_km")[["Site ID", "BSC", "Distance_km"]]
                return [
                    f"{nearest.iloc[i]['Site ID']} - {nearest.iloc[i]['BSC']} ({nearest.iloc[i]['Distance_km']:.2f} km)"
                    for i in range(3)
                ]

            nearest_data = [
                get_top3_nearest(row["Latitude"], row["Longitude"]) for _, row in df_sites.iterrows()
            ]
            nearest_df = pd.DataFrame(
                nearest_data,
                columns=["Nearest Site 1", "Nearest Site 2", "Nearest Site 3"]
            )

            df_result = pd.concat([df_sites.reset_index(drop=True), nearest_df], axis=1)

            # =========================================================
            # 4️⃣ Format Excel hasil
            # =========================================================
            output = BytesIO()
            df_result.to_excel(output, index=False, engine="openpyxl")
            output.seek(0)

            wb = load_workbook(output)
            ws = wb.active
            ws.freeze_panes = "A2"

            header_fill = PatternFill(start_color="0B3D91", end_color="0B3D91", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            header_align = Alignment(horizontal="center", vertical="center")
            thin = Side(border_style="thin", color="000000")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align
                cell.border = border

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    cell.border = border

            for col in ws.columns:
                max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(10, max_length + 2)

            buf = BytesIO()
            wb.save(buf)
            buf.seek(0)

            # =========================================================
            # 5️⃣ Output & Download
            # =========================================================
            st.success("✅ Proses selesai! Klik tombol di bawah untuk mengunduh hasilnya.")
            today = datetime.now().strftime("%Y%m%d")
            st.download_button(
                label="📥 Download Hasil Excel",
                data=buf,
                file_name=f"ATND_Result_{today}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.dataframe(df_result, use_container_width=True)

else:
    st.info("Silakan upload file ATND atau input manual data site terlebih dahulu.")

# =========================================================
# 6️⃣ Info tambahan
# =========================================================
st.markdown("---")
st.caption("💾 File Mapinfo tersimpan sementara selama aplikasi masih aktif. Upload ulang hanya jika ingin memperbarui data Mapinfo.")
