import streamlit as st
import pandas as pd
from geopy.distance import geodesic
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import folium
from streamlit_folium import st_folium

# =========================================================
# Konfigurasi dasar Streamlit
# =========================================================
st.set_page_config(page_title="📍 Site Nearest Finder", layout="wide")
st.title("📍 Site Nearest Finder")
st.write("Upload file **Mapinfo.xls** (tersimpan di memori). Lalu pilih upload file **ATND.xls** atau input site manual.")

# =========================================================
# 1️⃣ Upload & Simpan Mapinfo
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
    st.success(f"✅ Mapinfo berhasil dimuat ({len(df_map)} site).")
elif st.session_state.mapinfo_df is not None:
    df_map = st.session_state.mapinfo_df
    st.info(f"📂 Menggunakan Mapinfo tersimpan di memori ({len(df_map)} site).")
else:
    st.warning("⚠️ Upload Mapinfo.xls terlebih dahulu.")
    st.stop()

# =========================================================
# 2️⃣ Input: Upload ATND atau Manual
# =========================================================
st.markdown("---")
st.subheader("🗂️ Pilih Cara Input Site yang Akan Dicek")

input_mode = st.radio("Pilih metode input:", ["Upload ATND.xls", "Input Manual"])
df_sites = None

required_cols_sites = ["Site ID", "NE ID", "Sitename", "Longitude", "Latitude"]

if input_mode == "Upload ATND.xls":
    uploaded_sites = st.file_uploader("📂 Upload ATND.xls", type=["xls", "xlsx"])
    if uploaded_sites:
        df_sites = pd.read_excel(uploaded_sites)
        if not set(required_cols_sites).issubset(df_sites.columns):
            st.error(f"Kolom wajib hilang di ATND.xls: {set(required_cols_sites) - set(df_sites.columns)}")
            st.stop()
        st.success(f"✅ File ATND dimuat ({len(df_sites)} site).")

else:
    st.markdown("Masukkan data site manual (kolom sama seperti ATND.xls):")
    if "manual_sites" not in st.session_state:
        st.session_state.manual_sites = pd.DataFrame(columns=required_cols_sites)

    manual_df = st.data_editor(
        st.session_state.manual_sites,
        num_rows="dynamic",
        key="manual_editor",
        use_container_width=True
    )

    df_sites = manual_df.dropna(subset=["Site ID", "Longitude", "Latitude"], how="any")

# =========================================================
# 3️⃣ Tombol Proses
# =========================================================
if df_sites is not None and not df_sites.empty:
    if st.button("🚀 Proses Sekarang"):
        with st.spinner("Menghitung site terdekat..."):
            df_map = st.session_state.mapinfo_df.copy()

            def get_top3_nearest(lat, lon):
                target = (lat, lon)
                df_map["Distance_km"] = df_map.apply(
                    lambda r: geodesic(target, (r["Latitude"], r["Longitude"])).km, axis=1
                )
                nearest = df_map.nsmallest(3, "Distance_km")[["Site ID", "BSC", "Distance_km"]]
                return [
                    f"{nearest.iloc[i]['Site ID']} - {nearest.iloc[i]['BSC']} ({nearest.iloc[i]['Distance_km']:.2f} km)"
                    for i in range(3)
                ]

            nearest_data = [
                get_top3_nearest(row["Latitude"], row["Longitude"]) for _, row in df_sites.iterrows()
            ]
            nearest_df = pd.DataFrame(nearest_data, columns=["Nearest Site 1", "Nearest Site 2", "Nearest Site 3"])
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

            # =========================================================
            # 6️⃣ Peta Satelit Interaktif
            # =========================================================
            st.markdown("---")
            st.subheader("🛰️ Peta Satelit Lokasi Site dan Tetangga Terdekat")

            if not df_sites.empty:
                # Titik tengah peta
                avg_lat = df_sites["Latitude"].mean()
                avg_lon = df_sites["Longitude"].mean()
                m = folium.Map(location=[avg_lat, avg_lon], zoom_start=10, tiles="Esri.WorldImagery")

                # Titik Mapinfo
                for _, row in df_map.iterrows():
                    folium.CircleMarker(
                        location=[row["Latitude"], row["Longitude"]],
                        radius=4,
                        color="lightblue",
                        fill=True,
                        fill_opacity=0.5,
                        popup=f"{row['Site ID']} - {row['BSC']}"
                    ).add_to(m)

                # Titik input (merah)
                for _, row in df_sites.iterrows():
                    folium.Marker(
                        location=[row["Latitude"], row["Longitude"]],
                        popup=f"<b>{row['Site ID']}</b><br>{row['Sitename']}",
                        icon=folium.Icon(color="red", icon="info-sign")
                    ).add_to(m)

                st_folium(m, width=1000, height=600)
else:
    st.info("Silakan upload file ATND atau input manual data site terlebih dahulu.")

st.markdown("---")
st.caption("💾 File Mapinfo tersimpan sementara selama aplikasi aktif. Upload ulang hanya jika ingin memperbarui data.")
