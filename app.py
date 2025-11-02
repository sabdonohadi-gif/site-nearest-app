import streamlit as st
import pandas as pd
import pydeck as pdk
from geopy.distance import geodesic
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook

st.set_page_config(page_title="Site Nearest Finder", layout="wide")

st.title("📡 Site Nearest Finder (Versi XLS + Manual Input)")
st.write("Upload `Mapinfo.xls` (berisi Site ID, Longitude, Latitude, BSC) dan input site secara manual tanpa perlu file ATND.")

# =======================================
# 1️⃣ Upload / Simpan Mapinfo
# =======================================
if "mapinfo_df" not in st.session_state:
    st.session_state.mapinfo_df = None

st.header("🗺️ Upload Mapinfo (XLS)")
uploaded_mapinfo = st.file_uploader("Pilih file Mapinfo.xls", type=["xls", "xlsx"])

if uploaded_mapinfo:
    df_map = pd.read_excel(uploaded_mapinfo)
    required_cols = {"Site ID", "Longitude", "Latitude", "BSC"}
    if not required_cols.issubset(df_map.columns):
        st.error(f"❌ Kolom wajib hilang di file: {required_cols - set(df_map.columns)}")
    else:
        st.session_state.mapinfo_df = df_map.copy()
        st.success(f"✅ Mapinfo berhasil dimuat ({len(df_map)} site).")
        st.dataframe(df_map.head())
elif st.session_state.mapinfo_df is not None:
    st.info("📂 Menggunakan Mapinfo yang sudah diunggah sebelumnya.")
    df_map = st.session_state.mapinfo_df
else:
    st.warning("⚠️ Harap upload file Mapinfo terlebih dahulu sebelum memproses.")
    st.stop()

# =======================================
# 2️⃣ Input Manual Site(s)
# =======================================
st.header("✏️ Input Manual Site untuk Dicari 3 Terdekat")

with st.form("manual_input_form"):
    st.markdown("Masukkan beberapa site sekaligus (pisahkan baris per site):")
    st.text("Contoh:\nSITE001,106.82,-6.18\nSITE002,106.77,-6.12")
    manual_text = st.text_area("Masukkan data (Site ID, Longitude, Latitude):", height=150)
    submit_manual = st.form_submit_button("✅ Tambahkan Site Manual")

if "manual_sites" not in st.session_state:
    st.session_state.manual_sites = pd.DataFrame(columns=["No", "Site ID", "Longitude", "Latitude"])

if submit_manual and manual_text.strip():
    rows = [r.strip() for r in manual_text.splitlines() if r.strip()]
    parsed_data = []
    for row in rows:
        parts = [p.strip() for p in row.split(",")]
        if len(parts) == 3:
            try:
                parsed_data.append({
                    "Site ID": parts[0],
                    "Longitude": float(parts[1]),
                    "Latitude": float(parts[2])
                })
            except ValueError:
                st.warning(f"⚠️ Format angka salah di baris: {row}")
        else:
            st.warning(f"⚠️ Format salah, gunakan format: SiteID,Longitude,Latitude → {row}")
    if parsed_data:
        new_df = pd.DataFrame(parsed_data)
        st.session_state.manual_sites = pd.concat(
            [st.session_state.manual_sites[["Site ID", "Longitude", "Latitude"]], new_df],
            ignore_index=True
        )
        st.session_state.manual_sites.insert(0, "No", range(1, len(st.session_state.manual_sites) + 1))
        st.success(f"✅ {len(parsed_data)} site berhasil ditambahkan.")

if not st.session_state.manual_sites.empty:
    st.subheader("📋 Daftar Site Input Manual")
    df_show = st.session_state.manual_sites.copy()
    df_show["No"] = range(1, len(df_show) + 1)
    st.dataframe(df_show)

# =======================================
# 3️⃣ Proses Hitung Nearest
# =======================================
if st.button("🚀 Proses Cari 3 Nearest Site"):
    if st.session_state.manual_sites.empty:
        st.error("❌ Harap input minimal 1 site manual terlebih dahulu.")
        st.stop()

    df_map = st.session_state.mapinfo_df.copy()
    df_sites = st.session_state.manual_sites.copy()

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

    nearest_data = []
    for _, row in df_sites.iterrows():
        nearest_data.append(get_top3_nearest(row["Latitude"], row["Longitude"]))

    nearest_df = pd.DataFrame(nearest_data, columns=["Nearest Site 1", "Nearest Site 2", "Nearest Site 3"])
    df_result = pd.concat([df_sites.reset_index(drop=True), nearest_df], axis=1)

    # =======================================
    # 4️⃣ Simpan ke Excel
    # =======================================
    output = BytesIO()
    today = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Nearest_Result_{today}.xlsx"
    df_result.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.success("✅ Proses selesai! Berikut hasilnya:")
    st.dataframe(df_result)

    st.download_button(
        label="⬇️ Download Hasil Excel",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # =======================================
    # 5️⃣ Peta Interaktif (Satelit Mode)
    # =======================================
    markers = []
    lines = []

    for _, row in df_sites.iterrows():
        markers.append({
            "site_id": row["Site ID"],
            "bsc": "-",
            "lon": row["Longitude"],
            "lat": row["Latitude"],
            "type": "input"
        })
        nearest_sites = df_result.loc[df_result["Site ID"] == row["Site ID"], ["Nearest Site 1", "Nearest Site 2", "Nearest Site 3"]].values[0]
        for ns in nearest_sites:
            site_id, bsc_info = ns.split(" - ")
            bsc = bsc_info.split("(")[0].strip()
            nearest_row = df_map[df_map["Site ID"] == site_id]
            if not nearest_row.empty:
                n_row = nearest_row.iloc[0]
                markers.append({
                    "site_id": n_row["Site ID"],
                    "bsc": n_row["BSC"],
                    "lon": n_row["Longitude"],
                    "lat": n_row["Latitude"],
                    "type": "nearest"
                })
                lines.append({
                    "from_lat": row["Latitude"],
                    "from_lon": row["Longitude"],
                    "to_lat": n_row["Latitude"],
                    "to_lon": n_row["Longitude"]
                })

    if markers:
        marker_df = pd.DataFrame(markers)
        marker_df["coordinates"] = marker_df.apply(lambda r: [r["lon"], r["lat"]], axis=1)
        marker_df["color"] = marker_df["type"].apply(lambda t: [0, 116, 217] if t == "input" else [34, 139, 34])

        scatter = pdk.Layer(
            "ScatterplotLayer",
            data=marker_df,
            get_position="coordinates",
            get_fill_color="color",
            get_radius=120,
            radius_min_pixels=5,
            radius_max_pixels=25,
            pickable=True,
            auto_highlight=True,
        )

        if lines:
            lines_df = pd.DataFrame(lines)
            lines_df["from"] = lines_df.apply(lambda r: [r["from_lon"], r["from_lat"]], axis=1)
            lines_df["to"] = lines_df.apply(lambda r: [r["to_lon"], r["to_lat"]], axis=1)
            line_layer = pdk.Layer(
                "LineLayer",
                data=lines_df,
                get_source_position="from",
                get_target_position="to",
                get_width=2,
                get_color=[150, 150, 150],
                pickable=False,
            )
            layers = [scatter, line_layer]
        else:
            layers = [scatter]

        # Map style tetap Satelit
        map_style = "mapbox://styles/mapbox/satellite-v9"

        center_lat = marker_df["lat"].mean()
        center_lon = marker_df["lon"].mean()

        view_state = pdk.ViewState(latitude=center_lat, longitude=center_lon, zoom=10, pitch=0)

        tooltip_html = {
            "html": "<b>Site ID:</b> {site_id}<br><b>BSC:</b> {bsc}",
            "style": {
                "backgroundColor": "white",
                "color": "black",
                "fontSize": "12px",
                "border": "1px solid gray",
                "padding": "4px",
                "borderRadius": "4px",
            },
        }

        r = pdk.Deck(
            layers=layers,
            initial_view_state=view_state,
            tooltip=tooltip_html,
            map_style=map_style,
        )

        st.subheader("🛰️ Peta Interaktif (Mode Satelit) — Hover titik untuk lihat Site ID & BSC")
        st.pydeck_chart(r, use_container_width=True)
