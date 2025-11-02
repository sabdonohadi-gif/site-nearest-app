import streamlit as st
import pandas as pd
import numpy as np
import pydeck as pdk
from math import radians, sin, cos, sqrt, atan2

# ================================================
# 🧭 Fungsi untuk hitung jarak antar koordinat (Haversine)
# ================================================
def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0  # radius bumi dalam km
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlon, dlat = lon2 - lon1, lat2 - lat1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))
    return R * c


# ================================================
# 🚀 Streamlit UI
# ================================================
st.set_page_config(page_title="Nearest Site Finder", layout="wide")

st.title("📍 Nearest Site Finder + Mapinfo View")
st.write("Upload file Mapinfo (CSV) dan cari 3 site terdekat berdasarkan koordinat input atau file ATND.")

# ================================================
# 📁 Upload Mapinfo
# ================================================
uploaded_mapinfo = st.file_uploader("📂 Upload Mapinfo CSV", type=["csv"])
mapinfo_df = None

if uploaded_mapinfo:
    try:
        mapinfo_df = pd.read_csv(uploaded_mapinfo)
        mapinfo_df.columns = map(str.lower, mapinfo_df.columns)
        # pastikan kolom wajib ada
        required_cols = {"site_id", "lat", "lon"}
        if not required_cols.issubset(set(mapinfo_df.columns)):
            st.error("❌ File Mapinfo harus punya kolom: site_id, lat, lon")
            mapinfo_df = None
        else:
            st.success(f"✅ Mapinfo berhasil dimuat ({len(mapinfo_df)} baris)")
    except Exception as e:
        st.error(f"Gagal membaca file Mapinfo: {e}")

# ================================================
# 🧾 Input Site (manual atau file ATND)
# ================================================
st.divider()
st.subheader("🛰️ Input Site")

input_option = st.radio("Pilih sumber input site:", ["Manual", "Upload ATND (CSV)"])

input_sites = pd.DataFrame()

if input_option == "Manual":
    site_id = st.text_input("Masukkan Site ID (contoh: TNG001):")
    lat = st.number_input("Latitude:", format="%.6f")
    lon = st.number_input("Longitude:", format="%.6f")
    if site_id and lat and lon:
        input_sites = pd.DataFrame([{"site_id": site_id, "lat": lat, "lon": lon}])
elif input_option == "Upload ATND (CSV)":
    uploaded_atnd = st.file_uploader("📄 Upload ATND CSV", type=["csv"])
    if uploaded_atnd:
        try:
            input_sites = pd.read_csv(uploaded_atnd)
            input_sites.columns = map(str.lower, input_sites.columns)
            if not {"site_id", "lat", "lon"}.issubset(set(input_sites.columns)):
                st.error("❌ File ATND harus punya kolom: site_id, lat, lon")
                input_sites = pd.DataFrame()
            else:
                st.success(f"✅ ATND berhasil dimuat ({len(input_sites)} baris)")
        except Exception as e:
            st.error(f"Gagal membaca file ATND: {e}")

# ================================================
# 🔍 Proses cari nearest site
# ================================================
if mapinfo_df is not None and not input_sites.empty:
    if st.button("🚀 Proses Cari 3 Nearest Site"):
        nearest_results = []
        markers = []
        lines = []

        for _, row in input_sites.iterrows():
            site_id = row["site_id"]
            lat, lon = row["lat"], row["lon"]

            mapinfo_df["distance_km"] = mapinfo_df.apply(
                lambda x: haversine(lat, lon, x["lat"], x["lon"]), axis=1
            )
            nearest = mapinfo_df.nsmallest(3, "distance_km")

            for _, nrow in nearest.iterrows():
                nearest_results.append({
                    "input_site": site_id,
                    "nearest_site": nrow["site_id"],
                    "distance_km": round(nrow["distance_km"], 3)
                })
                lines.append({
                    "from_lon": lon,
                    "from_lat": lat,
                    "to_lon": nrow["lon"],
                    "to_lat": nrow["lat"]
                })

            # marker input
            markers.append({"site_id": site_id, "lat": lat, "lon": lon, "type": "input", "sitename": site_id})
            # marker nearest
            for _, nrow in nearest.iterrows():
                markers.append({
                    "site_id": nrow["site_id"],
                    "lat": nrow["lat"],
                    "lon": nrow["lon"],
                    "type": "nearest",
                    "sitename": nrow.get("sitename", nrow["site_id"])
                })

        nearest_df = pd.DataFrame(nearest_results)
        st.success("✅ Proses selesai!")
        st.dataframe(nearest_df)

        # ================================================
        # 🌍 Peta Interaktif dengan Tooltip + Theme
        # ================================================
        st.divider()
        st.subheader("🗺️ Peta Interaktif")

        style_option = st.selectbox("🗺️ Pilih gaya peta:", ["Light", "Dark", "Satellite"])
        map_styles = {
            "Light": "mapbox://styles/mapbox/light-v9",
            "Dark": "mapbox://styles/mapbox/dark-v9",
            "Satellite": "mapbox://styles/mapbox/satellite-v9"
        }

        if st.button("🔍 Zoom ke semua titik"):
            zoom_auto = True
        else:
            zoom_auto = False

        # deduplicate markers
        unique_markers = {(m["site_id"], m["lat"], m["lon"], m["type"], m["sitename"]): m for m in markers}
        marker_df = pd.DataFrame(unique_markers.values())
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

        # line layer
        lines_df = pd.DataFrame(lines)
        if not lines_df.empty:
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

        # zoom center
        center_lat, center_lon = marker_df["lat"].mean(), marker_df["lon"].mean()
        view_state = pdk.ViewState(latitude=center_lat, longitude=center_lon, zoom=11 if zoom_auto else 10, pitch=0)

        tooltip_html = {
            "html": "<b>Site ID:</b> {site_id}<br><b>Name/BSC:</b> {sitename}",
            "style": {
                "backgroundColor": "white",
                "color": "black",
                "fontSize": "12px",
                "border": "1px solid gray",
                "padding": "4px",
                "borderRadius": "4px",
            },
        }

        deck = pdk.Deck(
            layers=layers,
            initial_view_state=view_state,
            tooltip=tooltip_html,
            map_style=map_styles[style_option],
        )

        st.pydeck_chart(deck, use_container_width=True)

else:
    st.info("Silakan upload Mapinfo dan input site terlebih dahulu.")
