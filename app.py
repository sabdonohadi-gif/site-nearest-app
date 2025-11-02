# ==========================
# 🌍 Build Pydeck Map (interaktif)
# ==========================
if markers:
    # deduplicate markers by (site_id, lon, lat, type, sitename)
    unique_markers = {(m["site_id"], m["lon"], m["lat"], m["type"], m["sitename"]): m for m in markers}
    marker_list = list(unique_markers.values())

    marker_df = pd.DataFrame(marker_list)

    # pydeck expects [lon, lat] for positions
    marker_df["coordinates"] = marker_df.apply(lambda r: [r["lon"], r["lat"]], axis=1)

    # color column: input = blue, nearest = green
    marker_df["color"] = marker_df["type"].apply(lambda t: [0, 116, 217] if t == "input" else [34, 139, 34])

    # Scatterplot Layer (titik)
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

    # Line Layer (garis penghubung)
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

    # Otomatis cari pusat peta dari semua titik
    center_lat = marker_df["lat"].mean()
    center_lon = marker_df["lon"].mean()

    view_state = pdk.ViewState(
        latitude=center_lat,
        longitude=center_lon,
        zoom=11,
        pitch=0,
    )

    # Tooltip HTML rapi
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

    r = pdk.Deck(
        layers=layers,
        initial_view_state=view_state,
        tooltip=tooltip_html,
        map_style="mapbox://styles/mapbox/light-v9",
    )

    st.subheader("🗺️ Peta Interaktif — Hover titik untuk lihat Site ID")
    st.pydeck_chart(r, use_container_width=True)
else:
    st.info("Tidak ada marker untuk ditampilkan di peta.")
