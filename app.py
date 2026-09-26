import io
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(
    page_title="GRSM 8 — Koordinat → Area & Alamat",
    page_icon="📍",
    layout="wide",
)

st.title("📍 GRSM 8 — Koordinat → Area & Alamat")
st.caption("Upload Excel → proses Lat/Long → isi Area & ALAMAT → download Excel.")

st.info(
    "File yang kamu upload sudah kamu filter hanya untuk GRSM 8, "
    "jadi aplikasi ini akan memproses SEMUA baris yang memiliki koordinat "
    "di sheet Exclusive M1 dan Exclusive M3. Baris tanpa Lat/Long akan dilewati."
)

# ------------------------------------------------------------------
# Reverse geocoding
# ------------------------------------------------------------------

@st.cache_data(show_spinner=False, ttl=60 * 60 * 24)
def geocode_photon(lat, lon):
    """Reverse geocode one coordinate using Photon/OpenStreetMap."""
    url = "https://photon.komoot.io/reverse"
    headers = {
        "User-Agent": "GRSM8-Coordinate-Geocoder/2.0"
    }

    for attempt in range(4):
        try:
            response = requests.get(
                url,
                params={"lat": float(lat), "lon": float(lon)},
                headers=headers,
                timeout=25,
            )

            if response.status_code == 429:
                time.sleep(2 ** attempt)
                continue

            response.raise_for_status()
            data = response.json()

            features = data.get("features", [])
            if not features:
                return {
                    "area": "",
                    "alamat": "",
                    "status": "TIDAK DITEMUKAN",
                }

            props = features[0].get("properties", {})

            street = (
                props.get("street")
                or props.get("name")
                or ""
            )
            housenumber = props.get("housenumber") or ""

            if street and housenumber:
                alamat = f"{street} No. {housenumber}"
            else:
                alamat = street

            # Photon can return city/county/district depending on location.
            # For this workbook we want kabupaten/kota.
            area = (
                props.get("city")
                or props.get("county")
                or props.get("district")
                or props.get("state_district")
                or ""
            )

            return {
                "area": str(area),
                "alamat": str(alamat),
                "status": "OK" if (area or alamat) else "TIDAK LENGKAP",
            }

        except Exception as exc:
            if attempt < 3:
                time.sleep(1.5 * (attempt + 1))
            else:
                return {
                    "area": "",
                    "alamat": "",
                    "status": f"ERROR: {type(exc).__name__}",
                }

    return {
        "area": "",
        "alamat": "",
        "status": "ERROR: rate limit",
    }


def geocode_google(lat, lon, api_key):
    """Reverse geocode one coordinate using Google Maps Geocoding API."""
    url = "https://maps.googleapis.com/maps/api/geocode/json"

    response = requests.get(
        url,
        params={
            "latlng": f"{float(lat)},{float(lon)}",
            "key": api_key,
            "language": "id",
        },
        timeout=25,
    )
    response.raise_for_status()
    data = response.json()

    if data.get("status") != "OK" or not data.get("results"):
        return {
            "area": "",
            "alamat": "",
            "status": data.get("status", "TIDAK DITEMUKAN"),
        }

    result = data["results"][0]
    alamat = result.get("formatted_address", "")
    area = ""

    for component in result.get("address_components", []):
        types = component.get("types", [])
        if "administrative_area_level_2" in types:
            area = component.get("long_name", "")
            break

    if not area:
        for component in result.get("address_components", []):
            types = component.get("types", [])
            if "locality" in types:
                area = component.get("long_name", "")
                break

    return {
        "area": area,
        "alamat": alamat,
        "status": "OK",
    }


def clean_coord(value):
    try:
        if value is None or str(value).strip() == "":
            return None
        return round(float(value), 7)
    except (ValueError, TypeError):
        return None


def get_columns(ws):
    headers = []
    for cell in ws[1]:
        headers.append(
            str(cell.value).strip().upper()
            if cell.value is not None
            else ""
        )
    return {header: idx + 1 for idx, header in enumerate(headers)}


def process_workbook(uploaded_file, provider, google_key=None):
    workbook = load_workbook(
        io.BytesIO(uploaded_file.getvalue()),
        data_only=False,
    )

    # Process these sheets when present. If one is absent, simply skip it.
    sheet_names = [
        name for name in ["Exclusive M1", "Exclusive M3"]
        if name in workbook.sheetnames
    ]

    if not sheet_names:
        # Fallback: process every sheet if the workbook has different names.
        sheet_names = workbook.sheetnames

    # First pass: collect rows and unique coordinates.
    jobs = []
    coordinate_set = set()
    skipped_no_coord = 0
    already_complete = 0

    for sheet_name in sheet_names:
        ws = workbook[sheet_name]
        cols = get_columns(ws)

        missing = [
            c for c in ["AREA", "LAT", "LONG", "ALAMAT"]
            if c not in cols
        ]
        if missing:
            continue

        for row_num in range(2, ws.max_row + 1):
            lat = clean_coord(ws.cell(row_num, cols["LAT"]).value)
            lon = clean_coord(ws.cell(row_num, cols["LONG"]).value)

            if lat is None or lon is None:
                skipped_no_coord += 1
                continue

            old_area = ws.cell(row_num, cols["AREA"]).value
            old_address = ws.cell(row_num, cols["ALAMAT"]).value

            # If both fields are already filled, don't overwrite them.
            if (
                old_area not in (None, "")
                and old_address not in (None, "")
            ):
                already_complete += 1
                continue

            key = (lat, lon)
            coordinate_set.add(key)
            jobs.append((sheet_name, row_num, key))

    total_rows = len(jobs)
    unique_coordinates = len(coordinate_set)

    if total_rows == 0:
        return (
            None,
            {
                "total_rows": 0,
                "unique_coordinates": 0,
                "processed": 0,
                "errors": 0,
                "skipped_no_coord": skipped_no_coord,
                "already_complete": already_complete,
                "message": (
                    "Tidak ada baris yang perlu diproses. "
                    "Pastikan kolom Lat dan Long terisi."
                ),
            },
        )

    # Reverse geocode unique coordinates once.
    coordinates = list(coordinate_set)
    results = {}
    errors = 0

    progress = st.progress(0, text="Menyiapkan koordinat...")
    status = st.empty()

    # Four workers keeps the process substantially faster than one request
    # at a time while avoiding an aggressive request flood.
    max_workers = 4

    def lookup(coord):
        lat, lon = coord
        if provider == "Google Maps":
            return coord, geocode_google(lat, lon, google_key)
        return coord, geocode_photon(lat, lon)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {
            executor.submit(lookup, coord): coord
            for coord in coordinates
        }

        for index, future in enumerate(as_completed(future_map), start=1):
            coord = future_map[future]

            try:
                _, result = future.result()
            except Exception as exc:
                result = {
                    "area": "",
                    "alamat": "",
                    "status": f"ERROR: {type(exc).__name__}",
                }

            results[coord] = result

            if str(result.get("status", "")).startswith("ERROR"):
                errors += 1

            pct = index / max(unique_coordinates, 1)
            progress.progress(
                pct,
                text=f"Reverse geocoding {index:,} / {unique_coordinates:,}"
            )
            status.write(
                f"Koordinat selesai: **{index:,} / {unique_coordinates:,}**"
            )

    # Second pass: write results into workbook.
    rows_written = 0

    for sheet_name, row_num, coord in jobs:
        ws = workbook[sheet_name]
        cols = get_columns(ws)

        result = results.get(coord, {})
        area = result.get("area", "")
        address = result.get("alamat", "")

        if area:
            ws.cell(row_num, cols["AREA"]).value = area

        if address:
            ws.cell(row_num, cols["ALAMAT"]).value = address

        rows_written += 1

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    progress.progress(1.0, text="Selesai.")
    status.empty()

    return output.getvalue(), {
        "total_rows": total_rows,
        "unique_coordinates": unique_coordinates,
        "processed": rows_written,
        "errors": errors,
        "skipped_no_coord": skipped_no_coord,
        "already_complete": already_complete,
        "message": "",
    }


# ------------------------------------------------------------------
# UI
# ------------------------------------------------------------------

uploaded = st.file_uploader(
    "Upload file Excel",
    type=["xlsx"],
    help="Upload file yang sudah kamu filter hanya GRSM 8.",
)

provider = st.radio(
    "Sumber reverse geocoding",
    ["OpenStreetMap / Photon", "Google Maps"],
    horizontal=True,
)

google_key = None

if provider == "Google Maps":
    google_key = st.text_input(
        "Google Maps Geocoding API Key",
        type="password",
    )

if uploaded:
    st.success(
        f"File siap diproses: **{uploaded.name}** "
        f"({uploaded.size / 1024:.1f} KB)"
    )

    if st.button(
        "🚀 Proses Koordinat",
        type="primary",
        disabled=(provider == "Google Maps" and not google_key),
    ):
        try:
            with st.spinner("Membaca Excel..."):
                output, stats = process_workbook(
                    uploaded,
                    provider,
                    google_key,
                )

            if output is None:
                st.warning(stats["message"])
            else:
                st.success("🎉 Proses selesai!")

                col1, col2, col3, col4 = st.columns(4)
                col1.metric(
                    "Baris diproses",
                    f"{stats['processed']:,}",
                )
                col2.metric(
                    "Koordinat unik",
                    f"{stats['unique_coordinates']:,}",
                )
                col3.metric(
                    "Tanpa koordinat",
                    f"{stats['skipped_no_coord']:,}",
                )
                col4.metric(
                    "Error",
                    f"{stats['errors']:,}",
                )

                st.download_button(
                    "📥 Download Excel hasil",
                    data=output,
                    file_name="EXCLUSIVE TOKO M1 M3_GRSM8_terisi.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    type="primary",
                )

        except Exception as exc:
            st.error("Aplikasi mengalami error saat memproses file.")
            st.exception(exc)

st.markdown("---")
st.caption(
    "OpenStreetMap / Photon digunakan sebagai default tanpa API key. "
    "Hasil reverse geocoding dapat berbeda dari label Google Maps."
)
