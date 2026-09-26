import io
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(
    page_title="GRSM 8 — Koordinat → Area, Alamat & Lokasi",
    page_icon="📍",
    layout="wide",
)

st.title("📍 GRSM 8 — Koordinat → Area, Alamat & Lokasi")
st.caption("Upload Excel → proses Lat/Long → isi Area, ALAMAT, Kategori Lokasi & Nama Lokasi → download Excel.")


# ------------------------------------------------------------------
# POI / lokasi identification
# ------------------------------------------------------------------

def normalize_location_category(osm_key, osm_value, feature_type=""):
    """Convert OSM/Photon tags into a simple Indonesian category."""
    key = str(osm_key or "").lower().strip()
    value = str(osm_value or "").lower().strip()
    ftype = str(feature_type or "").lower().strip()

    mapping = {
        ("amenity", "school"): "Sekolah",
        ("amenity", "college"): "Perguruan Tinggi",
        ("amenity", "university"): "Perguruan Tinggi",
        ("amenity", "kindergarten"): "Sekolah",
        ("amenity", "hospital"): "Rumah Sakit",
        ("amenity", "clinic"): "Klinik",
        ("amenity", "pharmacy"): "Apotek",
        ("amenity", "place_of_worship"): "Tempat Ibadah",
        ("amenity", "restaurant"): "Restoran",
        ("amenity", "cafe"): "Kafe",
        ("amenity", "fast_food"): "Restoran",
        ("amenity", "fuel"): "SPBU",
        ("amenity", "bank"): "Bank",
        ("amenity", "post_office"): "Kantor Pos",
        ("amenity", "police"): "Kepolisian",
        ("amenity", "fire_station"): "Pemadam Kebakaran",
        ("amenity", "townhall"): "Kantor Pemerintah",
        ("amenity", "community_centre"): "Fasilitas Komunitas",
        ("shop", "supermarket"): "Supermarket",
        ("shop", "convenience"): "Toko",
        ("shop", "department_store"): "Toko",
        ("shop", "mall"): "Pusat Perbelanjaan",
        ("shop", "bakery"): "Toko Roti",
        ("shop", "clothes"): "Toko Pakaian",
        ("shop", "hardware"): "Toko Bangunan",
        ("shop", "car"): "Dealer Mobil",
        ("shop", "motorcycle"): "Dealer Motor",
        ("shop", "mobile_phone"): "Toko HP",
        ("tourism", "hotel"): "Hotel",
        ("tourism", "guest_house"): "Penginapan",
        ("tourism", "motel"): "Penginapan",
        ("tourism", "attraction"): "Tempat Wisata",
        ("leisure", "sports_centre"): "Pusat Olahraga",
        ("leisure", "stadium"): "Stadion",
        ("office", "government"): "Kantor Pemerintah",
        ("office", "company"): "Kantor",
        ("office", "ngo"): "Organisasi",
        ("public_transport", "station"): "Stasiun",
        ("railway", "station"): "Stasiun",
        ("highway", "bus_stop"): "Halte",
    }

    if (key, value) in mapping:
        return mapping[(key, value)]

    # Common OSM values that are useful even when the exact combination
    # is not in the mapping above.
    if key == "shop":
        return "Toko"
    if key == "school" or value in {"school", "kindergarten", "college", "university"}:
        return "Sekolah"
    if value in {"church", "chapel"}:
        return "Gereja"
    if value in {"mosque", "musalla"}:
        return "Masjid"
    if value in {"temple", "shrine"}:
        return "Tempat Ibadah"
    if key == "place" and value in {"city", "town", "village", "hamlet"}:
        return "Permukiman"
    if ftype in {"house", "residential"}:
        return "Permukiman"

    return ""


def is_poi_feature(props):
    """True when Photon feature looks like a named POI rather than a road."""
    key = str(props.get("osm_key") or "").lower().strip()
    value = str(props.get("osm_value") or "").lower().strip()
    name = str(props.get("name") or "").strip()

    if not name:
        return False

    # A highway/street name by itself is not a location/POI.
    if key == "highway" and value not in {"bus_stop", "services", "rest_area"}:
        return False

    poi_keys = {
        "amenity", "shop", "tourism", "leisure", "office",
        "public_transport", "railway", "healthcare", "education",
        "craft", "building", "historic", "sport"
    }
    return key in poi_keys


def extract_poi_from_photon(props):
    """Return POI category + name from a Photon feature when available."""
    if not is_poi_feature(props):
        return "", ""

    category = normalize_location_category(
        props.get("osm_key"),
        props.get("osm_value"),
        props.get("type"),
    )
    name = str(props.get("name") or "").strip()

    # Don't call a generic OSM object name a POI if it has no usable category.
    if not category and name:
        key = str(props.get("osm_key") or "").lower().strip()
        if key in {"building", "historic", "sport", "craft", "healthcare", "education"}:
            category = "Fasilitas/Lokasi"

    return category, name

st.info(
    "File yang kamu upload sudah kamu filter hanya untuk GRSM 8, "
    "jadi aplikasi ini akan memproses SEMUA baris yang memiliki koordinat "
    "di sheet Exclusive M1 dan Exclusive M3. Baris tanpa Lat/Long akan dilewati."
)

st.info(
    "Kolom KATEGORI LOKASI dan NAMA LOKASI akan diisi jika koordinat "
    "teridentifikasi sebagai POI/tempat bernama di data OpenStreetMap. "
    "Jika tidak ada POI yang terdaftar, kolom tersebut dibiarkan kosong "
    "dan ALAMAT tetap diisi semaksimal mungkin."
)

# ------------------------------------------------------------------
# Reverse geocoding
# ------------------------------------------------------------------

@st.cache_data(show_spinner=False, ttl=60 * 60 * 24)
def geocode_photon(lat, lon):
    """Primary reverse geocoder using Photon/OpenStreetMap."""
    url = "https://photon.komoot.io/reverse"
    headers = {"User-Agent": "GRSM8-Coordinate-Geocoder/4.0 (+https://github.com/nurulirtiah/grsm8-geocoder)"}

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
                    "kecamatan": "",
                    "kategori_lokasi": "",
                    "nama_lokasi": "",
                    "status": "TIDAK DITEMUKAN",
                }

            props = features[0].get("properties", {})
            kategori_lokasi, nama_lokasi = extract_poi_from_photon(props)
            road = props.get("street") or props.get("name") or ""
            house = props.get("housenumber") or ""
            kec = props.get("district") or props.get("city_district") or props.get("locality") or ""
            desa = props.get("village") or props.get("suburb") or props.get("neighbourhood") or ""
            area = props.get("city") or props.get("county") or props.get("municipality") or props.get("town") or props.get("state_district") or ""

            first = f"{road} No. {house}" if road and house else road
            parts = []
            for value in [first, desa, kec, area]:
                value = str(value).strip() if value else ""
                if value and value not in parts:
                    parts.append(value)
            alamat = ", ".join(parts)

            return {
                "area": str(area),
                "alamat": alamat,
                "kecamatan": str(kec),
                "kategori_lokasi": kategori_lokasi,
                "nama_lokasi": nama_lokasi,
                "status": "OK" if (area or alamat) else "TIDAK LENGKAP",
            }
        except Exception as exc:
            if attempt < 3:
                time.sleep(1.5 * (attempt + 1))
            else:
                return {
                    "area": "",
                    "alamat": "",
                    "kecamatan": "",
                    "kategori_lokasi": "",
                    "nama_lokasi": "",
                    "status": f"ERROR: {type(exc).__name__}",
                }

    return {"area": "", "alamat": "", "kecamatan": "", "status": "ERROR: rate limit"}


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24 * 30)
def geocode_nominatim_fallback(lat, lon):
    """Fallback only for incomplete Photon results. One request at a time."""
    url = "https://nominatim.openstreetmap.org/reverse"
    headers = {
        "User-Agent": "GRSM8-Coordinate-Geocoder/4.0 (+https://github.com/nurulirtiah/grsm8-geocoder)",
        "Accept-Language": "id,en",
    }
    response = requests.get(
        url,
        params={
            "lat": float(lat),
            "lon": float(lon),
            "format": "jsonv2",
            "addressdetails": 1,
            "namedetails": 1,
            "zoom": 18,
        },
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    addr = data.get("address", {})

    road = addr.get("road") or addr.get("pedestrian") or addr.get("footway") or addr.get("path") or ""
    house = addr.get("house_number") or ""
    desa = addr.get("village") or addr.get("suburb") or addr.get("neighbourhood") or addr.get("hamlet") or ""
    kec = addr.get("district") or addr.get("city_district") or addr.get("municipality") or ""
    area = addr.get("county") or addr.get("city") or addr.get("town") or addr.get("municipality") or ""
    state = addr.get("state") or ""

    category = normalize_location_category(
        data.get("category") or data.get("class"),
        data.get("type"),
        data.get("type"),
    )
    name_details = data.get("namedetails") or {}
    poi_name = (
        name_details.get("name")
        or data.get("name")
        or ""
    )
    # Do not use a road name as the POI name.
    if poi_name and str(poi_name).strip() == str(road).strip():
        poi_name = ""

    first = f"{road} No. {house}" if road and house else road
    parts = []
    for value in [first, desa, kec, area, state]:
        value = str(value).strip() if value else ""
        if value and value not in parts:
            parts.append(value)

    alamat = ", ".join(parts)
    return {
        "area": str(area),
        "alamat": alamat,
        "kecamatan": str(kec),
        "kategori_lokasi": category,
        "nama_lokasi": str(poi_name).strip(),
        "status": "OK_FALLBACK" if (area or alamat) else "TIDAK DITEMUKAN",
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
            "kategori_lokasi": "",
            "nama_lokasi": "",
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

    result_types = set(result.get("types", []))
    google_category = ""
    google_name = result.get("name", "") or ""

    if "school" in result_types:
        google_category = "Sekolah"
    elif "church" in result_types:
        google_category = "Gereja"
    elif "mosque" in result_types:
        google_category = "Masjid"
    elif "hospital" in result_types:
        google_category = "Rumah Sakit"
    elif "pharmacy" in result_types:
        google_category = "Apotek"
    elif "restaurant" in result_types:
        google_category = "Restoran"
    elif "cafe" in result_types:
        google_category = "Kafe"
    elif "gas_station" in result_types:
        google_category = "SPBU"
    elif "bank" in result_types:
        google_category = "Bank"
    elif "store" in result_types or "shopping_mall" in result_types:
        google_category = "Toko/Pusat Perbelanjaan"

    return {
        "area": area,
        "alamat": alamat,
        "kategori_lokasi": google_category,
        "nama_lokasi": google_name,
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


def process_workbook(uploaded_file, provider, google_key=None, use_fallback=False):
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
            c for c in ["AREA", "LAT", "LONG", "ALAMAT", "KATEGORI LOKASI", "NAMA LOKASI"]
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
            old_category = ws.cell(row_num, cols["KATEGORI LOKASI"]).value
            old_name = ws.cell(row_num, cols["NAMA LOKASI"]).value

            # Only skip a row when all four output fields already have data.
            if all(
                value not in (None, "")
                for value in [old_area, old_address, old_category, old_name]
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
                "fallback_used": 0,
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

    # Fallback only for incomplete Photon results. Nominatim's public service
    # requires single-threaded use and at most 1 request/second.
    fallback_count = 0
    if provider == "OpenStreetMap / Photon" and use_fallback:
        # Nominatim fallback is used only when the address/area itself is
        # incomplete. We do NOT trigger one Nominatim request merely because
        # a POI name is absent, otherwise a large workbook could cause
        # hundreds/thousands of slow public-service requests.
        incomplete = [
            coord for coord, result in results.items()
            if not result.get("area") or not result.get("alamat")
        ]
        if incomplete:
            fallback_progress = st.progress(0, text="Melengkapi hasil yang belum lengkap...")
            fallback_status = st.empty()
            for idx, coord in enumerate(incomplete, start=1):
                try:
                    time.sleep(1.05)
                    fallback = geocode_nominatim_fallback(*coord)
                    current = results.get(coord, {})
                    if fallback.get("area"):
                        current["area"] = fallback["area"]
                    if fallback.get("alamat"):
                        current["alamat"] = fallback["alamat"]
                    if fallback.get("kategori_lokasi"):
                        current["kategori_lokasi"] = fallback["kategori_lokasi"]
                    if fallback.get("nama_lokasi"):
                        current["nama_lokasi"] = fallback["nama_lokasi"]
                    current["status"] = fallback.get("status", current.get("status", ""))
                    results[coord] = current
                    fallback_count += 1
                except Exception:
                    pass
                fallback_progress.progress(idx / len(incomplete), text=f"Melengkapi {idx:,} / {len(incomplete):,}")
                fallback_status.write(f"Fallback selesai: **{idx:,} / {len(incomplete):,}**")
            fallback_status.empty()

    # Second pass: write results into workbook.
    rows_written = 0

    for sheet_name, row_num, coord in jobs:
        ws = workbook[sheet_name]
        cols = get_columns(ws)

        result = results.get(coord, {})
        area = result.get("area", "")
        address = result.get("alamat", "")
        category = result.get("kategori_lokasi", "")
        location_name = result.get("nama_lokasi", "")

        if area:
            ws.cell(row_num, cols["AREA"]).value = area

        if address:
            ws.cell(row_num, cols["ALAMAT"]).value = address

        if category:
            ws.cell(row_num, cols["KATEGORI LOKASI"]).value = category

        if location_name:
            ws.cell(row_num, cols["NAMA LOKASI"]).value = location_name

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
        "fallback_used": fallback_count,
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

use_fallback = False
if provider == "OpenStreetMap / Photon":
    use_fallback = st.checkbox(
        "Lengkapi hasil yang masih kosong dengan Nominatim (lebih lambat, 1 koordinat/detik)",
        value=True,
        help="Fallback hanya dipakai untuk koordinat yang hasil Photon-nya belum lengkap. Hasil dicache agar koordinat yang sama tidak diminta ulang."
    )

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
                    use_fallback,
                )

            if output is None:
                st.warning(stats["message"])
            else:
                st.success("🎉 Proses selesai!")

                col1, col2, col3, col4, col5 = st.columns(5)
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
                col5.metric(
                    "Fallback",
                    f"{stats.get('fallback_used', 0):,}",
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
