import io
import time
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(
    page_title="GRSM 8 — Koordinat → Area, Alamat & Lokasi (v7)",
    page_icon="📍",
    layout="wide",
)

st.title("📍 GRSM 8 — Koordinat → Area, Alamat & Lokasi (v7)")
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
    """Only accept explicitly named, place-like OSM features as POIs."""
    key = str(props.get("osm_key") or "").lower().strip()
    name = str(props.get("name") or "").strip()

    if not name or key in {"highway", "place"}:
        return False

    poi_keys = {
        "amenity", "shop", "tourism", "leisure", "office",
        "public_transport", "railway", "healthcare", "education",
        "craft", "historic", "sport"
    }
    return key in poi_keys


def looks_like_address_fragment(name):
    low = str(name or "").strip().lower()
    if not low:
        return True
    if re.fullmatch(r"(rt|rw)\s*[\d /-]+", low):
        return True
    prefixes = (
        "rt ", "rw ", "rt.", "rw.", "rukun tetangga", "rukun warga",
        "dusun ", "lingkungan ", "blok ", "jalan ", "jl. ", "jl ",
        "no. ", "nomor "
    )
    return low.startswith(prefixes)

def extract_poi_from_photon(props):
    """Return POI category/name only when the map has an explicit POI tag."""
    if not is_poi_feature(props):
        return "", ""

    category = normalize_location_category(
        props.get("osm_key"),
        props.get("osm_value"),
        props.get("type"),
    )
    name = str(props.get("name") or "").strip()

    if looks_like_address_fragment(name) or not category:
        return "", ""

    return category, name

@st.cache_data(show_spinner=False, ttl=60 * 60 * 24 * 30)
def geocode_photon(lat, lon):
    """Primary reverse geocoder using Photon/OpenStreetMap."""
    url = "https://photon.komoot.io/reverse"
    headers = {"User-Agent": "GRSM8-Coordinate-Geocoder/6.0 (+https://github.com/nurulirtiah/grsm8-geocoder)"}

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
                "area": clean_area_label(area),
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
        "User-Agent": "GRSM8-Coordinate-Geocoder/6.0 (+https://github.com/nurulirtiah/grsm8-geocoder)",
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
        "area": clean_area_label(area),
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


def prepare_workbook(uploaded_file):
    """Read workbook and collect rows/unique coordinates without geocoding."""
    file_bytes = uploaded_file.getvalue()
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=False)

    sheet_names = [
        name for name in ["Exclusive M1", "Exclusive M3"]
        if name in workbook.sheetnames
    ]
    if not sheet_names:
        sheet_names = workbook.sheetnames

    jobs = []
    coordinate_set = set()
    skipped_no_coord = 0
    already_complete = 0

    for sheet_name in sheet_names:
        ws = workbook[sheet_name]
        cols = get_columns(ws)
        required = ["AREA", "LAT", "LONG", "ALAMAT", "KATEGORI LOKASI", "NAMA LOKASI"]
        missing = [c for c in required if c not in cols]
        if missing:
            continue

        for row_num in range(2, ws.max_row + 1):
            lat = clean_coord(ws.cell(row_num, cols["LAT"]).value)
            lon = clean_coord(ws.cell(row_num, cols["LONG"]).value)
            if lat is None or lon is None:
                skipped_no_coord += 1
                continue

            values = [
                ws.cell(row_num, cols["AREA"]).value,
                ws.cell(row_num, cols["ALAMAT"]).value,
                ws.cell(row_num, cols["KATEGORI LOKASI"]).value,
                ws.cell(row_num, cols["NAMA LOKASI"]).value,
            ]
            if all(value not in (None, "") for value in values):
                already_complete += 1
                continue

            key = (lat, lon)
            coordinate_set.add(key)
            jobs.append((sheet_name, row_num, key))

    return {
        "file_bytes": file_bytes,
        "sheet_names": sheet_names,
        "jobs": jobs,
        "coordinates": list(coordinate_set),
        "skipped_no_coord": skipped_no_coord,
        "already_complete": already_complete,
    }


def write_results_to_workbook(prepared, results):
    """Create an Excel output using all results collected so far."""
    workbook = load_workbook(io.BytesIO(prepared["file_bytes"]), data_only=False)

    for sheet_name, row_num, coord in prepared["jobs"]:
        result = results.get(coord, {})
        if not result:
            continue

        ws = workbook[sheet_name]
        cols = get_columns(ws)

        if result.get("area"):
            ws.cell(row_num, cols["AREA"]).value = result["area"]
        if result.get("alamat"):
            ws.cell(row_num, cols["ALAMAT"]).value = result["alamat"]
        if result.get("kategori_lokasi"):
            ws.cell(row_num, cols["KATEGORI LOKASI"]).value = result["kategori_lokasi"]
        if result.get("nama_lokasi"):
            ws.cell(row_num, cols["NAMA LOKASI"]).value = result["nama_lokasi"]

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.getvalue()


def process_coordinate_batch(coordinates, provider, google_key=None):
    """Process a small batch so Streamlit does not hold thousands of requests at once."""
    results = {}
    errors = 0

    # Two workers is intentionally conservative for the public Photon service.
    max_workers = 2

    def lookup(coord):
        lat, lon = coord
        if provider == "Google Maps":
            return coord, geocode_google(lat, lon, google_key)
        return coord, geocode_photon(lat, lon)

    progress = st.progress(0, text=f"Memproses 0 / {len(coordinates):,} batch ini...")
    status = st.empty()

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(lookup, coord): coord for coord in coordinates}
        for index, future in enumerate(as_completed(future_map), start=1):
            coord = future_map[future]
            try:
                _, result = future.result()
            except Exception as exc:
                result = {
                    "area": "",
                    "alamat": "",
                    "kategori_lokasi": "",
                    "nama_lokasi": "",
                    "status": f"ERROR: {type(exc).__name__}",
                }

            results[coord] = result
            if str(result.get("status", "")).startswith("ERROR"):
                errors += 1

            progress.progress(
                index / max(len(coordinates), 1),
                text=f"Batch: {index:,} / {len(coordinates):,}",
            )
            status.write(f"Koordinat selesai: **{index:,} / {len(coordinates):,}**")

    status.empty()
    return results, errors



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
        value=False,
        help="Sengaja dimatikan dulu. Nominatim publik dibatasi 1 request/detik.",
    )

if provider == "Google Maps":
    google_key = st.text_input("Google Maps Geocoding API Key", type="password")

if uploaded:
    st.success(
        f"File siap diproses: **{uploaded.name}** ({uploaded.size / 1024:.1f} KB)"
    )

    import hashlib
    file_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()
    session_key = f"grsm8_{file_hash}_{provider}"

    if st.session_state.get("active_session") != session_key:
        prepared = prepare_workbook(uploaded)
        st.session_state.active_session = session_key
        st.session_state.prepared = prepared
        st.session_state.results = {}
        st.session_state.errors = 0

    prepared = st.session_state.prepared
    results = st.session_state.results
    all_coordinates = prepared["coordinates"]
    pending = [coord for coord in all_coordinates if coord not in results]
    total_unique = len(all_coordinates)
    completed = len(results)

    st.info(
        f"**Progres saat ini: {completed:,} / {total_unique:,} koordinat.** "
        "Aplikasi diproses bertahap supaya kalau berhenti, progres sebelumnya tidak hilang."
    )

    if total_unique:
        st.progress(
            completed / total_unique,
            text=f"Total: {completed:,} / {total_unique:,} koordinat selesai",
        )

    col1, col2, col3 = st.columns(3)
    col1.metric("Koordinat unik", f"{total_unique:,}")
    col2.metric("Sudah selesai", f"{completed:,}")
    col3.metric("Sisa", f"{len(pending):,}")

    if pending:
        batch_size = 100
        label = (
            f"🚀 Proses {min(batch_size, len(pending)):,} koordinat berikutnya"
            if completed
            else f"🚀 Mulai proses {min(batch_size, len(pending)):,} koordinat"
        )

        if st.button(
            label,
            type="primary",
            disabled=(provider == "Google Maps" and not google_key),
        ):
            batch = pending[:batch_size]
            batch_results, batch_errors = process_coordinate_batch(
                batch,
                provider,
                google_key,
            )
            st.session_state.results.update(batch_results)
            st.session_state.errors += batch_errors
            st.rerun()

    else:
        st.success("🎉 Semua koordinat unik sudah selesai diproses!")

    # Optional Nominatim fallback is intentionally offered only after Photon/Google
    # batches have completed. This avoids unexpectedly generating hundreds of
    # public Nominatim requests during the main run.
    if not pending and use_fallback and provider == "OpenStreetMap / Photon":
        incomplete = [
            coord for coord, result in results.items()
            if not result.get("area") or not result.get("alamat")
        ]
        if incomplete:
            st.warning(
                f"Masih ada {len(incomplete):,} koordinat dengan Area/Alamat belum lengkap. "
                "Fallback Nominatim belum dijalankan otomatis."
            )

    if results:
        output = write_results_to_workbook(prepared, results)
        done = len(results) >= total_unique
        st.download_button(
            "📥 Download hasil sementara" if not done else "📥 Download Excel hasil FINAL",
            data=output,
            file_name=(
                "EXCLUSIVE TOKO M1 M3_GRSM8_progress.xlsx"
                if not done
                else "EXCLUSIVE TOKO M1 M3_GRSM8_terisi.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        if not done:
            st.caption(
                "💡 Kamu boleh download hasil sementara kapan saja. "
                "Setelah itu klik tombol proses lagi untuk melanjutkan batch berikutnya."
            )

st.markdown("---")
st.caption(
    "OpenStreetMap / Photon digunakan sebagai default tanpa API key. "
    "Versi ini memproses 100 koordinat per batch dan menyimpan progres di sesi browser."
)
