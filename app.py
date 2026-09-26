import io
import re
import time
import hashlib
import random
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
st.caption("Upload Excel → proses semua koordinat → isi Area, ALAMAT, Kategori Lokasi & Nama Lokasi → download Excel.")

# ================================================================
# Konfigurasi
# ================================================================
PHOTON_URL = "https://photon.komoot.io/reverse"
APP_UA = "GRSM8-Geocoder/8.0 (+https://github.com/nurulirtiah/grsm8-geocoder)"
PHOTON_WORKERS = 4
REQUEST_TIMEOUT = 12
MAX_RETRIES = 4

POI_KEYS = {
    "amenity", "shop", "tourism", "leisure", "office", "public_transport",
    "railway", "healthcare", "education", "craft", "historic", "sport"
}

CATEGORY_MAP = {
    ("amenity", "school"): "Sekolah",
    ("amenity", "college"): "Perguruan Tinggi",
    ("amenity", "university"): "Perguruan Tinggi",
    ("amenity", "kindergarten"): "Sekolah",
    ("amenity", "hospital"): "Rumah Sakit",
    ("amenity", "clinic"): "Klinik",
    ("amenity", "doctors"): "Klinik",
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
    ("amenity", "library"): "Perpustakaan",
    ("amenity", "marketplace"): "Pasar",
    ("amenity", "bus_station"): "Terminal",
    ("shop", "supermarket"): "Supermarket",
    ("shop", "convenience"): "Toko",
    ("shop", "department_store"): "Toko",
    ("shop", "mall"): "Pusat Perbelanjaan",
    ("shop", "bakery"): "Toko Roti",
    ("shop", "clothes"): "Toko Pakaian",
    ("shop", "hardware"): "Toko Bangunan",
    ("shop", "car"): "Dealer Mobil",
    ("shop", "car_repair"): "Bengkel",
    ("shop", "motorcycle"): "Dealer Motor",
    ("shop", "mobile_phone"): "Toko HP",
    ("shop", "beauty"): "Salon/Kecantikan",
    ("shop", "hairdresser"): "Salon",
    ("tourism", "hotel"): "Hotel",
    ("tourism", "guest_house"): "Penginapan",
    ("tourism", "motel"): "Penginapan",
    ("tourism", "attraction"): "Tempat Wisata",
    ("tourism", "museum"): "Museum",
    ("leisure", "sports_centre"): "Pusat Olahraga",
    ("leisure", "stadium"): "Stadion",
    ("office", "government"): "Kantor Pemerintah",
    ("office", "company"): "Kantor",
    ("office", "ngo"): "Organisasi",
    ("public_transport", "station"): "Stasiun",
    ("railway", "station"): "Stasiun",
    ("railway", "halt"): "Stasiun/Halte",
    ("healthcare", "hospital"): "Rumah Sakit",
    ("healthcare", "clinic"): "Klinik",
    ("historic", "monument"): "Monumen",
    ("sport", "pitch"): "Lapangan Olahraga",
}

RELIGIOUS_VALUES = {
    "church": "Gereja",
    "chapel": "Gereja",
    "mosque": "Masjid",
    "musalla": "Masjid",
    "temple": "Tempat Ibadah",
    "shrine": "Tempat Ibadah",
}


def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "").strip())


def clean_area_label(value):
    """Normalize a likely Kabupaten/Kota label without inventing one."""
    value = clean_text(value)
    if not value:
        return ""
    value = re.sub(r"^(kabupaten|kab\.?|kota)\s+", "", value, flags=re.I)
    return value.strip(" ,-")


def clean_coord(value):
    try:
        if value is None or str(value).strip() == "":
            return None
        number = float(value)
        if not (-90 <= number <= 90):
            return None
        return round(number, 7)
    except (ValueError, TypeError):
        return None


def looks_like_address_fragment(name):
    low = clean_text(name).lower()
    if not low:
        return True
    if re.fullmatch(r"(?:rt|rw)\s*[\d /.-]+", low):
        return True
    bad_starts = (
        "rt ", "rw ", "rt.", "rw.", "rukun tetangga", "rukun warga",
        "dusun ", "lingkungan ", "blok ", "jalan ", "jl. ", "jl ",
        "no. ", "nomor ",
    )
    return low.startswith(bad_starts)


def category_for(props):
    key = clean_text(props.get("osm_key")).lower()
    value = clean_text(props.get("osm_value")).lower()
    if value in RELIGIOUS_VALUES:
        return RELIGIOUS_VALUES[value]
    if (key, value) in CATEGORY_MAP:
        return CATEGORY_MAP[(key, value)]
    if key == "shop":
        return "Toko"
    if value in {"school", "kindergarten", "college", "university"}:
        return "Sekolah"
    return ""


def explicit_poi(props):
    """Return category/name only when OSM explicitly describes a named POI."""
    key = clean_text(props.get("osm_key")).lower()
    name = clean_text(props.get("name"))
    if not name or key not in POI_KEYS or looks_like_address_fragment(name):
        return "", ""
    category = category_for(props)
    if not category:
        return "", ""
    return category, name


def choose_area(props):
    # For Indonesia, county is preferred because it most closely maps to Kab/Kota.
    # If absent, fall back through locality levels rather than using a village/district.
    for key in ("county", "city", "town", "municipality", "state_district"):
        value = clean_text(props.get(key))
        if value:
            return clean_area_label(value)
    return ""


def build_address(props):
    road = clean_text(props.get("street") or "")
    house = clean_text(props.get("housenumber") or "")
    desa = clean_text(
        props.get("village") or props.get("suburb") or props.get("neighbourhood") or ""
    )
    kec = clean_text(props.get("district") or props.get("city_district") or props.get("locality") or "")
    area = choose_area(props)
    state = clean_text(props.get("state") or "")

    first = f"{road} No. {house}" if road and house else road
    parts = []
    for value in (first, desa, kec, area, state):
        if value and value not in parts:
            parts.append(value)
    return ", ".join(parts), kec, area


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24 * 30)
def geocode_photon_cached(lat, lon):
    """Reverse geocode one coordinate. Cached per coordinate so reruns resume."""
    headers = {"User-Agent": APP_UA, "Accept-Language": "id,en"}
    last_error = ""

    for attempt in range(MAX_RETRIES):
        try:
            response = requests.get(
                PHOTON_URL,
                params={
                    "lat": float(lat),
                    "lon": float(lon),
                    "limit": 5,
                },
                headers=headers,
                timeout=REQUEST_TIMEOUT,
            )

            if response.status_code in (429, 500, 502, 503, 504):
                last_error = f"HTTP {response.status_code}"
                time.sleep(min(6, (1.2 ** attempt) + random.uniform(0.2, 0.8)))
                continue

            response.raise_for_status()
            data = response.json()
            features = data.get("features") or []
            if not features:
                return empty_result("TIDAK DITEMUKAN")

            # First feature is used for the address. Among all returned features,
            # independently look for a named, explicitly typed POI.
            address_feature = features[0].get("properties", {}) or {}
            address, kec, area = build_address(address_feature)

            poi_category = ""
            poi_name = ""
            for feature in features:
                props = feature.get("properties", {}) or {}
                category, name = explicit_poi(props)
                if category and name:
                    poi_category, poi_name = category, name
                    break

            # If the first feature is itself a POI, it should win over a less relevant
            # later feature because reverse results are proximity/relevance ordered.
            first_cat, first_name = explicit_poi(address_feature)
            if first_cat and first_name:
                poi_category, poi_name = first_cat, first_name

            status = "OK" if area or address else "TIDAK LENGKAP"
            return {
                "area": area,
                "alamat": address,
                "kecamatan": kec,
                "kategori_lokasi": poi_category,
                "nama_lokasi": poi_name,
                "status": status,
                "error": "",
            }

        except Exception as exc:
            last_error = type(exc).__name__
            if attempt < MAX_RETRIES - 1:
                time.sleep(min(5, 0.8 * (attempt + 1) + random.uniform(0.1, 0.5)))

    raise RuntimeError(last_error or "Photon gagal")


def empty_result(status=""):
    return {
        "area": "",
        "alamat": "",
        "kecamatan": "",
        "kategori_lokasi": "",
        "nama_lokasi": "",
        "status": status,
        "error": status if str(status).startswith("ERROR") else "",
    }


def geocode_photon(lat, lon):
    """Use cached successful results; do not cache final failures so retry can recover."""
    try:
        return geocode_photon_cached(lat, lon)
    except Exception as exc:
        return empty_result(f"ERROR: {type(exc).__name__}: {exc}")


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24 * 30)
def geocode_nominatim(lat, lon):
    """Slow emergency fallback for one coordinate only."""
    url = "https://nominatim.openstreetmap.org/reverse"
    headers = {
        "User-Agent": APP_UA,
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
        timeout=20,
    )
    response.raise_for_status()
    data = response.json()
    addr = data.get("address") or {}

    road = clean_text(addr.get("road") or addr.get("pedestrian") or addr.get("footway") or "")
    house = clean_text(addr.get("house_number") or "")
    desa = clean_text(addr.get("village") or addr.get("suburb") or addr.get("neighbourhood") or addr.get("hamlet") or "")
    kec = clean_text(addr.get("district") or addr.get("city_district") or "")
    area = clean_area_label(addr.get("county") or addr.get("city") or addr.get("town") or addr.get("municipality") or "")
    state = clean_text(addr.get("state") or "")

    first = f"{road} No. {house}" if road and house else road
    parts = []
    for value in (first, desa, kec, area, state):
        if value and value not in parts:
            parts.append(value)

    poi_name = clean_text((data.get("namedetails") or {}).get("name") or data.get("name") or "")
    if poi_name and poi_name.lower() == road.lower():
        poi_name = ""

    category = category_for({
        "osm_key": data.get("category") or data.get("class"),
        "osm_value": data.get("type"),
    })
    if not category:
        category = RELIGIOUS_VALUES.get(clean_text(data.get("type")).lower(), "")

    return {
        "area": area,
        "alamat": ", ".join(parts),
        "kecamatan": kec,
        "kategori_lokasi": category if poi_name else "",
        "nama_lokasi": poi_name,
        "status": "OK_FALLBACK",
        "error": "",
    }


def get_columns(ws):
    headers = {}
    for idx, cell in enumerate(ws[1], start=1):
        header = clean_text(cell.value).upper() if cell.value is not None else ""
        if header:
            headers[header] = idx
    return headers


def ensure_output_columns(ws):
    """Ensure Area/Alamat/POI columns exist, with POI columns immediately after ALAMAT."""
    cols = get_columns(ws)

    if "AREA" not in cols:
        # Append missing Area at the end.
        ws.cell(1, ws.max_column + 1).value = "Area"
        cols = get_columns(ws)

    if "ALAMAT" not in cols:
        ws.cell(1, ws.max_column + 1).value = "ALAMAT"
        cols = get_columns(ws)

    # Insert missing POI columns right after ALAMAT, preserving requested layout.
    alamat_col = cols["ALAMAT"]
    if "KATEGORI LOKASI" not in cols:
        ws.insert_cols(alamat_col + 1, 1)
        ws.cell(1, alamat_col + 1).value = "KATEGORI LOKASI"
        cols = get_columns(ws)
    if "NAMA LOKASI" not in cols:
        alamat_col = get_columns(ws)["ALAMAT"]
        kategori_col = get_columns(ws)["KATEGORI LOKASI"]
        insert_at = max(alamat_col, kategori_col) + 1
        ws.insert_cols(insert_at, 1)
        ws.cell(1, insert_at).value = "NAMA LOKASI"

    return get_columns(ws)


def prepare_workbook(file_bytes):
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=False)
    sheet_names = [n for n in ("Exclusive M1", "Exclusive M3") if n in workbook.sheetnames]
    if not sheet_names:
        sheet_names = workbook.sheetnames

    jobs = []
    coordinate_set = set()
    skipped_no_coord = 0
    complete_rows = 0

    for sheet_name in sheet_names:
        ws = workbook[sheet_name]
        cols = ensure_output_columns(ws)
        for row_num in range(2, ws.max_row + 1):
            lat = clean_coord(ws.cell(row_num, cols["LAT"]).value) if "LAT" in cols else None
            lon = clean_coord(ws.cell(row_num, cols["LONG"]).value) if "LONG" in cols else None
            if lat is None or lon is None:
                skipped_no_coord += 1
                continue

            values = [
                ws.cell(row_num, cols["AREA"]).value,
                ws.cell(row_num, cols["ALAMAT"]).value,
                ws.cell(row_num, cols["KATEGORI LOKASI"]).value,
                ws.cell(row_num, cols["NAMA LOKASI"]).value,
            ]
            if all(clean_text(v) for v in values):
                complete_rows += 1
                continue

            coord = (lat, lon)
            coordinate_set.add(coord)
            jobs.append((sheet_name, row_num, coord))

    return {
        "file_bytes": file_bytes,
        "sheet_names": sheet_names,
        "jobs": jobs,
        "coordinates": sorted(coordinate_set),
        "skipped_no_coord": skipped_no_coord,
        "complete_rows": complete_rows,
    }


def merge_result(old, new):
    """Only replace missing fields; never erase a good value with a blank."""
    merged = dict(old or {})
    for key in ("area", "alamat", "kecamatan", "kategori_lokasi", "nama_lokasi"):
        if clean_text(new.get(key)):
            merged[key] = new[key]
        else:
            merged.setdefault(key, "")
    merged["status"] = new.get("status", merged.get("status", ""))
    merged["error"] = new.get("error", "")
    return merged


def write_results(prepared, results):
    workbook = load_workbook(io.BytesIO(prepared["file_bytes"]), data_only=False)

    for sheet_name in prepared["sheet_names"]:
        ws = workbook[sheet_name]
        cols = ensure_output_columns(ws)
        for sheet, row_num, coord in prepared["jobs"]:
            if sheet != sheet_name:
                continue
            result = results.get(coord)
            if not result:
                continue
            for field, header in (
                ("area", "AREA"),
                ("alamat", "ALAMAT"),
                ("kategori_lokasi", "KATEGORI LOKASI"),
                ("nama_lokasi", "NAMA LOKASI"),
            ):
                value = clean_text(result.get(field))
                if value:
                    ws.cell(row_num, cols[header]).value = value

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.getvalue()


def process_all_coordinates(prepared, provider, google_key, use_nominatim, results):
    """Run all pending coordinates in one user action, but internally in safe chunks.

    The chunks are implementation detail only: the user sees one continuous progress bar.
    Results are written to session_state after every completed future, so a rerun can resume.
    """
    coordinates = prepared["coordinates"]
    pending = [
        c for c in coordinates
        if c not in results or str(results.get(c, {}).get("status", "")).startswith("ERROR")
    ]
    if not pending:
        return 0

    progress = st.progress(0, text=f"Memproses 0 / {len(pending):,} koordinat tersisa...")
    status = st.empty()
    errors = 0
    done = 0

    def lookup(coord):
        lat, lon = coord
        if provider == "Google Maps":
            return coord, geocode_google(lat, lon, google_key)
        return coord, geocode_photon(lat, lon)

    # Keep the number of in-flight futures bounded. This is NOT a user-facing batch.
    chunk_size = 80
    for start in range(0, len(pending), chunk_size):
        chunk = pending[start:start + chunk_size]
        with ThreadPoolExecutor(max_workers=PHOTON_WORKERS) as executor:
            futures = {executor.submit(lookup, coord): coord for coord in chunk}
            for future in as_completed(futures):
                coord = futures[future]
                try:
                    _, result = future.result()
                except Exception as exc:
                    result = empty_result(f"ERROR: {type(exc).__name__}")

                # Persist immediately. If Streamlit reruns later, successful coordinates
                # are already in session_state and cached geocoding prevents duplicate calls.
                results[coord] = merge_result(results.get(coord), result)
                done += 1
                if str(result.get("status", "")).startswith("ERROR"):
                    errors += 1

                overall_done = len(coordinates) - len(pending) + done
                total = len(coordinates)
                progress.progress(
                    min(overall_done / max(total, 1), 1.0),
                    text=f"Reverse geocoding: {overall_done:,} / {total:,}",
                )
                if done == 1 or done % 25 == 0 or done == len(pending):
                    status.write(
                        f"Sudah selesai **{overall_done:,} / {total:,}** · "
                        f"error sementara **{errors:,}**"
                    )

        # Let Streamlit breathe between small internal chunks.
        time.sleep(0.05)

    # Optional emergency repair: only failed/incomplete coordinates, never the whole file.
    if use_nominatim and provider == "OpenStreetMap / Photon":
        repair = [
            c for c in pending
            if not results.get(c, {}).get("area") or not results.get(c, {}).get("alamat")
        ]
        if repair:
            status.write(f"Melengkapi {len(repair):,} koordinat yang masih kurang dengan Nominatim...")
            for idx, coord in enumerate(repair, start=1):
                try:
                    fallback = geocode_nominatim(*coord)
                    results[coord] = merge_result(results.get(coord), fallback)
                except Exception as exc:
                    # Keep the Photon result; do not turn a usable result into a blank.
                    results[coord]["error"] = f"Nominatim: {type(exc).__name__}"
                if idx < len(repair):
                    time.sleep(1.05)
                progress.progress(
                    min((len(coordinates) + idx) / max(len(coordinates) + len(repair), 1), 1.0),
                    text=f"Fallback Nominatim: {idx:,} / {len(repair):,}",
                )

    progress.empty()
    status.empty()
    return errors


def geocode_google(lat, lon, api_key):
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    response = requests.get(
        url,
        params={"latlng": f"{float(lat)},{float(lon)}", "key": api_key, "language": "id"},
        timeout=15,
    )
    response.raise_for_status()
    data = response.json()
    if data.get("status") != "OK" or not data.get("results"):
        return empty_result(data.get("status", "TIDAK DITEMUKAN"))

    result = data["results"][0]
    area = ""
    for component in result.get("address_components", []):
        types = component.get("types", [])
        if "administrative_area_level_2" in types:
            area = clean_area_label(component.get("long_name"))
            break

    types = set(result.get("types", []))
    cat = ""
    if "school" in types: cat = "Sekolah"
    elif "church" in types: cat = "Gereja"
    elif "mosque" in types: cat = "Masjid"
    elif "hospital" in types: cat = "Rumah Sakit"
    elif "pharmacy" in types: cat = "Apotek"
    elif "restaurant" in types: cat = "Restoran"
    elif "cafe" in types: cat = "Kafe"
    elif "gas_station" in types: cat = "SPBU"
    elif "bank" in types: cat = "Bank"
    elif "store" in types or "shopping_mall" in types: cat = "Toko/Pusat Perbelanjaan"

    return {
        "area": area,
        "alamat": clean_text(result.get("formatted_address")),
        "kecamatan": "",
        "kategori_lokasi": cat,
        "nama_lokasi": clean_text(result.get("name")) if cat else "",
        "status": "OK",
        "error": "",
    }


# ================================================================
# UI
# ================================================================

uploaded = st.file_uploader(
    "Upload file Excel",
    type=["xlsx"],
    help="Upload file yang sudah kamu filter hanya GRSM 8."
)

provider = st.radio(
    "Sumber reverse geocoding",
    ["OpenStreetMap / Photon", "Google Maps"],
    horizontal=True,
)

google_key = None
use_nominatim = False
if provider == "Google Maps":
    google_key = st.text_input("Google Maps Geocoding API Key", type="password")
else:
    use_nominatim = st.checkbox(
        "Gunakan Nominatim hanya untuk koordinat yang masih gagal/tidak lengkap (lambat)",
        value=False,
        help="Photon tetap menjadi sumber utama. Nominatim hanya dipakai untuk sisa yang benar-benar perlu diperbaiki."
    )

if uploaded:
    st.success(f"File siap diproses: **{uploaded.name}** ({uploaded.size / 1024:.1f} KB)")

    file_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()
    session_key = f"grsm8_v8_{file_hash}_{provider}"

    if st.session_state.get("active_session") != session_key:
        prepared = prepare_workbook(uploaded.getvalue())
        st.session_state.active_session = session_key
        st.session_state.prepared = prepared
        st.session_state.results = {}
        st.session_state.errors = 0

    prepared = st.session_state.prepared
    results = st.session_state.results
    total = len(prepared["coordinates"])
    completed = len(results)
    failed_coords = [c for c, r in results.items() if str(r.get("status", "")).startswith("ERROR")]
    pending = (total - completed) + len(failed_coords)

    st.info(
        f"**{completed:,} / {total:,} koordinat sudah selesai.** "
        "Sekali klik akan memproses semuanya; kalau ada request yang gagal, aplikasi mencoba ulang dan menyimpan hasil per koordinat."
    )

    st.progress(completed / max(total, 1), text=f"Total: {completed:,} / {total:,} koordinat")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Koordinat unik", f"{total:,}")
    c2.metric("Sudah selesai", f"{completed:,}")
    c3.metric("Sisa", f"{pending:,}")
    c4.metric("Error tercatat", f"{st.session_state.get('errors', 0):,}")

    if pending:
        if completed == 0:
            label = "🚀 Proses semua koordinat"
        elif failed_coords and pending == len(failed_coords):
            label = f"🔁 Coba ulang {len(failed_coords):,} koordinat yang gagal"
        else:
            label = "🚀 Lanjutkan proses koordinat yang tersisa"
        if st.button(label, type="primary", disabled=(provider == "Google Maps" and not google_key)):
            new_errors = process_all_coordinates(
                prepared,
                provider,
                google_key,
                use_nominatim,
                results,
            )
            st.session_state.errors += new_errors
            st.rerun()
    else:
        st.success("🎉 Semua koordinat unik sudah selesai diproses!")

    if results:
        output = write_results(prepared, results)
        done = completed >= total and not failed_coords
        st.download_button(
            "📥 Download Excel hasil FINAL" if done else "📥 Download hasil sementara",
            data=output,
            file_name="EXCLUSIVE TOKO M1 M3_GRSM8_terisi.xlsx" if done else "EXCLUSIVE TOKO M1 M3_GRSM8_progress.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        errors = [c for c, r in results.items() if str(r.get("status", "")).startswith("ERROR")]
        incomplete = [c for c, r in results.items() if not r.get("area") or not r.get("alamat")]
        if done:
            st.caption(
                f"Selesai. {len(errors):,} koordinat tercatat error setelah retry; "
                f"{len(incomplete):,} masih belum lengkap Area/Alamat. "
                "Koordinat yang sudah berhasil tersimpan di cache sehingga rerun tidak otomatis mengulang request yang sama."
            )

st.markdown("---")
st.caption(
    "v8: satu tombol untuk seluruh proses, retry per koordinat, cache per koordinat, "
    "POI hanya jika OSM memberi nama + kategori eksplisit, dan kolom POI otomatis dibuat bila belum ada."
)
