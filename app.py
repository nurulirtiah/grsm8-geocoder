import io
import json
import time
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(
    page_title="GRSM 8 — Koordinat → Alamat",
    page_icon="📍",
    layout="wide",
)

st.title("📍 GRSM 8 — Koordinat → Area & Alamat")
st.caption("Upload Excel → proses koordinat GRSM 8 → download Excel hasil.")

TARGET_GRSM = "MU - GRSM 8"
SHEETS_DEFAULT = ["Exclusive M3", "Exclusive M1"]

@st.cache_data(show_spinner=False)
def geocode_photon(lat, lon):
    url = "https://photon.komoot.io/reverse"
    params = {"lat": float(lat), "lon": float(lon)}
    headers = {"User-Agent": "GRSM8-Coordinate-Geocoder/1.0"}
    r = requests.get(url, params=params, headers=headers, timeout=20)
    r.raise_for_status()
    data = r.json()

    features = data.get("features", [])
    if not features:
        return {"area": "", "alamat": "", "status": "Tidak ditemukan"}

    p = features[0].get("properties", {})
    street = p.get("street") or p.get("name") or ""
    housenumber = p.get("housenumber") or ""

    if street and housenumber:
        alamat = f"{street} No. {housenumber}"
    else:
        alamat = street

    # Prefer city/regency-level field, then district/state district.
    area = (
        p.get("city")
        or p.get("county")
        or p.get("district")
        or p.get("state_district")
        or ""
    )

    return {
        "area": area,
        "alamat": alamat,
        "status": "OK" if (area or alamat) else "Tidak lengkap",
    }


def geocode_google(lat, lon, api_key):
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {
        "latlng": f"{float(lat)},{float(lon)}",
        "key": api_key,
        "language": "id",
    }
    r = requests.get(url, params=params, timeout=20)
    r.raise_for_status()
    data = r.json()

    if data.get("status") != "OK" or not data.get("results"):
        return {
            "area": "",
            "alamat": "",
            "status": data.get("status", "Tidak ditemukan"),
        }

    result = data["results"][0]
    alamat = result.get("formatted_address", "")
    area = ""

    for comp in result.get("address_components", []):
        types = comp.get("types", [])
        if "administrative_area_level_2" in types:
            area = comp.get("long_name", "")
            break

    if not area:
        for comp in result.get("address_components", []):
            types = comp.get("types", [])
            if "locality" in types:
                area = comp.get("long_name", "")
                break

    return {
        "area": area,
        "alamat": alamat,
        "status": "OK",
    }


def normalize_coord(v):
    try:
        return round(float(v), 6)
    except Exception:
        return None


def process_workbook(uploaded_file, provider, google_key=None):
    raw = uploaded_file.getvalue()
    wb = load_workbook(io.BytesIO(raw))
    cache = {}
    results_log = []

    # Cache by rounded coordinate so identical/near-identical points
    # do not generate duplicate API requests.
    def lookup(lat, lon):
        key = f"{lat:.6f},{lon:.6f}"
        if key in cache:
            return cache[key]

        if provider == "Google Maps":
            result = geocode_google(lat, lon, google_key)
        else:
            result = geocode_photon(lat, lon)

        cache[key] = result
        time.sleep(0.35 if provider == "Google Maps" else 0.7)
        return result

    sheets = [s for s in wb.sheetnames if s in SHEETS_DEFAULT]
    if not sheets:
        sheets = wb.sheetnames

    total_candidates = 0
    for sheet_name in sheets:
        ws = wb[sheet_name]
        headers = [str(c.value).strip().upper() if c.value is not None else "" for c in ws[1]]
        col = {h: i + 1 for i, h in enumerate(headers)}

        required = ["GRSM", "AREA", "LAT", "LONG", "ALAMAT"]
        missing = [x for x in required if x not in col]
        if missing:
            results_log.append(f"{sheet_name}: dilewati, kolom hilang {missing}")
            continue

        for row in range(2, ws.max_row + 1):
            grsm = ws.cell(row, col["GRSM"]).value
            if str(grsm).strip().upper() != TARGET_GRSM.upper():
                continue

            lat = normalize_coord(ws.cell(row, col["LAT"]).value)
            lon = normalize_coord(ws.cell(row, col["LONG"]).value)

            if lat is not None and lon is not None:
                # We process rows where either target field is empty.
                area_old = ws.cell(row, col["AREA"]).value
                addr_old = ws.cell(row, col["ALAMAT"]).value
                if area_old in (None, "") or addr_old in (None, ""):
                    total_candidates += 1

    progress = st.progress(0, text="Menyiapkan...")
    done = 0
    errors = 0
    skipped = 0

    for sheet_name in sheets:
        ws = wb[sheet_name]
        headers = [str(c.value).strip().upper() if c.value is not None else "" for c in ws[1]]
        col = {h: i + 1 for i, h in enumerate(headers)}

        required = ["GRSM", "AREA", "LAT", "LONG", "ALAMAT"]
        if any(x not in col for x in required):
            continue

        for row in range(2, ws.max_row + 1):
            grsm = ws.cell(row, col["GRSM"]).value
            if str(grsm).strip().upper() != TARGET_GRSM.upper():
                continue

            lat = normalize_coord(ws.cell(row, col["LAT"]).value)
            lon = normalize_coord(ws.cell(row, col["LONG"]).value)

            if lat is None or lon is None:
                skipped += 1
                continue

            area_old = ws.cell(row, col["AREA"]).value
            addr_old = ws.cell(row, col["ALAMAT"]).value

            if area_old not in (None, "") and addr_old not in (None, ""):
                skipped += 1
                continue

            try:
                result = lookup(lat, lon)

                if result["area"]:
                    ws.cell(row, col["AREA"]).value = result["area"]

                if result["alamat"]:
                    ws.cell(row, col["ALAMAT"]).value = result["alamat"]

                if result["status"] != "OK":
                    errors += 1

            except Exception as e:
                errors += 1
                results_log.append(
                    f"{sheet_name} row {row}: {type(e).__name__}: {e}"
                )

            done += 1
            pct = min(done / max(total_candidates, 1), 1)
            progress.progress(
                pct,
                text=f"Memproses {done:,} / {total_candidates:,} — {sheet_name}",
            )

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    progress.progress(1.0, text="Selesai.")
    return out.getvalue(), len(cache), done, skipped, errors, results_log


st.info(
    "Aplikasi ini hanya memproses baris dengan GRSM = "
    f"`{TARGET_GRSM}` pada sheet Exclusive M1 dan Exclusive M3."
)

uploaded = st.file_uploader(
    "Upload file Excel",
    type=["xlsx"],
    help="Gunakan file EXCLUSIVE TOKO M1 M3.xlsx",
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
        help="API key digunakan hanya selama proses. Untuk Google Maps Geocoding API, billing/API access perlu aktif pada project Google Cloud.",
    )
    if not google_key:
        st.warning("Masukkan API key Google Maps sebelum menjalankan proses.")

if uploaded:
    if st.button("🚀 Proses GRSM 8", type="primary", disabled=(provider == "Google Maps" and not google_key)):
        with st.spinner("Membaca workbook dan memproses koordinat..."):
            output, unique_coords, done, skipped, errors, logs = process_workbook(
                uploaded, provider, google_key
            )

        st.success("Selesai diproses.")
        c1, c2, c3 = st.columns(3)
        c1.metric("Koordinat unik diproses", f"{unique_coords:,}")
        c2.metric("Baris diproses", f"{done:,}")
        c3.metric("Baris error/tidak lengkap", f"{errors:,}")

        st.download_button(
            "📥 Download Excel hasil",
            data=output,
            file_name="EXCLUSIVE TOKO M1 M3_GRSM8_terisi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if logs:
            with st.expander("Log / catatan"):
                for item in logs[:100]:
                    st.write(item)

st.markdown("---")
st.caption(
    "Catatan: hasil OpenStreetMap/Photon dapat berbeda dari label Google Maps. "
    "Untuk hasil yang mengikuti Google Maps, pilih Google Maps dan gunakan Geocoding API."
)
