# GRSM 8 — Coordinate to Area & Address

Streamlit app untuk:
- upload Excel
- memproses hanya GRSM = `MU - GRSM 8`
- sheet `Exclusive M1` dan `Exclusive M3`
- membaca Lat/Long
- mengisi `Area` dan `ALAMAT`
- download kembali Excel hasil

## Deploy ke Streamlit Community Cloud

1. Buat repository GitHub baru, misalnya `grsm8-geocoder`.
2. Upload:
   - `app.py`
   - `requirements.txt`
3. Buka Streamlit Community Cloud.
4. Pilih repository tersebut dan file `app.py`.
5. Deploy.

## Sumber geocoding

### OpenStreetMap / Photon
Tidak perlu API key. Cocok untuk mencoba alur aplikasi, tetapi hasil alamat dapat berbeda dari Google Maps dan layanan publik memiliki batas penggunaan.

### Google Maps
Pilih `Google Maps` di aplikasi dan masukkan Google Maps Geocoding API key.
Google Cloud project perlu mengaktifkan Geocoding API dan billing sesuai kebijakan Google.

## Privasi
File Excel hanya diproses oleh aplikasi saat di-upload. Jangan memasukkan API key ke source code atau commit ke GitHub.

## Penting untuk data kerja
Jangan upload file Excel berisi data outlet ke repository GitHub. Repository cukup berisi kode aplikasi.
