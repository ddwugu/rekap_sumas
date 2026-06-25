import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import xml.etree.ElementTree as ET
from xml.dom import minidom
import re
import math

# ============================================================
# KONVERSI FORMAT KOORDINAT (FIXED)
# ============================================================

def dms_to_decimal(dms_string):
    """Convert DMS string e.g. -1° 52' 5.97\" to decimal degree
    FIXED: handle simbol derajat apapun (°, ˚, etc) dan multiple spaces"""
    if pd.isna(dms_string) or dms_string == '':
        return None
    try:
        # PENTING: ganti simbol derajat/menit/detik yang aneh jadi standar dulu
        s = str(dms_string).strip()
        # Normalize simbol: ˚→°, ′→', ″→"
        s = s.replace('˚', '°').replace('′', "'").replace('″', '"')
        
        # Regex lebih robust: \s+ handle multiple spaces
        pattern = r'(-?\d+)\s*°?\s*(\d+)\s*\'?\s*([\d.]+)\s*"?'
        match = re.match(pattern, s)
        if match:
            degrees = float(match.group(1))
            minutes = float(match.group(2))
            seconds = float(match.group(3))
            decimal = abs(degrees) + minutes / 60 + seconds / 3600
            if degrees < 0:
                decimal = -decimal
            return decimal
    except:
        pass
    return None


def dm_to_decimal(dm_string):
    """Convert DM string e.g. -1° 52.0995' to decimal degree
    FIXED: handle simbol derajat apapun dan multiple spaces"""
    if pd.isna(dm_string) or dm_string == '':
        return None
    try:
        s = str(dm_string).strip()
        # Normalize simbol
        s = s.replace('˚', '°').replace('′', "'").replace('″', '"')
        
        # Format dengan simbol derajat dan menit: -1° 52.0995'
        pattern = r'(-?\d+)\s*°?\s*([\d.]+)\s*\'?'
        match = re.match(pattern, s)
        if match and ('°' in s or "'" in s):
            degrees = float(match.group(1))
            minutes = float(match.group(2))
            decimal = abs(degrees) + minutes / 60
            if degrees < 0:
                decimal = -decimal
            return decimal
        # Kalau cuma angka biasa, anggap sudah decimal degree
        return float(s)
    except:
        return None


def utm_to_latlon(x, y, zone=48, hemisphere='S'):
    """Konversi UTM ke Lat/Lon (approximate, zone 48 default - area Jambi/Sumatera)"""
    if pd.isna(x) or pd.isna(y):
        return None, None
    try:
        a = 6378137.0
        e = 0.081819191
        k0 = 0.9996

        x = float(x) - 500000.0
        y = float(y)

        if hemisphere == 'S':
            y = y - 10000000.0

        M = y / k0
        mu = M / (a * (1 - e**2/4 - 3*e**4/64 - 5*e**6/256))

        e1 = (1 - math.sqrt(1 - e**2)) / (1 + math.sqrt(1 - e**2))

        phi1 = mu + (3*e1/2 - 27*e1**3/32) * math.sin(2*mu) + \
               (21*e1**2/16 - 55*e1**4/32) * math.sin(4*mu) + \
               (151*e1**3/96) * math.sin(6*mu)

        C1 = e**2 * math.cos(phi1)**2 / (1 - e**2)
        T1 = math.tan(phi1)**2
        N1 = a / math.sqrt(1 - e**2 * math.sin(phi1)**2)
        R1 = a * (1 - e**2) / (1 - e**2 * math.sin(phi1)**2)**1.5
        D = x / (N1 * k0)

        lat = phi1 - (N1 * math.tan(phi1) / R1) * \
              (D**2/2 - (5 + 3*T1 + 10*C1 - 4*C1**2 - 9*e**2) * D**4/24 +
               (61 + 90*T1 + 298*C1 + 45*T1**2 - 252*e**2 - 3*C1**2) * D**6/720)

        lon = (D - (1 + 2*T1 + C1) * D**3/6 +
               (5 - 2*C1 + 28*T1 - 3*C1**2 + 8*e**2 + 24*T1**2) * D**5/120) / math.cos(phi1)

        lat = math.degrees(lat)
        lon = math.degrees(lon) + (zone - 1) * 6 - 180 + 3

        return lon, lat
    except:
        return None, None


def parse_coord_value(value):
    """
    Parse satu nilai koordinat (string/angka) apapun formatnya (DMS/DM/DD/UTM-raw)
    jadi angka float mentah + tipe yang terdeteksi.
    FIXED: normalize simbol dulu sebelum regex
    """
    if pd.isna(value) or value == '':
        return None, None

    s = str(value).strip()
    # Normalize simbol segera
    s = s.replace('˚', '°').replace('′', "'").replace('″', '"')

    # Ada simbol derajat → pasti DMS atau DM, hasil akhir = degree (angle)
    if '°' in s:
        val = dms_to_decimal(s)
        if val is None:
            val = dm_to_decimal(s)
        if val is not None:
            return val, 'angle'
        return None, None

    # Angka polos
    try:
        s_clean = s.rstrip(',').strip()
        if '..' in s_clean or '.-' in s_clean:
            s_clean = s_clean.replace('.-', '.').replace('..', '.')
        num = float(s_clean)
    except:
        return None, None

    # Magnitude kecil (<=180) → kemungkinan besar sudah decimal degree
    if abs(num) <= 180:
        return num, 'angle'

    # Magnitude besar → kemungkinan UTM (easting/northing dalam meter)
    return num, 'utm'


def classify_lat_lon(val_a, val_b):
    """
    Diberi 2 nilai degree (val_a, val_b) yang sudah pasti 'angle' (bukan UTM),
    tentukan mana latitude (-90..90) dan mana longitude (bisa lebih besar).
    Return (lat, lon) atau (None, None) kalau tidak bisa ditentukan.
    """
    a_is_lat = abs(val_a) <= 90
    b_is_lat = abs(val_b) <= 90

    # Kalau cuma satu yang valid sebagai lat (di luar range utk yang lain) → jelas
    if a_is_lat and not b_is_lat:
        return val_a, val_b
    if b_is_lat and not a_is_lat:
        return val_b, val_a

    # Keduanya valid sebagai lat (misal sama2 di bawah 90) → pakai heuristik magnitude
    # Konvensi umum Indonesia: longitude (95-141) > |latitude| (-11 to 6)
    if a_is_lat and b_is_lat:
        if abs(val_a) >= abs(val_b):
            return val_b, val_a   # yang magnitude lebih besar = longitude
        else:
            return val_a, val_b

    # Keduanya di luar range lat → tidak valid sama sekali
    return None, None


# ============================================================
# DETEKSI KOLOM & EKSTRAKSI KOORDINAT PER BARIS
# ============================================================

def find_column(df, candidates):
    """Cari nama kolom yang cocok (case-insensitive) dari daftar kandidat"""
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None


def detect_columns(df):
    """
    Deteksi kolom nama & koordinat.
    Prioritas: cari nama kolom eksplisit dulu (Decimal Degree, DMS, DM, UTM, Sumber Info).
    Kalau tidak ada satupun yang cocok, fallback ke POSISI kolom:
      kolom ke-1 = nama, kolom ke-2 & ke-3 = koordinat (urutan X/Y bebas, auto-detect).
    """
    name_col = find_column(df, [
        'Nama', 'Name', 'Sumber Info 1', 'Nama Titik', 'Nama Sumur', 'ID', 'Sumur'
    ])

    info2_col = find_column(df, ['Sumber Info 2'])
    info3_col = find_column(df, ['Sumber Info 3'])

    lon_dd = find_column(df, ['Longitude(Decimal Degree)', 'Longitude (Decimal Degree)', 'Lon(DD)', 'Longitude DD'])
    lat_dd = find_column(df, ['Latitude(Decimal Degree)', 'Latitude (Decimal Degree)', 'Lat(DD)', 'Latitude DD'])
    lon_dm = find_column(df, ['Longitude(DM)', 'Lon(DM)'])
    lat_dm = find_column(df, ['Latitude(DM)', 'Lat(DM)'])
    lon_dms = find_column(df, ['Longitude(DMS)', 'Lon(DMS)'])
    lat_dms = find_column(df, ['Latitude(DMS)', 'Lat(DMS)'])
    x_utm = find_column(df, ['X(UTM)', 'X (UTM)'])
    y_utm = find_column(df, ['Y(UTM)', 'Y (UTM)'])

    has_explicit_coord = any([
        lon_dd and lat_dd, lon_dm and lat_dm, lon_dms and lat_dms, x_utm and y_utm
    ])

    cols = {
        'name': name_col,
        'info2': info2_col,
        'info3': info3_col,
        'lon_dd': lon_dd, 'lat_dd': lat_dd,
        'lon_dm': lon_dm, 'lat_dm': lat_dm,
        'lon_dms': lon_dms, 'lat_dms': lat_dms,
        'x_utm': x_utm, 'y_utm': y_utm,
        'coord_col_a': None,
        'coord_col_b': None,
        'mode': 'explicit' if has_explicit_coord else 'positional',
    }

    if not has_explicit_coord:
        # Fallback POSITIONAL: kolom ke-1 = nama, kolom ke-2 & ke-3 = koordinat
        all_cols = list(df.columns)
        if len(all_cols) >= 3:
            cols['name'] = cols['name'] or all_cols[0]
            cols['coord_col_a'] = all_cols[1]
            cols['coord_col_b'] = all_cols[2]
        elif len(all_cols) == 2:
            # Tidak ada kolom nama terpisah, cuma 2 kolom koordinat
            cols['coord_col_a'] = all_cols[0]
            cols['coord_col_b'] = all_cols[1]

    return cols


def get_coordinates(row, cols, utm_zone=48, utm_hemisphere='S'):
    """
    Ambil lon, lat dari satu baris, dengan prioritas:
    1. Decimal Degree (kolom eksplisit)
    2. DM (kolom eksplisit)
    3. DMS (kolom eksplisit)
    4. UTM (kolom eksplisit X(UTM)/Y(UTM))
    5. Kolom generic X/Y → auto-detect format dari isi value
    Mengembalikan (lon, lat, format_terpakai)
    """
    # 1. Decimal Degree eksplisit
    if cols['lon_dd'] and cols['lat_dd']:
        lon_raw = row.get(cols['lon_dd'])
        lat_raw = row.get(cols['lat_dd'])
        try:
            lon = float(str(lon_raw).strip().rstrip(','))
            lat_str = str(lat_raw).strip().rstrip(',')
            if '..' in lat_str or '.-' in lat_str:
                lat_str = lat_str.replace('.-', '.').replace('..', '.')
            lat = float(lat_str)
            if pd.notna(lon) and pd.notna(lat):
                return lon, lat, 'Decimal Degree'
        except:
            pass

    # 2. DM eksplisit
    if cols['lon_dm'] and cols['lat_dm']:
        lon = dm_to_decimal(row.get(cols['lon_dm']))
        lat = dm_to_decimal(row.get(cols['lat_dm']))
        if lon is not None and lat is not None:
            return lon, lat, 'DM'

    # 3. DMS eksplisit
    if cols['lon_dms'] and cols['lat_dms']:
        lon = dms_to_decimal(row.get(cols['lon_dms']))
        lat = dms_to_decimal(row.get(cols['lat_dms']))
        if lon is not None and lat is not None:
            return lon, lat, 'DMS'

    # 4. UTM eksplisit
    if cols['x_utm'] and cols['y_utm']:
        x_val = row.get(cols['x_utm'])
        y_val = row.get(cols['y_utm'])
        if pd.notna(x_val) and pd.notna(y_val):
            lon, lat = utm_to_latlon(x_val, y_val, zone=utm_zone, hemisphere=utm_hemisphere)
            if lon is not None and lat is not None:
                return lon, lat, 'UTM'

    # 5. Mode POSITIONAL: kolom ke-2 & ke-3, urutan bebas, auto-detect format & auto-swap
    if cols.get('coord_col_a') and cols.get('coord_col_b'):
        raw_a = row.get(cols['coord_col_a'])
        raw_b = row.get(cols['coord_col_b'])

        val_a, type_a = parse_coord_value(raw_a)
        val_b, type_b = parse_coord_value(raw_b)

        if val_a is None or val_b is None:
            return None, None, None

        # Kasus 1: keduanya sudah berupa angle (degree, dari DD/DMS/DM) → classify mana lat/lon
        if type_a == 'angle' and type_b == 'angle':
            lat, lon = classify_lat_lon(val_a, val_b)
            if lat is not None and lon is not None:
                return lon, lat, 'Auto-detect (angle, auto-swap)'
            return None, None, None

        # Kasus 2: keduanya UTM raw (meter) → easting selalu lebih kecil dari northing (di Indonesia selatan)
        if type_a == 'utm' and type_b == 'utm':
            # Easting UTM Indonesia: ~100,000-900,000 | Northing (selatan, +10jt offset): ~9,000,000-10,000,000
            if val_a > val_b:
                easting, northing = val_b, val_a
            else:
                easting, northing = val_a, val_b
            lon, lat = utm_to_latlon(easting, northing, zone=utm_zone, hemisphere=utm_hemisphere)
            if lon is not None and lat is not None:
                return lon, lat, 'UTM (auto-detect, auto-swap)'
            return None, None, None

        # Kasus 3: campuran (jarang terjadi, data tidak konsisten) → gagal
        return None, None, None

    return None, None, None


def get_point_name(row, cols, idx):
    name_col = cols.get('name')
    parts = []
    if name_col:
        v = str(row.get(name_col, '')).strip()
        if v and v != 'nan':
            parts.append(v)
    if cols.get('info2'):
        v = str(row.get(cols['info2'], '')).strip()
        if v and v != 'nan':
            parts.append(v)
    if cols.get('info3'):
        v = str(row.get(cols['info3'], '')).strip()
        if v and v != 'nan':
            parts.append(v)

    if parts:
        return ' / '.join(parts)
    return f"Point_{idx+1}"


# ============================================================
# BUAT KML / KMZ
# ============================================================

def create_kml_content(df, cols, title="Excel to KMZ", utm_zone=48, utm_hemisphere='S'):
    kml = ET.Element('kml', xmlns="http://www.opengis.net/kml/2.2")
    document = ET.SubElement(kml, 'Document')

    name = ET.SubElement(document, 'name')
    name.text = title

    style_yellow = ET.SubElement(document, 'Style', id="yellowPin")
    icon_style = ET.SubElement(style_yellow, 'IconStyle')
    color = ET.SubElement(icon_style, 'color')
    color.text = "ff00ffff"
    icon = ET.SubElement(icon_style, 'Icon')
    href = ET.SubElement(icon, 'href')
    href.text = "http://maps.google.com/mapfiles/kml/paddle/ylw-circle.png"

    success_count = 0
    failed_count = 0
    failed_rows = []
    format_used_count = {}

    for idx, row in df.iterrows():
        try:
            lon, lat, fmt = get_coordinates(row, cols, utm_zone=utm_zone, utm_hemisphere=utm_hemisphere)

            if lon is None or lat is None:
                failed_count += 1
                failed_rows.append(idx + 1)
                continue

            format_used_count[fmt] = format_used_count.get(fmt, 0) + 1

            point_name = get_point_name(row, cols, idx)

            placemark = ET.SubElement(document, 'Placemark')
            pm_name = ET.SubElement(placemark, 'name')
            pm_name.text = point_name

            description = ET.SubElement(placemark, 'description')
            desc_text = f"""
            <![CDATA[
            <b>Nama:</b> {point_name}<br/>
            <b>Longitude:</b> {lon}<br/>
            <b>Latitude:</b> {lat}<br/>
            <b>Format asal:</b> {fmt}
            ]]>
            """
            description.text = desc_text

            style_url = ET.SubElement(placemark, 'styleUrl')
            style_url.text = "#yellowPin"

            point = ET.SubElement(placemark, 'Point')
            coordinates = ET.SubElement(point, 'coordinates')
            coordinates.text = f"{lon},{lat},0"

            success_count += 1

        except Exception:
            failed_count += 1
            failed_rows.append(idx + 1)
            continue

    xml_str = ET.tostring(kml, encoding='utf-8')
    dom = minidom.parseString(xml_str)
    pretty_xml = dom.toprettyxml(indent="  ", encoding='utf-8')

    return pretty_xml, success_count, failed_count, failed_rows, format_used_count


def create_kmz(kml_content):
    kmz_buffer = BytesIO()
    with zipfile.ZipFile(kmz_buffer, 'w', zipfile.ZIP_DEFLATED) as kmz:
        kmz.writestr('doc.kml', kml_content)
    kmz_buffer.seek(0)
    return kmz_buffer


# ============================================================
# STREAMLIT APP
# ============================================================

def main():
    st.set_page_config(
        page_title="Excel to KMZ - Multi Format",
        page_icon="🗺️",
        layout="wide"
    )

    st.title("🗺️ Excel to KMZ - Multi Format Auto-Detect (FIXED)")
    st.markdown("**Mendukung: Decimal Degree, DMS, DM, UTM — kolom dan format auto-detect**")
    st.info("✅ FIXED: Sekarang support simbol derajat apapun (°, ˚) dan spacing irregular")

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=['xlsx', 'xls']
    )

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)

            st.success(f"✅ File loaded. Total rows: {len(df)}")

            st.subheader("📊 Data Preview")
            st.dataframe(df.head(10), use_container_width=True)

            cols = detect_columns(df)

            st.subheader("📍 Kolom Terdeteksi")
            detected = []
            if cols['name']:
                detected.append(f"Nama → `{cols['name']}`")
            if cols['lon_dd'] and cols['lat_dd']:
                detected.append(f"Decimal Degree → `{cols['lon_dd']}` / `{cols['lat_dd']}`")
            if cols['lon_dm'] and cols['lat_dm']:
                detected.append(f"DM → `{cols['lon_dm']}` / `{cols['lat_dm']}`")
            if cols['lon_dms'] and cols['lat_dms']:
                detected.append(f"DMS → `{cols['lon_dms']}` / `{cols['lat_dms']}`")
            if cols['x_utm'] and cols['y_utm']:
                detected.append(f"UTM → `{cols['x_utm']}` / `{cols['y_utm']}`")
            if cols.get('coord_col_a') and cols.get('coord_col_b'):
                name_info = f" | Nama → `{cols['name']}`" if cols['name'] else " | (kolom nama tidak ada)"
                detected.append(
                    f"Mode posisi kolom (urutan X/Y bebas, auto-detect format) → "
                    f"`{cols['coord_col_a']}` / `{cols['coord_col_b']}`{name_info}"
                )

            if detected:
                for d in detected:
                    st.write("✅ " + d)
            else:
                st.error("❌ Tidak ada kolom koordinat yang terdeteksi. Cek nama kolom Excel.")
                st.info(f"Kolom ditemukan di file: {', '.join(df.columns.tolist())}")
                return

            # Setting UTM (kalau dipakai)
            needs_utm = bool(cols['x_utm'] and cols['y_utm']) or bool(cols.get('coord_col_a') and cols.get('coord_col_b'))
            utm_zone, utm_hemisphere = 48, 'S'
            if needs_utm:
                with st.expander("⚙️ Setting UTM (kalau data mengandung koordinat UTM)"):
                    c1, c2 = st.columns(2)
                    with c1:
                        utm_zone = st.number_input("UTM Zone", min_value=1, max_value=60, value=48)
                    with c2:
                        utm_hemisphere = st.selectbox("Hemisphere", ['S', 'N'], index=0)

            st.subheader("🔄 Convert to KMZ")
            title = st.text_input("Judul KMZ", value="Excel to KMZ")

            if st.button("Convert to KMZ", type="primary", use_container_width=True):
                with st.spinner("Converting..."):
                    kml_content, success, failed, failed_rows, fmt_count = create_kml_content(
                        df, cols, title=title, utm_zone=utm_zone, utm_hemisphere=utm_hemisphere
                    )
                    kmz_buffer = create_kmz(kml_content)

                    st.success("✅ Conversion completed!")

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Rows", len(df))
                    with col2:
                        st.metric("✅ Converted", success)
                    with col3:
                        st.metric("❌ Failed", failed)

                    if fmt_count:
                        st.write("**Format koordinat terpakai:**")
                        for fmt, count in fmt_count.items():
                            st.write(f"- {fmt}: {count} titik")

                    if failed_rows:
                        with st.expander(f"⚠️ {len(failed_rows)} baris gagal — lihat detail"):
                            st.write(f"Baris (Excel, mulai dari 1): {failed_rows}")

                    st.download_button(
                        label="📥 Download KMZ File",
                        data=kmz_buffer,
                        file_name=f"{title.replace(' ', '_')}.kmz",
                        mime="application/vnd.google-earth.kmz",
                        use_container_width=True
                    )

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.exception(e)

    else:
        st.info("👆 Upload Excel file (format apa saja, kolom akan dideteksi otomatis)")
        st.markdown("""
        ### 📋 Format kolom yang didukung (salah satu cukup):

        | Format | Nama Kolom |
        |---|---|
        | Decimal Degree | `Longitude(Decimal Degree)`, `Latitude(Decimal Degree)` |
        | DM | `Longitude(DM)`, `Latitude(DM)` |
        | DMS | `Longitude(DMS)`, `Latitude(DMS)` |
        | UTM | `X(UTM)`, `Y(UTM)` |
        | Generic | `X`, `Y` (format isi dideteksi otomatis per baris: DMS/DD/UTM) |

        Kolom nama (opsional): `Nama`, `Name`, `Sumber Info 1`, `Sumber Info 2`, `Sumber Info 3`
        
        ### ✅ Improvements:
        - Handle simbol derajat aneh (°, ˚) otomatis
        - Handle spacing irregular di DMS format
        - Detect kolom "Sumur" sebagai nama
        """)

if __name__ == "__main__":
    main()
