import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import xml.etree.ElementTree as ET
from xml.dom import minidom
import re
import math

# ============================================================
# KONVERSI FORMAT KOORDINAT
# ============================================================

def dms_to_decimal(dms_string):
    """Convert DMS string e.g. -1° 52' 5.97\" to decimal degree"""
    if pd.isna(dms_string) or dms_string == '':
        return None
    try:
        s = str(dms_string).strip()
        s = s.replace('˚', '°').replace('′', "'").replace('″', '"')
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
    """Convert DM string e.g. -1° 52.0995' to decimal degree"""
    if pd.isna(dm_string) or dm_string == '':
        return None
    try:
        s = str(dm_string).strip()
        s = s.replace('˚', '°').replace('′', "'").replace('″', '"')
        pattern = r'(-?\d+)\s*°?\s*([\d.]+)\s*\'?'
        match = re.match(pattern, s)
        if match and ('°' in s or "'" in s):
            degrees = float(match.group(1))
            minutes = float(match.group(2))
            decimal = abs(degrees) + minutes / 60
            if degrees < 0:
                decimal = -decimal
            return decimal
        return float(s)
    except:
        return None


def utm_to_latlon(x, y, zone=48, hemisphere='S'):
    """Konversi UTM ke Lat/Lon (zone 48 default - area Jambi/Sumatera)"""
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


# ============================================================
# PRESISI PENUH: helper raw string
# ============================================================

def clean_decimal_str(value):
    """
    Kalau value adalah angka decimal polos (dari excel), kembalikan STRING aslinya
    (dibersihkan trailing koma/spasi) tanpa konversi ulang → zero rounding.
    Return None kalau bukan angka polos.
    """
    if pd.isna(value) or value == '':
        return None
    s = str(value).strip().rstrip(',').strip()
    if '..' in s or '.-' in s:
        s = s.replace('.-', '.').replace('..', '.')
    try:
        float(s)
        return s
    except:
        return None


def fmt_coord(val, raw_str=None):
    """
    Format koordinat untuk output KML.
    - Kalau raw_str tersedia & value-nya sama → pakai raw_str VERBATIM (presisi = excel).
    - Kalau hasil konversi (DMS/DM/UTM) → tulis 10 desimal (~0.01 mm, lebih dari cukup).
    """
    if raw_str is not None:
        try:
            if float(raw_str) == val:
                return raw_str
        except:
            pass
    return f"{val:.10f}".rstrip('0').rstrip('.')


def parse_coord_value(value):
    """Parse satu nilai koordinat → (float, tipe, raw_str_kalau_decimal_polos)"""
    if pd.isna(value) or value == '':
        return None, None, None

    s = str(value).strip()
    s = s.replace('˚', '°').replace('′', "'").replace('″', '"')

    if '°' in s:
        val = dms_to_decimal(s)
        if val is None:
            val = dm_to_decimal(s)
        if val is not None:
            return val, 'angle', None  # hasil konversi, tidak ada raw pass-through
        return None, None, None

    raw = clean_decimal_str(s)
    if raw is None:
        return None, None, None
    num = float(raw)

    if abs(num) <= 180:
        return num, 'angle', raw

    return num, 'utm', raw


def classify_lat_lon(val_a, val_b):
    """Tentukan mana lat (-90..90) dan mana lon. Return (lat, lon) atau (None, None)."""
    a_is_lat = abs(val_a) <= 90
    b_is_lat = abs(val_b) <= 90

    if a_is_lat and not b_is_lat:
        return val_a, val_b
    if b_is_lat and not a_is_lat:
        return val_b, val_a

    if a_is_lat and b_is_lat:
        # Konvensi Indonesia: longitude (95-141) > |latitude| (-11 to 6)
        if abs(val_a) >= abs(val_b):
            return val_b, val_a
        else:
            return val_a, val_b

    return None, None


# ============================================================
# DETEKSI KOLOM
# ============================================================

def find_column(df, candidates):
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None


def detect_columns(df):
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
        all_cols = list(df.columns)
        if len(all_cols) >= 3:
            cols['name'] = cols['name'] or all_cols[0]
            cols['coord_col_a'] = all_cols[1]
            cols['coord_col_b'] = all_cols[2]
        elif len(all_cols) == 2:
            cols['coord_col_a'] = all_cols[0]
            cols['coord_col_b'] = all_cols[1]

    return cols


def get_coordinates(row, cols, utm_zone=48, utm_hemisphere='S'):
    """
    Return (lon, lat, fmt, lon_raw, lat_raw).
    lon_raw/lat_raw = string ASLI dari excel kalau formatnya decimal polos → dipakai
    verbatim di KML (presisi 100% sama excel). None kalau hasil konversi.
    """
    # 1. Decimal Degree eksplisit
    if cols['lon_dd'] and cols['lat_dd']:
        lon_raw = clean_decimal_str(row.get(cols['lon_dd']))
        lat_raw = clean_decimal_str(row.get(cols['lat_dd']))
        if lon_raw is not None and lat_raw is not None:
            return float(lon_raw), float(lat_raw), 'Decimal Degree', lon_raw, lat_raw

    # 2. DM eksplisit
    if cols['lon_dm'] and cols['lat_dm']:
        lon = dm_to_decimal(row.get(cols['lon_dm']))
        lat = dm_to_decimal(row.get(cols['lat_dm']))
        if lon is not None and lat is not None:
            return lon, lat, 'DM', None, None

    # 3. DMS eksplisit
    if cols['lon_dms'] and cols['lat_dms']:
        lon = dms_to_decimal(row.get(cols['lon_dms']))
        lat = dms_to_decimal(row.get(cols['lat_dms']))
        if lon is not None and lat is not None:
            return lon, lat, 'DMS', None, None

    # 4. UTM eksplisit
    if cols['x_utm'] and cols['y_utm']:
        x_val = row.get(cols['x_utm'])
        y_val = row.get(cols['y_utm'])
        if pd.notna(x_val) and pd.notna(y_val):
            lon, lat = utm_to_latlon(x_val, y_val, zone=utm_zone, hemisphere=utm_hemisphere)
            if lon is not None and lat is not None:
                return lon, lat, 'UTM', None, None

    # 5. Mode POSITIONAL: auto-detect + auto-swap, raw string ikut di-swap
    if cols.get('coord_col_a') and cols.get('coord_col_b'):
        raw_val_a = row.get(cols['coord_col_a'])
        raw_val_b = row.get(cols['coord_col_b'])

        val_a, type_a, str_a = parse_coord_value(raw_val_a)
        val_b, type_b, str_b = parse_coord_value(raw_val_b)

        if val_a is None or val_b is None:
            return None, None, None, None, None

        if type_a == 'angle' and type_b == 'angle':
            lat, lon = classify_lat_lon(val_a, val_b)
            if lat is not None and lon is not None:
                # Map raw string mengikuti hasil swap
                if lat == val_a:
                    lat_raw, lon_raw = str_a, str_b
                else:
                    lat_raw, lon_raw = str_b, str_a
                return lon, lat, 'Auto-detect (angle, auto-swap)', lon_raw, lat_raw
            return None, None, None, None, None

        if type_a == 'utm' and type_b == 'utm':
            if val_a > val_b:
                easting, northing = val_b, val_a
            else:
                easting, northing = val_a, val_b
            lon, lat = utm_to_latlon(easting, northing, zone=utm_zone, hemisphere=utm_hemisphere)
            if lon is not None and lat is not None:
                return lon, lat, 'UTM (auto-detect, auto-swap)', None, None
            return None, None, None, None, None

        return None, None, None, None, None

    return None, None, None, None, None


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
    precision_samples = []  # untuk panel verifikasi

    for idx, row in df.iterrows():
        try:
            lon, lat, fmt, lon_raw, lat_raw = get_coordinates(
                row, cols, utm_zone=utm_zone, utm_hemisphere=utm_hemisphere
            )

            if lon is None or lat is None:
                failed_count += 1
                failed_rows.append(idx + 1)
                continue

            format_used_count[fmt] = format_used_count.get(fmt, 0) + 1

            # PRESISI PENUH: pakai string asli excel kalau ada
            lon_out = fmt_coord(lon, lon_raw)
            lat_out = fmt_coord(lat, lat_raw)

            if len(precision_samples) < 5:
                precision_samples.append({
                    'Baris': idx + 1,
                    'Lon (excel/asli)': lon_raw if lon_raw else '(hasil konversi)',
                    'Lat (excel/asli)': lat_raw if lat_raw else '(hasil konversi)',
                    'Lon → KML': lon_out,
                    'Lat → KML': lat_out,
                })

            point_name = get_point_name(row, cols, idx)

            placemark = ET.SubElement(document, 'Placemark')
            pm_name = ET.SubElement(placemark, 'name')
            pm_name.text = point_name

            description = ET.SubElement(placemark, 'description')
            desc_text = f"""
            <![CDATA[
            <b>Nama:</b> {point_name}<br/>
            <b>Longitude:</b> {lon_out}<br/>
            <b>Latitude:</b> {lat_out}<br/>
            <b>Format asal:</b> {fmt}
            ]]>
            """
            description.text = desc_text

            style_url = ET.SubElement(placemark, 'styleUrl')
            style_url.text = "#yellowPin"

            point = ET.SubElement(placemark, 'Point')
            coordinates = ET.SubElement(point, 'coordinates')
            coordinates.text = f"{lon_out},{lat_out},0"

            success_count += 1

        except Exception:
            failed_count += 1
            failed_rows.append(idx + 1)
            continue

    xml_str = ET.tostring(kml, encoding='utf-8')
    dom = minidom.parseString(xml_str)
    pretty_xml = dom.toprettyxml(indent="  ", encoding='utf-8')

    return pretty_xml, success_count, failed_count, failed_rows, format_used_count, precision_samples


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
        page_title="Excel to KMZ - Full Precision",
        page_icon="🗺️",
        layout="wide"
    )

    st.title("🗺️ Excel to KMZ - Multi Format (Full Precision)")
    st.markdown("**Decimal Degree, DMS, DM, UTM — auto-detect. Digit koordinat 100% sama dengan excel (raw pass-through).**")

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=['xlsx', 'xls']
    )

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)

            st.success(f"✅ File loaded. Total rows: {len(df)}")

            st.subheader("📊 Data Preview (full digit, tanpa rounding tampilan)")
            # Tampilkan sebagai string supaya tidak ada rounding display
            st.dataframe(df.head(10).astype(str), use_container_width=True)

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
                    f"Mode posisi kolom (urutan X/Y bebas, auto-detect + auto-swap) → "
                    f"`{cols['coord_col_a']}` / `{cols['coord_col_b']}`{name_info}"
                )

            if detected:
                for d in detected:
                    st.write("✅ " + d)
            else:
                st.error("❌ Tidak ada kolom koordinat yang terdeteksi. Cek nama kolom Excel.")
                st.info(f"Kolom ditemukan di file: {', '.join(df.columns.tolist())}")
                return

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
                    kml_content, success, failed, failed_rows, fmt_count, prec_samples = create_kml_content(
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

                    if prec_samples:
                        st.subheader("🔬 Verifikasi Presisi (sample 5 baris pertama)")
                        st.dataframe(pd.DataFrame(prec_samples), use_container_width=True)
                        st.caption("Nilai 'Lon/Lat → KML' persis yang ditulis ke file KMZ. Kalau sumbernya decimal polos, string excel dipakai verbatim.")

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
        | Generic | `X`, `Y` (format isi dideteksi otomatis per baris) |

        ### ✅ Full Precision:
        - Nilai decimal degree dari excel ditulis **verbatim** ke KML (string asli, bukan hasil re-format float)
        - Hasil konversi DMS/DM/UTM ditulis 10 desimal (~0.01 mm)
        - Preview & panel verifikasi presisi full digit
        """)

if __name__ == "__main__":
    main()
