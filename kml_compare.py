import streamlit as st
import pandas as pd
import zipfile
import io
import os
from xml.etree import ElementTree as ET
from math import radians, cos, sin, asin, sqrt
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="KML/KMZ Overlap Checker", page_icon="🗺️", layout="wide")
KML_NS = "http://www.opengis.net/kml/2.2"


# ── Helpers ──────────────────────────────────────────────────────────────────

def filename_to_label(uploaded_file):
    if uploaded_file is None:
        return ""
    base = os.path.splitext(uploaded_file.name)[0]
    return base.replace("_", " ").replace("-", " ")


def haversine_m(lat1, lon1, lat2, lon2):
    R = 6_371_000
    phi1, phi2 = radians(lat1), radians(lat2)
    a = sin((phi2 - phi1) / 2) ** 2 + cos(phi1) * cos(phi2) * sin((radians(lon2 - lon1)) / 2) ** 2
    return 2 * R * asin(sqrt(a))


def extract_kml_bytes(uploaded_file):
    name = uploaded_file.name.lower()
    raw = uploaded_file.read()
    if name.endswith(".kmz"):
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            kml_names = [n for n in z.namelist() if n.lower().endswith(".kml")]
            if not kml_names:
                st.error(f"Tidak ada file .kml di dalam {uploaded_file.name}")
                return None
            return z.read(kml_names[0])
    return raw


def parse_kml(kml_bytes, label="File"):
    ns = {"kml": KML_NS}
    try:
        root = ET.fromstring(kml_bytes)
    except ET.ParseError as e:
        st.error(f"Gagal parse KML ({label}): {e}")
        return []

    records = []
    for pm in root.findall(".//kml:Placemark", ns):
        name_el = pm.find("kml:name", ns)
        name = name_el.text.strip() if name_el is not None and name_el.text else ""
        ext_data = {}
        schema_data = pm.find(".//kml:SchemaData", ns)
        if schema_data is not None:
            for sd in schema_data.findall("kml:SimpleData", ns):
                k = sd.get("name", "")
                v = sd.text.strip() if sd.text else ""
                ext_data[k] = v
            if not name:
                name = (ext_data.get("NO_SUMUR") or ext_data.get("Name_1")
                        or ext_data.get("name") or "")
        pt = pm.find(".//kml:Point/kml:coordinates", ns)
        if pt is None or not pt.text:
            continue
        parts = pt.text.strip().split(",")
        if len(parts) < 2:
            continue
        try:
            lon, lat = float(parts[0]), float(parts[1])
        except ValueError:
            continue
        records.append({"name": name, "lat": lat, "lon": lon})
    return records


def build_comparison(recs_a, recs_b, threshold_m, label_a, label_b):
    """
    Strict 1-to-1 matching:
    - Build all pairs within threshold
    - Sort by distance (smallest first)
    - Greedy assign: each A max 1 match, each B max 1 match
    
    Result: OVERLAP + HANYA DI = total points (balanced)
    """
    
    # Build all candidate pairs within threshold
    candidates = []
    for i, ra in enumerate(recs_a):
        for j, rb in enumerate(recs_b):
            d = haversine_m(ra["lat"], ra["lon"], rb["lat"], rb["lon"])
            if d <= threshold_m:
                candidates.append((d, i, j))
    
    candidates.sort()  # Sort by distance (shortest first)
    
    # Greedy 1-to-1 assignment
    matched_a = {}  # a_idx → (b_idx, dist)
    matched_b = set()  # b_idx set
    
    for d, i, j in candidates:
        if i not in matched_a and j not in matched_b:
            matched_a[i] = (j, d)
            matched_b.add(j)
    
    # Build rows for A
    rows_a = []
    for i, ra in enumerate(recs_a):
        if i in matched_a:
            j, d = matched_a[i]
            rb = recs_b[j]
            keterangan = "✅ OVERLAP"
        else:
            # Find nearest B for reference (even if beyond threshold)
            best_d, best_j = None, None
            for j, rb in enumerate(recs_b):
                d = haversine_m(ra["lat"], ra["lon"], rb["lat"], rb["lon"])
                if best_d is None or d < best_d:
                    best_d, best_j = d, j
            j, d = best_j, best_d if best_j is not None else (None, None)
            rb = recs_b[j] if j is not None else {"name": "-", "lat": None, "lon": None}
            keterangan = "❌ HANYA DI FILE 1"
        
        if j is not None:
            rows_a.append({
                f"Nama ({label_a})":          ra["name"],
                f"Latitude ({label_a})":      ra["lat"],
                f"Longitude ({label_a})":     ra["lon"],
                f"Nama Pasangan ({label_b})": rb["name"],
                f"Latitude ({label_b})":      rb["lat"],
                f"Longitude ({label_b})":     rb["lon"],
                "Jarak (m)":                  round(d, 2) if d is not None else None,
                "Keterangan":                 keterangan,
            })
    
    df_a = pd.DataFrame(rows_a)
    
    # Build rows for B (unmatched only)
    rows_b = []
    for j, rb in enumerate(recs_b):
        if j not in matched_b:
            # Find nearest A for reference
            best_d, best_i = None, None
            for i, ra in enumerate(recs_a):
                d = haversine_m(rb["lat"], rb["lon"], ra["lat"], ra["lon"])
                if best_d is None or d < best_d:
                    best_d, best_i = d, i
            ra_near = recs_a[best_i] if best_i is not None else {"name": "-", "lat": None, "lon": None}
            
            rows_b.append({
                f"Nama ({label_a})":          ra_near["name"],
                f"Latitude ({label_a})":      ra_near["lat"],
                f"Longitude ({label_a})":     ra_near["lon"],
                f"Nama Pasangan ({label_b})": rb["name"],
                f"Latitude ({label_b})":      rb["lat"],
                f"Longitude ({label_b})":     rb["lon"],
                "Jarak (m)":                  round(best_d, 2) if best_d is not None else None,
                "Keterangan":                 "⚠️ HANYA DI FILE 2",
            })
    
    df_b = pd.DataFrame(rows_b)
    
    return df_a, df_b


def safe_sheet_name(name, max_len=28):
    for ch in r'\/*?:[]':
        name = name.replace(ch, "")
    return name[:max_len]


def to_excel_bytes(df_overlap, df_not_a, df_only_b, label_a, label_b, threshold_m, n_a, n_b):
    wb = Workbook()
    GREEN      = "C6EFCE"
    RED        = "FFC7CE"
    YELLOW     = "FFEB9C"
    BLUE_HDR   = "1F4E79"
    GRAY_HDR   = "595959"
    WHITE      = "FFFFFF"
    LIGHT_GRAY = "F2F2F2"
    thin   = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write_sheet(ws, df, sheet_title, note=""):
        col_widths = {}
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
        tc = ws.cell(row=1, column=1, value=sheet_title)
        tc.font = Font(bold=True, size=13, color=WHITE)
        tc.fill = PatternFill("solid", fgColor=BLUE_HDR)
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        header_row = 2
        if note:
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(df.columns))
            nc = ws.cell(row=2, column=1, value=note)
            nc.font = Font(italic=True, size=10, color="444444")
            nc.fill = PatternFill("solid", fgColor="EAF1F8")
            nc.alignment = Alignment(horizontal="left")
            header_row = 3

        for ci, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=header_row, column=ci, value=col_name)
            cell.font = Font(bold=True, size=10, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=GRAY_HDR)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            col_widths[ci] = max(len(str(col_name)), col_widths.get(ci, 0))
        ws.row_dimensions[header_row].height = 32

        ket_col = list(df.columns).index("Keterangan") + 1 if "Keterangan" in df.columns else None

        for ri, row in enumerate(df.itertuples(index=False), start=header_row + 1):
            row_bg = LIGHT_GRAY if ri % 2 == 0 else WHITE
            for ci, value in enumerate(row, 1):
                cell = ws.cell(row=ri, column=ci, value=value)
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(size=9)
                if ci == ket_col:
                    val_str = str(value) if value else ""
                    if "OVERLAP" in val_str and "TIDAK" not in val_str:
                        cell.fill = PatternFill("solid", fgColor=GREEN)
                        cell.font = Font(size=9, bold=True, color="006100")
                    elif "TIDAK" in val_str:
                        cell.fill = PatternFill("solid", fgColor=RED)
                        cell.font = Font(size=9, bold=True, color="9C0006")
                    elif "HANYA" in val_str:
                        cell.fill = PatternFill("solid", fgColor=YELLOW)
                        cell.font = Font(size=9, bold=True, color="7D5800")
                    else:
                        cell.fill = PatternFill("solid", fgColor=row_bg)
                else:
                    cell.fill = PatternFill("solid", fgColor=row_bg)
                col_widths[ci] = max(len(str(value)) if value is not None else 0, col_widths.get(ci, 0))

        for ci, width in col_widths.items():
            ws.column_dimensions[get_column_letter(ci)].width = min(max(width + 4, 12), 40)

    # Sheet 1: Ringkasan
    ws_sum = wb.active
    ws_sum.title = "Ringkasan"
    ws_sum.merge_cells("A1:B1")
    tc = ws_sum["A1"]
    tc.value = f"PERBANDINGAN: {label_a}  ×  {label_b}"
    tc.font = Font(bold=True, size=14, color=WHITE)
    tc.fill = PatternFill("solid", fgColor=BLUE_HDR)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 28

    summary_rows = [
        (f"Total koordinat {label_a}",                                    n_a),
        (f"Total koordinat {label_b}",                                    n_b),
        ("Threshold overlap (meter)",                                      threshold_m),
        ("Jumlah OVERLAP (1-to-1 matching)",                              len(df_overlap)),
        (f"Hanya di {label_a} (tidak ada pasangan di {label_b})",         len(df_not_a)),
        (f"Hanya di {label_b} (jarak > threshold)",                       len(df_only_b)),
        ("Maks. overlap teoritis = min(File 1, File 2)",                   min(n_a, n_b)),
    ]
    for ri, (k, v) in enumerate(summary_rows, start=2):
        ck = ws_sum.cell(row=ri, column=1, value=k)
        cv = ws_sum.cell(row=ri, column=2, value=v)
        bg = LIGHT_GRAY if ri % 2 == 0 else WHITE
        for c in [ck, cv]:
            c.border = border
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="left" if c.column == 1 else "center", vertical="center")
            c.font = Font(size=10)
        ck.font = Font(size=10, bold=True)
    ws_sum.column_dimensions["A"].width = 58
    ws_sum.column_dimensions["B"].width = 18

    # Sheet 2: OVERLAP
    if len(df_overlap):
        ws_ov = wb.create_sheet(safe_sheet_name("OVERLAP"))
        write_sheet(ws_ov, df_overlap,
                    f"KOORDINAT OVERLAP — {label_a} × {label_b}",
                    f"Threshold: {threshold_m} m  |  Total: {len(df_overlap)} titik  |  Metode: Strict 1-to-1 greedy matching")

    # Sheet 3: Hanya di A
    if len(df_not_a):
        ws_na = wb.create_sheet(safe_sheet_name(f"Hanya {label_a}"))
        write_sheet(ws_na, df_not_a,
                    f"HANYA DI {label_a.upper()} — tidak ada pasangan di {label_b}",
                    f"Total: {len(df_not_a)} titik  |  Jarak > {threshold_m}m ke nearest B")

    # Sheet 4: Hanya di B
    if len(df_only_b):
        ws_nb = wb.create_sheet(safe_sheet_name(f"Hanya {label_b}"))
        write_sheet(ws_nb, df_only_b,
                    f"HANYA DI {label_b.upper()} — tidak ada pasangan di {label_a}",
                    f"Total: {len(df_only_b)} titik  |  Tidak match dalam 1-to-1 assignment")

    # Sheet 5: Semua Data
    ws_all = wb.create_sheet("Semua Data")
    df_all = pd.concat([df_overlap, df_not_a, df_only_b], ignore_index=True)
    order = {"✅ OVERLAP": 0, "❌ HANYA DI FILE 1": 1, "⚠️ HANYA DI FILE 2": 2}
    df_all["_sort"] = df_all["Keterangan"].map(lambda x: order.get(x, 3))
    df_all = df_all.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
    write_sheet(ws_all, df_all,
                f"SEMUA DATA — {label_a} × {label_b}",
                f"Threshold: {threshold_m} m  |  Total baris: {len(df_all)}")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── UI ───────────────────────────────────────────────────────────────────────

st.title("🗺️ KML / KMZ Overlap Checker")
st.caption("Upload 2 file KML/KMZ → deteksi titik tumpang-tindih (1-to-1) → export Excel.")
st.info("✨ **STRICT 1-to-1 Matching:** Setiap titik hanya match ke 1 titik lainnya. OVERLAP + HANYA DI = total koordinat ✓")

with st.sidebar:
    st.header("⚙️ Pengaturan")
    threshold_m = st.number_input(
        "Threshold jarak overlap (meter)",
        min_value=1, max_value=10_000, value=50, step=1,
        help="Dua titik dianggap overlap jika jarak ≤ nilai ini"
    )
    st.markdown("---")
    st.markdown(
        "**Cara pakai:**\n"
        "1. Upload File 1 & File 2\n"
        "2. Klik **Proses**\n"
        "3. Download Excel\n\n"
        "**Metode matching:**\n"
        "- Build semua pairs dalam threshold\n"
        "- Sort by distance (terpendek dulu)\n"
        "- Greedy assign: 1 titik = 1 match\n\n"
        "**Hasil balanced:**\n"
        "OVERLAP + HANYA DI = total points ✓"
    )

col1, col2 = st.columns(2)
with col1:
    st.subheader("📂 File 1")
    file_a = st.file_uploader("Upload KML / KMZ", type=["kml", "kmz"], key="file_a")
    default_a = filename_to_label(file_a) if file_a else "File 1"
    label_a = st.text_input("Label", value=default_a, key="label_a")

with col2:
    st.subheader("📂 File 2")
    file_b = st.file_uploader("Upload KML / KMZ", type=["kml", "kmz"], key="file_b")
    default_b = filename_to_label(file_b) if file_b else "File 2"
    label_b = st.text_input("Label", value=default_b, key="label_b")

if st.button("🔍 Proses Perbandingan", type="primary", use_container_width=True):

    if not file_a or not file_b:
        st.warning("Upload dua file dulu.")
        st.stop()

    label_a = label_a.strip() or filename_to_label(file_a) or "File 1"
    label_b = label_b.strip() or filename_to_label(file_b) or "File 2"

    with st.spinner("Parsing KML..."):
        kml_bytes_a = extract_kml_bytes(file_a)
        kml_bytes_b = extract_kml_bytes(file_b)
        if kml_bytes_a is None or kml_bytes_b is None:
            st.stop()
        recs_a = parse_kml(kml_bytes_a, label_a)
        recs_b = parse_kml(kml_bytes_b, label_b)

    if not recs_a:
        st.error(f"Tidak ada titik terbaca dari {label_a}.")
        st.stop()
    if not recs_b:
        st.error(f"Tidak ada titik terbaca dari {label_b}.")
        st.stop()

    st.success(f"**{label_a}**: {len(recs_a)} titik  |  **{label_b}**: {len(recs_b)} titik  |  Max overlap teoritis: **{min(len(recs_a), len(recs_b))}**")

    with st.spinner("Menghitung overlap (1-to-1 matching)..."):
        df_a, df_b_unmatched = build_comparison(recs_a, recs_b, threshold_m, label_a, label_b)
        df_overlap = df_a[df_a["Keterangan"] == "✅ OVERLAP"].reset_index(drop=True)
        df_only_a = df_a[df_a["Keterangan"] == "❌ HANYA DI FILE 1"].reset_index(drop=True)
        df_only_b = df_b_unmatched  # All unmatched File 2

    # Sanity check - math must balance
    assert len(df_overlap) + len(df_only_a) == len(recs_a), f"BUG: {len(df_overlap)} + {len(df_only_a)} != {len(recs_a)}"
    assert len(df_overlap) + len(df_only_b) == len(recs_b), f"BUG: {len(df_overlap)} + {len(df_only_b)} != {len(recs_b)}"
    
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric(f"Total {label_a}", len(recs_a))
    m2.metric(f"Total {label_b}", len(recs_b))
    m3.metric("✅ OVERLAP", len(df_overlap))
    m4.metric(f"❌ Hanya {label_a}", len(df_only_a))
    m5.metric(f"⚠️ Hanya {label_b}", len(df_only_b))

    st.markdown("---")

    tab1, tab2, tab3 = st.tabs([
        f"✅ OVERLAP ({len(df_overlap)})",
        f"❌ Hanya {label_a} ({len(df_only_a)})",
        f"⚠️ Hanya {label_b} ({len(df_only_b)})",
    ])
    with tab1:
        st.dataframe(df_overlap, use_container_width=True) if len(df_overlap) else st.info("Tidak ada titik overlap.")
    with tab2:
        st.dataframe(df_only_a, use_container_width=True) if len(df_only_a) else st.info(f"Semua titik {label_a} punya pasangan.")
    with tab3:
        st.dataframe(df_only_b, use_container_width=True) if len(df_only_b) else st.info(f"Semua titik {label_b} punya pasangan.")

    st.markdown("---")

    with st.spinner("Generate Excel..."):
        excel_bytes = to_excel_bytes(
            df_overlap, df_only_a, df_only_b,
            label_a, label_b, threshold_m,
            len(recs_a), len(recs_b)
        )

    fn = f"Compare_{label_a}_vs_{label_b}.xlsx".replace(" ", "_")
    st.download_button(
        label="📥 Download Excel",
        data=excel_bytes,
        file_name=fn,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
