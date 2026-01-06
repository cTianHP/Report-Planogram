import pandas as pd
import streamlit as st
import numpy as np
from io import BytesIO
from datetime import datetime

import os

st.set_page_config(page_title="Report Planogram", page_icon="📑")

current_dir = os.path.dirname(__file__)


# Streamlit App
def process_excel(uploaded_file, Jenis_Lokasi, section, varian, shelve_code, skew, single_rack, posting, settingan_spaceman, tipe_equipment):

    # ===============================
    # 1. LOAD & CLEAN AWAL
    # ===============================
    df = pd.read_excel(uploaded_file, skiprows=6)
    df = df.dropna(axis=1, how='all')        # hapus kolom kosong
    df = df.drop(df.index[-1], errors="ignore")

    if 'PLU' not in df.columns:
        raise KeyError("Kolom 'PLU' tidak ditemukan.")

    df['PLU'] = df['PLU'].astype(str).str.strip()
    df['PLU'] = df['PLU'].replace("", np.nan)
    df = df.dropna(subset=['PLU'])           # Hapus baris PLU kosong

    # Convert kolom numerik
    columns_to_convert = ['Shelv', 'No. Urut', 'PLU', 'KI-KA', 'A-B']
    for column in columns_to_convert:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors='coerce')

    df['No. Urut'] = df['No. Urut'].astype(int)
    df['PLU'] = df['PLU'].astype(int)
    df['KI-KA'] = df['KI-KA'].astype(int)
    df['A-B'] = df['A-B'].astype(int)

    df = df.sort_values(by=['Shelv', 'No. Urut'])

    # ===============================
    # 2. LOAD FILE MAPPING
    # ===============================
    if settingan_spaceman == "inch":
        data_path = os.path.join(current_dir, 'data-inch')
    else:
        data_path = os.path.join(current_dir, 'data-cm')

    # Mapping file lubang
    if Jenis_Lokasi == "T" and section == "EA":
        default_lubang_path = os.path.join(data_path, 'SnackStorehub-lubang.xlsx')
    elif Jenis_Lokasi == "I" and section == "T3":
        default_lubang_path = os.path.join(data_path, 'FPG-Lubang.xlsx')
    elif Jenis_Lokasi == "I" and varian in ["T1C", "T1D"]:
        default_lubang_path = os.path.join(data_path, 'FPG-Lubang.xlsx')
    elif section == "AA":
        default_lubang_path = os.path.join(data_path, 'AA-lubang.xlsx')
    elif Jenis_Lokasi == "I" and section == "AW":
        default_lubang_path = os.path.join(data_path, 'WalkInChiller-lubang.xlsx')
    elif Jenis_Lokasi == "F" and section == "AC":
        default_lubang_path = os.path.join(data_path, 'ChillerFlagship-lubang.xlsx')
    elif Jenis_Lokasi == "I" and varian in ["ACD", "ACE"]:
        default_lubang_path = os.path.join(data_path, 'OpenChiller-lubang.xlsx')
    elif Jenis_Lokasi == "A" and single_rack == "F" and tipe_equipment == "Rak Reguler":
        default_lubang_path = os.path.join(data_path, 'RakDouble-A.xlsx')
    elif Jenis_Lokasi == "A" and single_rack == "T" and tipe_equipment == "Rak Reguler":
        default_lubang_path = os.path.join(data_path, 'RakSingle-A.xlsx')
    elif Jenis_Lokasi == "B" and single_rack == "F" and tipe_equipment == "Rak Reguler":
        default_lubang_path = os.path.join(data_path, 'RakDouble-B.xlsx')
    elif Jenis_Lokasi == "B" and single_rack == "T" and tipe_equipment == "Rak Reguler":
        default_lubang_path = os.path.join(data_path, 'RakSingle-B.xlsx')
    else:
        default_lubang_path = os.path.join(data_path, 'default-lubang.xlsx')

    default_lubang = pd.read_excel(default_lubang_path)
    map_posisi = pd.read_excel(os.path.join(data_path,'map-posisi.xlsx'))

    # ===============================
    # 3. GENERATE rack_number & shelve_number
    # ===============================
    def extract_rack(x):
        if pd.isna(x):
            return None
        parts = str(x).split('.')
        return int(parts[0]) if parts[0].isdigit() else None

    def extract_shelf(x):
        if pd.isna(x):
            return None
        parts = str(x).split('.')
        # jika tidak ada bagian belakang → None
        if len(parts) < 2 or not parts[1].isdigit():
            return None
        return int(parts[1])

    rack_number_series = df['Shelv'].apply(extract_rack)
    shelve_number_series = df['Shelv'].apply(extract_shelf)
    
    rack_number_series = rack_number_series.astype(int)
    shelve_number_series = shelve_number_series.astype(int)


    # ===============================
    # 4. LOGIKA HOLE
    # # ===============================
    if section == "AJ":
        hole_series = shelve_number_series
    else:
        hole_series = df['NOTCHES'].map(
            lambda x: default_lubang.loc[default_lubang['NOTCHES'] == x, 'HOLE'].values[0]
            if x in default_lubang['NOTCHES'].values else "ERROR"
        )

    # # Clean hole jika data tidak valid
    # mask_invalid = (
    #     df['PLU'].isna() |
    #     df['Shelv'].isna() |
    #     shelve_number_series.isna()
    # )
    # hole_series[mask_invalid] = None

    # ===============================
    # 5. LOGIKA shelve_code (Lokasi A + Rak Reguler)
    # ===============================

    # DEFAULT: semua ikut parameter shelve_code
    # shelve_code_series = pd.Series([shelve_code] * len(df))

    # if Jenis_Lokasi == "A" and tipe_equipment == "Rak Reguler":

    #     # Step 1: default semua = 2
    #     shelve_code_series = pd.Series([2] * len(df))

    #     # Step 2: siapkan dataframe temp
    #     df_temp = df.copy()
    #     df_temp["rack_number"] = rack_number_series
    #     df_temp["shelve_number"] = shelve_number_series

    #     # Step 3: drop shelve yang tidak valid sebelum hitung max
    #     df_valid = df_temp.dropna(subset=["shelve_number"])

    #     # Step 4: hitung shelf paling dasar (max per rack)
    #     max_shelf = df_valid.groupby("rack_number")["shelve_number"].transform("max")

    #     # Step 5: gabungkan max shelf ke df_temp by index
    #     df_temp.loc[df_valid.index, "max_shelf"] = max_shelf

    #     # Step 6: assign shelve_code = 1 untuk shelf paling dasar
    #     shelve_code_series[df_temp["shelve_number"] == df_temp["max_shelf"]] = 1

    # else:
    #     shelve_code_series = pd.Series([shelve_code] * len(df))

    # ===============================
    # 6. BENTUK DATA AKHIR
    # ===============================
    data = {
        'location_code': Jenis_Lokasi,
        'section_code': section,
        'variant_code': varian,
        'rack_number': rack_number_series,
        'shelve_number': shelve_number_series,
        'shelve_code': shelve_code,
        'hole': hole_series,
        'skew': skew,
        'single_rack': single_rack,
        'position': df['POSISI'].map(
            lambda x: map_posisi.loc[map_posisi['POSISI'] == x, 'KODE'].values[0]
            if x in map_posisi['POSISI'].values else None
        ),
        'number': df['No. Urut'],
        'plu': df['PLU'],
        'tierkk': df['KI-KA'],
        'tierab': df['A-B'],
        'posting': posting
    }

    data_display = {
        'location_code': Jenis_Lokasi,
        'variant_code': varian,
        'rack_number': rack_number_series,
        'shelve_number': shelve_number_series,
        'number': df['No. Urut'],
        'plu': df['PLU'],
        'desc': df['DESC'],
        'tierkk': df['KI-KA'],
        'tierab': df['A-B']
    }

    new_df = pd.DataFrame(data).dropna()
    display_df = pd.DataFrame(data_display).dropna()

    # ===============================
    # 7. SORT FINAL
    # ===============================
    new_df = new_df.sort_values(
        by=['rack_number', 'shelve_number', 'number']
    ).reset_index(drop=True)

    display_df = display_df.sort_values(
        by=['rack_number', 'shelve_number', 'number']
    ).reset_index(drop=True)

    cols_int = ['rack_number', 'shelve_number', 'number']

    for c in cols_int:
        new_df[c] = pd.to_numeric(new_df[c], errors='coerce').astype('Int64')
        display_df[c] = pd.to_numeric(display_df[c], errors='coerce').astype('Int64')
    
    return new_df, display_df



st.title("Report Planogram App") 

# Input fields for variables
settingan_spaceman = st.selectbox("Settingan Ukuran Spaceman (WAJIB !!!)", ["inch", "cm"], index=0)
Jenis_Lokasi = st.selectbox("Jenis Lokasi", ["A", "B", "I", "T", "F", "G", "X", "Q"], index=0)

tipe_equipment = st.selectbox(
    "Equipment",
    ["Chiller", "Rak Reguler"],
    index=0
)
if tipe_equipment == "Chiller":
    single_rack_value = "F"
else:
    tipe_rak = st.selectbox("Jenis Rak", ["Rak Double", "Rak Single"], index=0)

    if tipe_rak == "Rak Double":
        single_rack_value = "F"
    else:
        single_rack_value = "T"

section = st.text_input("Section", "AC")
varian = st.text_input("Varian", "ACH")
shelve_code = st.number_input("Shelve Code", value=10, step=1)
skew = st.selectbox("Skew", ["F","T"], index=0)
single_rack = st.text_input("Single Rack", value=single_rack_value, disabled=True)
posting = st.text_input("Posting", value="T", disabled=True)


# File Upload
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx','xls'])

if uploaded_file is not None:
    try:
        processed_df, display_df = process_excel(uploaded_file, Jenis_Lokasi, section, varian, shelve_code, skew, single_rack, posting, settingan_spaceman, tipe_equipment)
        
        # ===============================
        # CEK HOLE ERROR + DETAIL LOKASI
        # ===============================
        if 'hole' in processed_df.columns:
            error_rows = processed_df[processed_df['hole'] == "ERROR"]

            if not error_rows.empty:
                lokasi_error = (
                    error_rows[['rack_number', 'shelve_number']]
                    .dropna()
                    .drop_duplicates()
                    .sort_values(['rack_number', 'shelve_number'])
                )

                lokasi_text = ", ".join(
                    f"Rak {int(r)} – Shelving {int(s)}"
                    for r, s in lokasi_error.values
                )

                st.warning(
                    f"⚠️ HOLE ERROR ditemukan pada lokasi berikut:\n"
                    f"{lokasi_text}\n\n"
                    "Silakan periksa NOTCHES atau file mapping lubang yang digunakan."
                )
        
        st.write("Filter Results")
        rack_options = sorted(processed_df['rack_number'].dropna().astype(int).unique())
        shelve_options = sorted(processed_df['shelve_number'].dropna().astype(int).unique())

        selected_racks = st.multiselect("Select Rack Numbers", options=rack_options)
        selected_shelves = st.multiselect("Select Shelve Numbers", options=shelve_options)

        if selected_racks or selected_shelves:
            filtered_df = processed_df[
                (processed_df['rack_number'].isin(selected_racks) if selected_racks else True) &
                (processed_df['shelve_number'].isin(selected_shelves) if selected_shelves else True)
            ]
            display_filtered_df = display_df[
                (display_df['rack_number'].isin(selected_racks) if selected_racks else True) &
                (display_df['shelve_number'].isin(selected_shelves) if selected_shelves else True)
            ]
        else:
            filtered_df = processed_df
            display_filtered_df = display_df

        st.write("Cek Report Planogram: ")
        st.dataframe(display_filtered_df)
        
        st.write("Report Planogram Siap Saji:")
        st.dataframe(filtered_df)

        buffer = BytesIO()
        filtered_df.to_excel(buffer, index=False, engine='xlsxwriter')
        buffer.seek(0)
        
        current_date = datetime.today().date()

        st.download_button(
            label="Download File Report Planogram",
            data=buffer,
            file_name=f"report_planogram_{Jenis_Lokasi}_{varian}_{current_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        
        # ======================================================
        # SECTION BARU - SIMULASI TAMBAH ITEM PLANOGRAM
        # ======================================================
        st.markdown("---")
        st.subheader("🧪 Tambah Item Planogram")

        # ===============================
        # INIT SIMULASI DF
        # ===============================
        if "simulasi_df" not in st.session_state:
            st.session_state.simulasi_df = filtered_df.copy()

        sim_df = st.session_state.simulasi_df

        # ===============================
        # FILTER SIMULASI
        # ===============================
        st.write("Filter Simulasi")

        sim_rack_opts = sorted(sim_df['rack_number'].dropna().astype(int).unique())
        sim_shelf_opts = sorted(sim_df['shelve_number'].dropna().astype(int).unique())

        sim_racks = st.multiselect("Rack (Simulasi)", sim_rack_opts)
        sim_shelves = st.multiselect("Shelving (Simulasi)", sim_shelf_opts)

        if sim_racks or sim_shelves:
            sim_filtered = sim_df[
                (sim_df['rack_number'].isin(sim_racks) if sim_racks else True) &
                (sim_df['shelve_number'].isin(sim_shelves) if sim_shelves else True)
            ]
        else:
            sim_filtered = sim_df

        # ===============================
        # FORM TAMBAH ITEM
        # ===============================
        with st.form("form_simulasi", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)

            with c1:
                in_rack = st.selectbox("Rack", sim_rack_opts)

            with c2:
                shelf_opts = sorted(
                    sim_df[sim_df['rack_number'] == in_rack]['shelve_number']
                    .dropna().astype(int).unique()
                )
                in_shelf = st.selectbox("Shelving", shelf_opts)

            with c3:
                exist_numbers = sim_df[
                    (sim_df['rack_number'] == in_rack) &
                    (sim_df['shelve_number'] == in_shelf)
                ]['number'].dropna().astype(int)

                max_no = exist_numbers.max() if not exist_numbers.empty else 0
                in_number = st.selectbox("Nomor Urut", list(range(1, max_no + 2)))

            in_plu = st.number_input(
                "PLU",
                min_value=1,
                step=1,
                value=1
            )
            in_pos = st.selectbox("Position", ["A", "B", "U", "D", "X", "Y"])

            in_kk = st.number_input(
                "Tier KI-KA",
                min_value=1,
                step=1,
                value=1
            )

            in_ab = st.number_input(
                "Tier A-B",
                min_value=1,
                step=1,
                value=1
            )

            submit_sim = st.form_submit_button("➕ Tambahkan")

        # ===============================
        # SUBMIT LOGIC (AUTO REFRESH)
        # ===============================
        if submit_sim:
            df_new = sim_df.copy()

            # Ambil hole berdasarkan rak & shelving
            hole_val = (
                df_new[
                    (df_new['rack_number'] == in_rack) &
                    (df_new['shelve_number'] == in_shelf)
                ]['hole']
                .dropna()
                .mode()
            )
            hole_val = hole_val.iloc[0] if not hole_val.empty else None

            # Geser nomor urut
            shift_mask = (
                (df_new['rack_number'] == in_rack) &
                (df_new['shelve_number'] == in_shelf) &
                (df_new['number'] >= in_number)
            )
            df_new.loc[shift_mask, 'number'] += 1

            # Logic D → atas jadi U
            if in_pos == "D":
                above_mask = (
                    (df_new['rack_number'] == in_rack) &
                    (df_new['shelve_number'] == in_shelf) &
                    (df_new['number'] == in_number - 1)
                )
                df_new.loc[above_mask, 'position'] = "U"

            # Baris baru
            new_row = {
                'location_code': Jenis_Lokasi,
                'section_code': section,
                'variant_code': varian,
                'rack_number': in_rack,
                'shelve_number': in_shelf,
                'shelve_code': shelve_code,
                'hole': hole_val,
                'skew': skew,
                'single_rack': single_rack,
                'position': in_pos,
                'number': in_number,
                'plu': in_plu if in_plu else None,
                'tierkk': in_kk if in_kk else None,
                'tierab': in_ab if in_ab else None,
                'posting': posting
            }

            st.session_state.simulasi_df = pd.concat(
                [df_new, pd.DataFrame([new_row])],
                ignore_index=True
            )

            st.success("✅ Item berhasil ditambahkan")

            # 🔥 INI KUNCI AUTO REFRESH
            st.rerun()

        # ===============================
        # TAMPILKAN HASIL SIMULASI
        # ===============================
        st.write("Hasil Simulasi Planogram:")
        st.dataframe(
            sim_filtered.sort_values(['rack_number', 'shelve_number', 'number'])
        )


    except Exception as e:
        st.error(f"An error occurred: {e}")
