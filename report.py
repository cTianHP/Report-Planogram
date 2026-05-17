import pandas as pd
import streamlit as st
import numpy as np
from io import BytesIO
from datetime import datetime

import os

st.set_page_config(page_title="Report Planogram", page_icon="📑")

current_dir = os.path.dirname(__file__)


# Streamlit App
def process_excel(uploaded_file, kode_cabang, Jenis_Lokasi, section, varian, shelve_code, skew, single_rack, settingan_spaceman, tipe_equipment, jumlah_rak_roti=None):

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
    # elif section == "AA":
    #     default_lubang_path = os.path.join(data_path, 'AA-lubang.xlsx')
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
    elif tipe_equipment == "Standing Freezer":
        default_lubang_path = os.path.join(data_path, 'lubang-STD Freezer.xlsx')
    elif tipe_equipment == "Rak Roti":
        if jumlah_rak_roti is None:
            raise ValueError("jumlah_rak_roti harus diisi untuk Rak Roti")
        if Jenis_Lokasi == "A":
            default_lubang_path = os.path.join(
                data_path,
                f'lubang-Rak Roti-A-{jumlah_rak_roti}.xlsx'
            )
        elif Jenis_Lokasi == "B":
            default_lubang_path = os.path.join(
                data_path,
                f'lubang-Rak Roti-B-{jumlah_rak_roti}.xlsx'
            )
        else:
            raise ValueError(f"Rak Roti tidak tersedia untuk lokasi {Jenis_Lokasi}")
    else:
        default_lubang_path = os.path.join(data_path, 'default-lubang.xlsx')

    default_lubang = pd.read_excel(default_lubang_path)
    # ===============================
    # TAMBAH KODE DI FILE LUBANG
    # ===============================
    if tipe_equipment in ["Standing Freezer", "Rak Roti"]:
        default_lubang['Kode'] = (
            default_lubang['Rak'].astype(str) + "-" +
            default_lubang['Shelving'].astype(str)
        )
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
    
    kode_series = (
        rack_number_series.astype(str) + "-" +
        shelve_number_series.astype(str)
    )


    # ===============================
    # 4. LOGIKA HOLE
    # # ===============================
    if tipe_equipment in ["Standing Freezer", "Rak Roti"]:
        hole_series = kode_series.map(
            lambda x: default_lubang.loc[
                default_lubang['Kode'] == x, 'Lubang'
            ].values[0]
            if x in default_lubang['Kode'].values else "ERROR"
        )

    elif section == "AJ":
        hole_series = shelve_number_series

    else:
        hole_series = df['NOTCHES'].map(
            lambda x: default_lubang.loc[
                default_lubang['NOTCHES'] == x, 'HOLE'
            ].values[0]
            if x in default_lubang['NOTCHES'].values else "ERROR"
        )

    # ===============================
    # 6. BENTUK DATA AKHIR
    # ===============================
    data = {
        'kode_cabang': kode_cabang,
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
        'tierab': df['A-B']
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
kode_cabang = st.text_input("Kode Cabang", value="KZ01")
Jenis_Lokasi = st.selectbox("Jenis Lokasi", ["A", "B", "I", "T", "F", "G", "X", "Q"], index=0)

tipe_equipment = st.selectbox(
    "Equipment",
    ["Chiller","Standing Freezer","Freezer Nugget","Freezer Ice Cream", "Rak Roti", "Rak Reguler"],
    index=0
)
single_rack_value = "F"
jumlah_rak_roti = None

if tipe_equipment == "Rak Reguler":
    tipe_rak = st.selectbox("Jenis Rak", ["Rak Double", "Rak Single"], index=0)

    if tipe_rak == "Rak Double":
        single_rack_value = "F"
    else:
        single_rack_value = "T"

elif tipe_equipment == "Rak Roti":
    jumlah_rak_roti = st.selectbox("Jumlah Rak Roti",[1,2,3], index=0)
    single_rack_value = "F"  

section = st.text_input("Section", "AC")
varian = st.text_input("Varian", "ACH")
shelve_code = st.number_input("Shelve Code", value=10, step=1)
skew = st.selectbox("Skew", ["F","T"], index=0)
single_rack = st.text_input("Single Rack", value=single_rack_value, disabled=True)


# File Upload
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx','xls'])

if uploaded_file is not None:
    try:
        processed_df, display_df = process_excel(
            uploaded_file, kode_cabang, Jenis_Lokasi, section, varian, shelve_code, skew, single_rack, settingan_spaceman, tipe_equipment, jumlah_rak_roti if tipe_equipment == "Rak Roti" else None
        )
        
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

        # 1. Pastikan opsi dikonversi ke standard Python list (.tolist())
        rack_options = sorted(processed_df['rack_number'].dropna().astype(int).unique().tolist())

        selected_racks = st.multiselect(
            "Select Rack Numbers",
            options=rack_options
        )

        # 2. FILTER BERTINGKAT: Opsi Shelve menyesuaikan Rack yang dipilih
        if selected_racks:
            # Jika ada rak yang dipilih, ambil opsi shelve hanya dari rak tersebut
            temp_df_for_shelve = processed_df[processed_df['rack_number'].isin(list(selected_racks))]
        else:
            # Jika tidak ada rak yang dipilih, tampilkan semua opsi shelve
            temp_df_for_shelve = processed_df

        shelve_options = sorted(temp_df_for_shelve['shelve_number'].dropna().astype(int).unique().tolist())

        selected_shelves = st.multiselect(
            "Select Shelve Numbers",
            options=shelve_options
        )

        # ===============================
        # FILTER DATA UTAMA
        # ===============================
        filtered_df = processed_df.copy()
        display_filtered_df = display_df.copy()

        # Gunakan list() di dalam isin() untuk mencegah error mismatch tipe data Pandas
        if selected_racks:
            filtered_df = filtered_df[filtered_df['rack_number'].isin(list(selected_racks))]
            display_filtered_df = display_filtered_df[display_filtered_df['rack_number'].isin(list(selected_racks))]

        if selected_shelves:
            filtered_df = filtered_df[filtered_df['shelve_number'].isin(list(selected_shelves))]
            display_filtered_df = display_filtered_df[display_filtered_df['shelve_number'].isin(list(selected_shelves))]

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
        # SECTION BARU - SIMULASI TAMBAH ITEM PLANOGRAM (INPUT MANUAL)
        # ======================================================
        st.markdown("---")
        st.subheader("🧪 Tambah Item Planogram")

        # ===============================
        # INIT SIMULASI DF (ambil dari filtered_df saat ini)
        # ===============================
        if "simulasi_df" not in st.session_state:
            sim_init = filtered_df.copy()
            sim_init['kode_cabang'] = kode_cabang

            # pastikan di depan
            cols = ['kode_cabang'] + [c for c in sim_init.columns if c != 'kode_cabang']
            sim_init = sim_init[cols]

            st.session_state.simulasi_df = sim_init
        sim_df = st.session_state.simulasi_df

        # ===============================
        # FILTER SIMULASI (hanya untuk melihat)
        # ===============================
        st.write("Filter Simulasi (untuk preview)")
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
        # FORM INPUT (SEMUA MANUAL, TANPA SELECTBOX)
        # ===============================
        st.write("Form Tambah Item (Isi manual semua field; klik Tambah untuk menyisipkan)")
        with st.form("form_simulasi_manual", clear_on_submit=True):
            r_col1, r_col2, r_col3 = st.columns(3)

            with r_col1:
                in_rack_txt = st.text_input("Rack (angka)", placeholder="contoh: 1")
            with r_col2:
                in_shelf_txt = st.text_input("Shelving (angka)", placeholder="contoh: 4")
            with r_col3:
                in_number_txt = st.text_input("Nomor Urut (angka)", placeholder="contoh: 3")

            p_col1, p_col2 = st.columns(2)
            with p_col1:
                in_plu_txt = st.text_input("PLU (angka)", placeholder="contoh: 417806")
            with p_col2:
                in_pos = st.text_input("Position (mis. U/D/A/B/X/Y)", placeholder="mis: D")

            t_col1, t_col2 = st.columns(2)
            with t_col1:
                in_kk_txt = st.text_input("Tier KI-KA (angka)", value="1")
            with t_col2:
                in_ab_txt = st.text_input("Tier A-B (angka)", value="1")

            submit_sim = st.form_submit_button("➕ Tambahkan")

        # ===============================
        # PROSES SUBMIT (VALIDASI + INSERT)
        # ===============================
        if submit_sim:
            # --- VALIDASI: wajib numeric untuk fields yang dibutuhkan ---
            errors = []
            # validate rack
            if not in_rack_txt or not in_rack_txt.strip().isdigit():
                errors.append("Rack harus diisi dengan angka.")
            if not in_shelf_txt or not in_shelf_txt.strip().isdigit():
                errors.append("Shelving harus diisi dengan angka.")
            if not in_number_txt or not in_number_txt.strip().isdigit():
                errors.append("Nomor Urut harus diisi dengan angka.")
            if not in_plu_txt or not in_plu_txt.strip().isdigit():
                errors.append("PLU harus diisi dengan angka.")
            # tiers optional but we try to parse; default handled below
            if in_kk_txt and (not in_kk_txt.strip().isdigit()):
                errors.append("Tier KI-KA harus angka jika diisi.")
            if in_ab_txt and (not in_ab_txt.strip().isdigit()):
                errors.append("Tier A-B harus angka jika diisi.")

            if errors:
                st.error(" • ".join(errors))
            else:
                # parse to int
                in_rack = int(in_rack_txt.strip())
                in_shelf = int(in_shelf_txt.strip())
                in_number = int(in_number_txt.strip())
                in_plu = int(in_plu_txt.strip())
                in_kk = int(in_kk_txt.strip()) if in_kk_txt.strip() else 1
                in_ab = int(in_ab_txt.strip()) if in_ab_txt.strip() else 1
                in_position = in_pos.strip() if in_pos and in_pos.strip() else None

                # get current sim df fresh
                df_new = st.session_state.simulasi_df.copy()
                # ensure number column int
                df_new['number'] = df_new['number'].astype(int)

                # determine hole for that rack+shelf (if exists)
                hole_series = df_new[
                    (df_new['rack_number'] == in_rack) &
                    (df_new['shelve_number'] == in_shelf)
                ]['hole'].dropna().mode()
                hole_val = hole_series.iloc[0] if not hole_series.empty else None

                # shift numbers for that rack+shelf
                shift_mask = (
                    (df_new['rack_number'] == in_rack) &
                    (df_new['shelve_number'] == in_shelf) &
                    (df_new['number'] >= in_number)
                )
                df_new.loc[shift_mask, 'number'] += 1

                # special logic: if new position is D then item above becomes U
                if in_position == "D" and in_number > 1:
                    above_mask = (
                        (df_new['rack_number'] == in_rack) &
                        (df_new['shelve_number'] == in_shelf) &
                        (df_new['number'] == in_number - 1)
                    )
                    df_new.loc[above_mask, 'position'] = "U"

                # build new row (allow None for optionals)
                new_row = {
                    'kode_cabang': kode_cabang,
                    'location_code': Jenis_Lokasi,
                    'section_code': section,
                    'variant_code': varian,
                    'rack_number': in_rack,
                    'shelve_number': in_shelf,
                    'shelve_code': shelve_code,
                    'hole': hole_val,
                    'skew': skew,
                    'single_rack': single_rack,
                    'position': in_position,
                    'number': int(in_number),
                    'plu': int(in_plu),
                    'tierkk': int(in_kk),
                    'tierab': int(in_ab)
                }

                # append and re-sort
                df_new = pd.concat([df_new, pd.DataFrame([new_row])], ignore_index=True)
                df_new = df_new.sort_values(['rack_number', 'shelve_number', 'number']).reset_index(drop=True)

                # save back
                st.session_state.simulasi_df = df_new

                st.success("✅ Item berhasil disisipkan ke simulasi pada posisi yang dipilih. Form telah direset.")

                # Streamlit akan rerun setelah form submit with clear_on_submit=True,
                # jadi form inputs akan tampil kosong (sesuai permintaan).

        # ===============================
        # TAMPILAN HASIL SIMULASI (SETELAH SUBMIT AKAN TERUPDATE)
        # ===============================
        st.write("Hasil Simulasi Planogram:")
        
        # 1. Ambil data fresh dari session state
        sim_df = st.session_state.simulasi_df.copy()
        
        # 2. Filter bertahap: Lebih aman, mudah dibaca, dan bebas dari error bitwise
        if sim_racks:
            sim_df = sim_df[sim_df['rack_number'].isin(list(sim_racks))]
            
        if sim_shelves:
            sim_df = sim_df[sim_df['shelve_number'].isin(list(sim_shelves))]
            
        # Simpan hasil akhir ke sim_filtered
        sim_filtered = sim_df

        # 3. Urutkan dan tampilkan data
        # Ditambahkan penanganan error jika seandainya nama kolom 'number' berbeda
        try:
            sim_filtered_sorted = sim_filtered.sort_values(['rack_number', 'shelve_number', 'number'])
            st.dataframe(sim_filtered_sorted)
        except KeyError:
            # Fallback jika kolom 'number' tidak ditemukan
            sim_filtered_sorted = sim_filtered.sort_values(['rack_number', 'shelve_number'])
            st.dataframe(sim_filtered_sorted)



    except Exception as e:
        st.error(f"An error occurred: {e}")
