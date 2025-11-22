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
    # Load Excel file into a DataFrame
    df = pd.read_excel(uploaded_file, skiprows=6)
    df = df.dropna(axis=1, how='all')  # Hapus kolom kosong
    df = df.drop(df.index[-1])
    
    if 'PLU' not in df.columns:
        raise KeyError("Kolom 'PLU' tidak ditemukan. Pastikan header file Excel benar.")

    # 2. Hapus baris jika terdapat data kosong pada kolom 'PLU'
    df['PLU'] = df['PLU'].str.strip()
    df['PLU'] = df['PLU'].replace("", np.nan)
    df = df.dropna(subset=['PLU'])
    

    # 3. Convert kolom tertentu ke dalam format number
    columns_to_convert = ['Shelv', 'No. Urut', 'PLU', 'KI-KA', 'A-B']
    for column in columns_to_convert:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors='coerce')
            if column in ['No. Urut', 'PLU', 'KI-KA', 'A-B']:
                df[column] = df[column].astype(int)

    # 4. Sorting data frame berdasarkan kolom 'Shelv' terlebih dahulu, kemudian 'No. Urut'
    df = df.sort_values(by=['Shelv', 'No. Urut'])

    # Load mapping files
    if settingan_spaceman == "inch":
        data_path = os.path.join(current_dir, 'data-inch')
    elif settingan_spaceman == "cm":
        data_path = os.path.join(current_dir, 'data-cm')

    # Load mapping files based on conditions
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
    map_posisi = pd.read_excel(os.path.join(data_path,'map-posisi.xlsx'))  # File untuk mapping posisi


    # Create new DataFrame
    # --- Generate shelve_number & rack_number ---
    shelve_number_series = df['Shelv'].apply(
        lambda x: int(str(x).split('.')[1]) if pd.notnull(x) and '.' in str(x) else None
    )

    rack_number_series = df['Shelv'].apply(
        lambda x: int(str(x).split('.')[0]) if pd.notnull(x) else None
    )

    # --- Logic Hole Based on Section ---
    if section == "AJ":
        hole_series = shelve_number_series
    else:
        hole_series = df['NOTCHES'].map(
            lambda x: default_lubang[default_lubang['NOTCHES'] == x]['HOLE'].values[0]
            if x in default_lubang['NOTCHES'].values else None
        )


    # -------------------------------------------------------------
    #  LOGIKA BARU UNTUK shelve_code (Lokasi A + Rak Reguler)
    # -------------------------------------------------------------

    # Default → gunakan input user
    shelve_code_series = pd.Series([shelve_code] * len(df))

    if Jenis_Lokasi == "A" and tipe_equipment == "Rak Reguler":

        # Set semua = 2 dulu
        shelve_code_series = pd.Series([2] * len(df))

        # Temp untuk identifikasi leveling dasar
        df_temp = df.copy()
        df_temp["rack_number"] = rack_number_series
        df_temp["shelve_number"] = shelve_number_series

        # Cari shelve_number terbesar per rack
        max_shelve_per_rack = df_temp.groupby("rack_number")["shelve_number"].transform("max")

        # Jika shelve_number == maksimum → shelve_code = 1
        shelve_code_series[df_temp["shelve_number"] == max_shelve_per_rack] = 1


    # -------------------------------------------------------------
    #  FINAL DICTIONARY DATA
    # -------------------------------------------------------------

    data = {
        'location_code': Jenis_Lokasi,
        'section_code': section,
        'variant_code': varian,
        'rack_number': rack_number_series,
        'shelve_number': shelve_number_series,
        'shelve_code': shelve_code_series,  # <-- sudah diperbarui
        'hole': hole_series,
        'skew': skew,
        'single_rack': single_rack,
        'position': df['POSISI'].map(
            lambda x: map_posisi[map_posisi['POSISI'] == x]['KODE'].values[0]
            if x in map_posisi['POSISI'].values else None
        ),
        'number': df['No. Urut'],
        'plu': df['PLU'],
        'tierkk': df['KI-KA'],
        'tierab': df['A-B'],
        'posting': posting
    }


    # -------------------------------------------------------------
    #  DISPLAY DATA (TIDAK BERUBAH)
    # -------------------------------------------------------------

    data_display = {
        'location_code': Jenis_Lokasi,
        'variant_code': varian,
        'rack_number': rack_number_series,
        'shelve_number': shelve_number_series,
        'number': df['No. Urut'],
        'plu': df['PLU'],
        'desc': df['DESC'],
        'tierkk': df['KI-KA'],
        'tierab': df['A-B'],
    }


    new_df = pd.DataFrame(data)
    new_df = new_df.reset_index()
    
    display_df = pd.DataFrame(data_display)
    display_df = display_df.reset_index()

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

        st.write("Filter Results")
        selected_racks = st.multiselect("Select Rack Numbers", options=sorted(processed_df['rack_number'].dropna().unique()))
        selected_shelves = st.multiselect("Select Shelve Numbers", options=sorted(processed_df['shelve_number'].dropna().unique()))

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
        st.dataframe(display_filtered_df.drop(columns=['index']))
        
        st.write("Report Planogram Siap Saji:")
        st.dataframe(filtered_df.drop(columns=['index']))

        buffer = BytesIO()
        filtered_df.drop(columns=['index']).to_excel(buffer, index=False, engine='xlsxwriter')
        buffer.seek(0)
        
        current_date = datetime.today().date()

        st.download_button(
            label="Download File Report Planogram",
            data=buffer,
            file_name=f"report_planogram_{Jenis_Lokasi}_{varian}_{current_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"An error occurred: {e}")
