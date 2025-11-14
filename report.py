import pandas as pd
import streamlit as st
import numpy as np
from io import BytesIO
from datetime import datetime

import os

st.set_page_config(page_title="Report Planogram", page_icon="📑")

current_dir = os.path.dirname(__file__)


# Streamlit App
def process_excel(uploaded_file, Jenis_Lokasi, section, varian, shelve_code, skew, single_rack, posting):
    # Load Excel file into a DataFrame
    df = pd.read_excel(uploaded_file, skiprows=6)
    df = df.dropna(axis=1, how='all')  # Hapus kolom kosong
    df = df.drop(df.index[-1])
    # # Pastikan ada cukup baris untuk menghindari kesalahan akses
    # if len(df) <= 6:
    #     raise ValueError("File Excel tidak memiliki cukup data untuk diproses.")
    
    # df = df.iloc[6:].reset_index(drop=True)  # Mulai dari baris ke-7
    # df.columns = df.iloc[0].fillna('Unnamed')  # Gunakan baris pertama sebagai header
    # df = df.drop(df.index[0]).reset_index(drop=True)  # Hapus baris pertama yang sudah digunakan
    
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
    # Load mapping files
    data_path = os.path.join(current_dir, 'data')

    # Load mapping files based on conditions
    if Jenis_Lokasi == "T" and section == "EA":
        default_lubang_path = os.path.join(data_path, 'SnackStorehub-lubang.xlsx')
    elif Jenis_Lokasi == "I" and section == "T3":
        default_lubang_path = os.path.join(data_path, 'FPG-Lubang.xlsx')
    elif Jenis_Lokasi == "I" and ["T1C", "T1D"]:
        default_lubang_path = os.path.join(data_path, 'FPG-Lubang.xlsx')
    elif section == "AA":
        default_lubang_path = os.path.join(data_path, 'AA-lubang.xlsx')
    elif Jenis_Lokasi == "I" and section == "AW":
        default_lubang_path = os.path.join(data_path, 'WalkInChiller-lubang.xlsx')
    elif Jenis_Lokasi == "F" and section == "AC":
        default_lubang_path = os.path.join(data_path, 'ChillerFlagship-lubang.xlsx')
    elif varian in ["ACD", "ACE"]:
        default_lubang_path = os.path.join(data_path, 'OpenChiller-lubang.xlsx')
    else:
        default_lubang_path = os.path.join(data_path, 'default-lubang.xlsx')

    default_lubang = pd.read_excel(default_lubang_path) 
    map_posisi = pd.read_excel(os.path.join(data_path,'map-posisi.xlsx'))  # File untuk mapping posisi

    # # Ensure required columns exist in the mapping files
    # if 'JENLOK' not in default_lubang.columns or 'NOTCHES' not in default_lubang.columns or 'HOLE' not in default_lubang.columns:
    #     raise KeyError("Kolom 'JENLOK', 'NOTCHES', atau 'HOLE' tidak ditemukan dalam file default-lubang.xlsx")

    # if 'POSISI' not in map_posisi.columns or 'KODE' not in map_posisi.columns:
    #     raise KeyError("Kolom 'POSISI' atau 'KODE' tidak ditemukan dalam file map-posisi.xlsx")

    # Create new DataFrame
    data = {
        'location_code': Jenis_Lokasi,
        'section_code': section,
        'variant_code': varian,
        'rack_number': df['Shelv'].apply(lambda x: int(str(x).split('.')[0]) if pd.notnull(x) else None),
        'shelve_number': df['Shelv'].apply(lambda x: int(str(x).split('.')[1]) if pd.notnull(x) and '.' in str(x) else None),
        'shelve_code': shelve_code,
        'hole': df['NOTCHES'].map(lambda x: default_lubang[(default_lubang['NOTCHES'] == x)]['HOLE'].values[0] if x in default_lubang['NOTCHES'].values else None),
        'skew': skew,
        'single_rack': single_rack,
        'position': df['POSISI'].map(lambda x: map_posisi[map_posisi['POSISI'] == x]['KODE'].values[0] if x in map_posisi['POSISI'].values else None),
        'number': df['No. Urut'],
        'plu': df['PLU'],
        'tierkk': df['KI-KA'],
        'tierab': df['A-B'],
        'posting': posting
    }
    
    data_display = {
        'location_code': Jenis_Lokasi,
        'section_code': section,
        'variant_code': varian,
        'rack_number': df['Shelv'].apply(lambda x: int(str(x).split('.')[0]) if pd.notnull(x) else None),
        'shelve_number': df['Shelv'].apply(lambda x: int(str(x).split('.')[1]) if pd.notnull(x) and '.' in str(x) else None),
        # 'shelve_code': shelve_code,
        'hole': df['NOTCHES'].map(lambda x: default_lubang[(default_lubang['NOTCHES'] == x)]['HOLE'].values[0] if x in default_lubang['NOTCHES'].values else None),
        # 'skew': skew,
        # 'single_rack': single_rack,
        # 'position': df['POSISI'].map(lambda x: map_posisi[map_posisi['POSISI'] == x]['KODE'].values[0] if x in map_posisi['POSISI'].values else None),
        'number': df['No. Urut'],
        'plu': df['PLU'],
        'desc': df['DESC'],
        'tierkk': df['KI-KA'],
        'tierab': df['A-B'],
        # 'posting': posting
    }

    new_df = pd.DataFrame(data)
    new_df = new_df.reset_index()
    
    display_df = pd.DataFrame(data_display)
    display_df = display_df.reset_index()

    return new_df, display_df

st.title("Report Planogram App") 

# Input fields for variables
Jenis_Lokasi = st.text_input("Jenis Lokasi", "I")
section = st.text_input("Section", "AB")
varian = st.text_input("Varian", "ABA")
shelve_code = st.number_input("Shelve Code", value=10, step=1)
skew = st.text_input("Skew", "F")
single_rack = st.text_input("Single Rack", "F")
posting = st.text_input("Posting", "T")

# File Upload
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx','xls'])

if uploaded_file is not None:
    try:
        processed_df, display_df = process_excel(uploaded_file, Jenis_Lokasi, section, varian, shelve_code, skew, single_rack, posting)

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
