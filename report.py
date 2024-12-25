import pandas as pd
import streamlit as st
from io import BytesIO

import os

st.set_page_config(page_title="Report Planogram", page_icon="📑")

current_dir = os.path.dirname(__file__)
csv_path = os.path.join(current_dir, 'hour.csv')

# Streamlit App
def process_excel(uploaded_file, Jenis_Lokasi, section, varian, shelve_code, skew, single_rack, posting):
    # Load Excel file into a DataFrame
    df = pd.read_excel(uploaded_file)

    # 1. Hapus kolom kosong di antara kolom lainnya
    df = df.dropna(axis=1, how='all')

    # 2. Hapus baris jika terdapat data kosong pada kolom 'PLU'
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
    data_path = os.path.join(current_dir, 'data')
    default_lubang = pd.read_excel(os.path.join(data_path,'default-lubang.xlsx'))  # File untuk mapping hole
    map_posisi = pd.read_excel(os.path.join(data_path,'map-posisi.xlsx'))  # File untuk mapping posisi

    # Ensure required columns exist in the mapping files
    if 'JENLOK' not in default_lubang.columns or 'NOTCHES' not in default_lubang.columns or 'HOLE' not in default_lubang.columns:
        raise KeyError("Kolom 'JENLOK', 'NOTCHES', atau 'HOLE' tidak ditemukan dalam file default-lubang.xlsx")

    if 'POSISI' not in map_posisi.columns or 'KODE' not in map_posisi.columns:
        raise KeyError("Kolom 'POSISI' atau 'KODE' tidak ditemukan dalam file map-posisi.xlsx")

    # Create new DataFrame
    data = {
        'location_code': Jenis_Lokasi,
        'section_code': section,
        'variant_code': varian,
        'rack_number': df['Shelv'].apply(lambda x: int(str(x).split('.')[0]) if pd.notnull(x) else None),
        'shelve_number': df['Shelv'].apply(lambda x: int(str(x).split('.')[1]) if pd.notnull(x) and '.' in str(x) else None),
        'shelve_code': shelve_code,
        'hole': df['NOTCHES'].map(lambda x: default_lubang[(default_lubang['NOTCHES'] == x) & (default_lubang['JENLOK'] == Jenis_Lokasi)]['HOLE'].values[0] if x in default_lubang['NOTCHES'].values else None),
        'skew': skew,
        'single_rack': single_rack,
        'position': df['POSISI'].map(lambda x: map_posisi[map_posisi['POSISI'] == x]['KODE'].values[0] if x in map_posisi['POSISI'].values else None),
        'number': df['No. Urut'],
        'plu': df['PLU'],
        'tierkk': df['KI-KA'],
        'tierab': df['A-B'],
        'posting': posting
    }

    new_df = pd.DataFrame(data)
    new_df = new_df.reset_index()

    return new_df

st.title("Report Planogram App") 

# Input fields for variables
Jenis_Lokasi = st.text_input("Jenis Lokasi", "I")
section = st.text_input("Section", "AC")
varian = st.text_input("Varian", "ACD")
shelve_code = st.number_input("Shelve Code", value=10, step=1)
skew = st.text_input("Skew", "F")
single_rack = st.text_input("Single Rack", "F")
posting = st.text_input("Posting", "T")

# File Upload
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx','xls'])

if uploaded_file is not None:
    try:
        # Process the uploaded file
        processed_df = process_excel(uploaded_file, Jenis_Lokasi, section, varian, shelve_code, skew, single_rack, posting)

        # Display the processed DataFrame
        st.write("Processed DataFrame:")
        st.dataframe(processed_df)

        # Download button for processed DataFrame
        buffer = BytesIO()
        processed_df.to_excel(buffer, index=False, engine='xlsxwriter')
        buffer.seek(0)

        st.download_button(
            label="Download Processed File",
            data=buffer,
            file_name="processed_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"An error occurred: {e}")
