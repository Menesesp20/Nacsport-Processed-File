import streamlit as st
import pandas as pd
import unidecode
from io import BytesIO

def fix_structure(file_path_input, file_path_output):
    input_data = pd.read_excel(file_path_input, sheet_name=None)
    value_mapping = {
        'T_FINALIZAÇÃO': ['FORA', 'FORA ADV', 'NO GOL', 'NO GOL ADV'],
        # Include your full value_mapping as defined earlier
    }
    target_columns = list(value_mapping.keys())
    key_columns = ['N#', 'Categoria', 'Início', 'Click', 'Fim', 'XY']
    def remove_accents(text):
        return unidecode.unidecode(text)

    for sheet_name, sheet_data in input_data.items():
        existing_key_columns = [col for col in key_columns if col in sheet_data.columns]
        des_columns = [col for col in sheet_data.columns if col.startswith('Des')]
        relevant_columns = existing_key_columns + des_columns
        sheet_data = sheet_data[relevant_columns].copy()
        for target_col in target_columns:
            sheet_data[target_col] = ''
        sheet_data['ATLETA'] = ''
        for index, row in sheet_data.iterrows():
            for des_col in des_columns:
                value = row[des_col]
                if pd.notna(value):
                    value_clean = remove_accents(value.strip().lower())
                    found = False
                    for target_col, values_list in value_mapping.items():
                        if value_clean in [remove_accents(v.lower().strip()) for v in values_list]:
                            sheet_data.at[index, target_col] = value
                            found = True
                            break
                    if not found:
                        sheet_data.at[index, 'ATLETA'] = value
        sheet_data = sheet_data.drop(columns=des_columns)
        ordered_columns = existing_key_columns + target_columns + ['ATLETA']
        sheet_data = sheet_data[ordered_columns]
        input_data[sheet_name] = sheet_data

    with pd.ExcelWriter(file_path_output, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in input_data.items():
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
    return file_path_output

# Streamlit App
st.title("Processador de arquivos do Excel")
st.write("Carregue seu arquivo Excel e nós corrigiremos sua estrutura!")

# File upload
uploaded_file = st.file_uploader("Escolha um arquivo do Excel", type=["xlsx"])

if uploaded_file is not None:
    # Process the uploaded file
    st.success("O arquivo foi carregado com sucesso!")
    output_file = BytesIO()
    uploaded_file.seek(0)
    fix_structure(uploaded_file, output_file)
    
    # Provide a download link
    st.download_button(
        label="Download do arquivo processado",
        data=output_file.getvalue(),
        file_name="arquivo_processado_Nacsport.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )