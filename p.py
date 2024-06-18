
from io import BytesIO
import pandas as pd
import streamlit as st
import re
import logging

# Definir o layout da página como 'wide' no início do script
st.set_page_config(layout="wide")

# Set up logging
logging.basicConfig(level=logging.DEBUG)

st.title('Converter lista de Endereços para Vazio Sanitário')

def agrupar_por_endereco(df_final):
    df_final['KM'] = pd.to_numeric(df_final['KM'], errors='coerce')
    grouped = df_final.groupby('Endereço')
    data = []

    for name, group in grouped:
        if 'BR 429' in name:
            grouped_by_setor = group.groupby('SETOR')
            for setor, sub_group in grouped_by_setor:
                group_sorted = sub_group.sort_values('KM')
                dados_agrupados = group_sorted.to_dict('records')
        
                data.append({'Endereço': f"{name} - {setor}", 'Dados Agrupados': dados_agrupados})
        else:
            group_sorted = group.sort_values('KM')
            dados_agrupados = group_sorted.to_dict('records')
            data.append({'Endereço': name, 'Dados Agrupados': dados_agrupados})
    df_agrupado = pd.DataFrame(data)
 
    return df_agrupado

def extract_coordinates(text):
    lat_lon_pattern = re.compile(r"Coordenadas:\s*([\d°\s,\.']+)\s*S\s*/\s*([\d°\s,\.']+)\s*W")
    match = lat_lon_pattern.search(text)
    if match:
        latitude = match.group(1).replace(" ", "").replace("'", "")
        longitude = match.group(2).replace(" ", "").replace("'", "")
        return latitude, longitude
    return None, None

def process_dataframe(df):
    """
    Processa o DataFrame para extrair informações e adicionar novas colunas.
    """
    df_final = df.copy()
    df_final['Codigo'] = df_final['Endereço e Informações'].apply(lambda x: re.search(r'\((.+?)\)', x).group(1) if re.search(r'\((.+?)\)', x) else None)
    df_final['Endereço'] = df_final['Endereço e Informações'].apply(lambda x: x.split(')', 1)[1].split(', ')[0].strip() if ')' in x else None)
    df_final['KM'] = pd.to_numeric(df_final['Endereço e Informações'].apply(lambda x: re.search(r'KM\s*(\d+,\d+|\d+)', x).group(1).replace(',', '.') if re.search(r'KM\s*(\d+,\d+|\d+)', x) else None), errors='coerce')
    df_final['SETOR'] = df_final['Endereço e Informações'].apply(lambda x: 'A' if re.search(r'429 SENT/A', x) else 'S' if re.search(r'429 SENT/S', x) else 'S' if re.search(r'429 SE', x) else 'Sem Setor.Cad')
    
    df_final[['Latitude', 'Longitude']] = df_final['Endereço e Informações'].apply(lambda x: pd.Series(extract_coordinates(x)))

    return df_final

def create_excel(df, filename):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    # Converta o DataFrame para uma planilha Excel
    df.to_excel(writer, index=False, sheet_name=filename, startrow=1)

    # Adicione o título acima do cabeçalho
    workbook = writer.book
    worksheet = writer.sheets[filename]
    title = f"Relatório para {filename}"
    worksheet.write(0, 0, title)

    # Configurações de impressão
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 0)
    worksheet.set_print_scale(100)

    # Adicione o efeito zebra
    format1 = workbook.add_format({'bg_color': '#D3D3D3'})
    format2 = workbook.add_format({'bg_color': '#FFFFFF'})
    for row in range(2, len(df) + 2):
        format_to_apply = format1 if row % 2 == 0 else format2
        worksheet.set_row(row, cell_format=format_to_apply)

    writer.close()
    
    output.seek(0)
    return output.getvalue()

def process_and_check_dataframe(df):
    df_final = process_dataframe(df)
    logging.debug("DataFrame final:")
    logging.debug(df_final)
    return df_final

def load_and_display_excel():
    uploaded_file = st.file_uploader("Escolha uma planilha Excel", type=['xlsx'])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        df_final = process_and_check_dataframe(df)
        st.write("DataFrame carregado:")
        
        st.dataframe(df_final)
        
        df_agrupado = agrupar_por_endereco(df_final)
        
        coordenadas_data = []

        for _, row in df_agrupado.iterrows():
            st.subheader(f"Endereço: {row['Endereço']}")
            df_temp = pd.DataFrame(row['Dados Agrupados'])
            st.write(df_temp)
            
            # Selecionar as colunas corretas
            selected_columns = ['Nome', 'Endereço e Informações', 'Nome do proprietario da terra', 'numero']
            df_excel = df_temp[selected_columns]
            
            logging.debug(f"Gerando planilha Excel para o endereço: {row['Endereço']}")
            logging.debug(df_excel)
            excel_bytes = create_excel(df_excel, row['Endereço'])
            
            st.download_button(label=f"Baixar planilha Excel para {row['Endereço']}",
                               data=excel_bytes,
                               file_name=f"{row['Endereço']}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key=f"button_{row['Endereço']}")
            
            for _, item in df_temp.iterrows():
                coordenadas_data.append({
                    'Nome': item['Nome'],
                    'Endereço e Informações': item['Endereço e Informações'],
                    'Nome do proprietario da terra': item['Nome do proprietario da terra'],
                    'numero': item['numero'],
                    'Latitude': item['Latitude'],
                    'Longitude': item['Longitude']
                })

        df_coordenadas = pd.DataFrame(coordenadas_data)
        st.write("Coordenadas gerais:")
        st.dataframe(df_coordenadas)
        
        # Exportar para GAIA (botão após a exibição do DataFrame de coordenadas)
        if st.button("Exportar CSV para GAIA"):
            csv_data = df_coordenadas.to_csv(index=False)
            st.download_button(
                label="Baixar CSV para GAIA",
                data=csv_data,
                file_name="dados_gaia.csv",
                mime="text/csv"
            )

        # Adicionar botão para imprimir o relatório completo em planilhas separadas
        if st.button("Imprimir Relatório Completo (Planilhas Separadas)"):
            excel_bytes_combined = create_combined_excel(df_agrupado)
            st.download_button(
                label="Baixar Relatório Completo",
                data=excel_bytes_combined,
                file_name="relatorio_completo_separado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Adicionar botão para imprimir o relatório completo em uma única planilha
        if st.button("Imprimir Relatório Completo (Uma Planilha)"):
            excel_bytes_single_sheet = create_single_sheet_excel(df_agrupado)
            st.download_button(
                label="Baixar Relatório Completo (Uma Planilha)",
                data=excel_bytes_single_sheet,
                file_name="relatorio_completo_unico.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def create_combined_excel(df_agrupado):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    # Para cada endereço, criar uma nova planilha no Excel
    for _, row in df_agrupado.iterrows():
        df_temp = pd.DataFrame(row['Dados Agrupados'])
        sheet_name = row['Endereço'][:31]  # Limitar nome da planilha a 31 caracteres
        df_temp.to_excel(writer, index=False, sheet_name=sheet_name, startrow=1)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Adicionar título e formatação de impressão
        title = f"Relatório para {row['Endereço']}"
        worksheet.write(0, 0, title)
        worksheet.set_landscape()
        worksheet.fit_to_pages(1, 0)
        worksheet.set_print_scale(100)

        # Adicionar efeito zebra
        format1 = workbook.add_format({'bg_color': '#D3D3D3'})
        format2 = workbook.add_format({'bg_color': '#FFFFFF'})
        for row_num in range(2, len(df_temp) + 2):
            format_to_apply = format1 if row_num % 2 == 0 else format2
            worksheet.set_row(row_num, cell_format=format_to_apply)

    writer.close()
    output.seek(0)
    return output.getvalue()

def create_single_sheet_excel(df_agrupado):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet("Relatório Completo")

    row_num = 0
    for _, row in df_agrupado.iterrows():
        df_temp = pd.DataFrame(row['Dados Agrupados'])
        
        # Adicionar título e cabeçalhos
        worksheet.write(row_num, 0, f"Relatório para {row['Endereço']}")
        df_temp.to_excel(writer, index=False, sheet_name="Relatório Completo", startrow=row_num + 1, header=True)

        # Adicionar quebra de página
        row_num += len(df_temp) + 3
        worksheet.set_row(row_num, None, None, {'hidden': True})  # Adicionar linha em branco como quebra

    # Configurações de impressão
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 0)
    worksheet.set_print_scale(100)

    writer.close()
    output.seek(0)
    return output.getvalue()

# Executa a função principal
if __name__ == "__main__":
    load_and_display_excel()