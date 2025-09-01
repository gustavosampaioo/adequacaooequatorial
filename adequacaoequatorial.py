import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import tempfile
import os

def main():
    st.set_page_config(
        page_title="Conversor Excel para KML",
        page_icon="🌍",
        layout="centered"
    )
    
    st.title("🌍 Conversor Excel para KML")
    st.markdown("---")
    
    # Upload do arquivo
    uploaded_file = st.file_uploader(
        "Selecione o arquivo Excel (.xlsx)",
        type=["xlsx"],
        help="O arquivo deve conter as colunas: B (Nome), C (Latitude), D (Longitude), F (Pasta), I (Descrição)"
    )
    
    if uploaded_file is not None:
        # Mostrar informações do arquivo
        st.success(f"Arquivo selecionado: {uploaded_file.name}")
        
        # Botão para gerar KML
        if st.button("🚀 Gerar Arquivo KML", use_container_width=True):
            try:
                # Ler o arquivo Excel
                df = pd.read_excel(uploaded_file)
                
                # Verificar se as colunas necessárias existem
                if len(df.columns) < 9:
                    st.error("❌ O arquivo Excel não tem colunas suficientes! Precisa ter pelo menos 9 colunas.")
                    return
                
                # Criar barra de progresso
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Criar o elemento raiz do KML
                kml = ET.Element("kml", xmlns="http://www.opengis.net/kml/2.2")
                document = ET.SubElement(kml, "Document")
                
                # Dicionário para armazenar pastas (folders)
                folders = {}
                total_rows = len(df)
                
                # Processar cada linha do DataFrame
                for index, row in df.iterrows():
                    # Atualizar progresso
                    progress = (index + 1) / total_rows
                    progress_bar.progress(progress)
                    status_text.text(f"Processando linha {index + 1} de {total_rows}...")
                    
                    # Obter os dados das colunas (index 0-based)
                    nome = str(row.iloc[1]) if pd.notna(row.iloc[1]) else f"Ponto_{index+1}"  # Coluna B (índice 1)
                    lat = row.iloc[2]  # Coluna C (índice 2)
                    lon = row.iloc[3]  # Coluna D (índice 3)
                    folder_name = str(row.iloc[5]) if pd.notna(row.iloc[5]) else "Geral"  # Coluna F (índice 5)
                    description = str(row.iloc[8]) if pd.notna(row.iloc[8]) else ""  # Coluna I (índice 8)
                    
                    # Verificar se latitude e longitude são válidos
                    try:
                        lat = float(lat)
                        lon = float(lon)
                    except (ValueError, TypeError):
                        continue  # Pular linha se coordenadas não forem válidas
                    
                    # Criar folder se não existir
                    if folder_name not in folders:
                        folder = ET.SubElement(document, "Folder")
                        name_element = ET.SubElement(folder, "name")
                        name_element.text = folder_name
                        folders[folder_name] = folder
                    
                    # Criar placemark
                    placemark = ET.SubElement(folders[folder_name], "Placemark")
                    
                    # Nome do placemark
                    name_el = ET.SubElement(placemark, "name")
                    name_el.text = nome
                    
                    # Descrição
                    if description:
                        desc_el = ET.SubElement(placemark, "description")
                        desc_el.text = description
                    
                    # Ponto
                    point = ET.SubElement(placemark, "Point")
                    coordinates = ET.SubElement(point, "coordinates")
                    coordinates.text = f"{lon},{lat},0"
                
                # Converter para string XML formatada
                rough_string = ET.tostring(kml, 'utf-8')
                parsed_string = minidom.parseString(rough_string)
                pretty_xml = parsed_string.toprettyxml(indent="  ")
                
                # Criar arquivo temporário
                with tempfile.NamedTemporaryFile(delete=False, suffix='.kml') as tmp_file:
                    tmp_file.write(pretty_xml.encode('utf-8'))
                    tmp_file_path = tmp_file.name
                
                # Limpar barra de progresso
                progress_bar.empty()
                status_text.empty()
                
                # Mostrar sucesso e botão de download
                st.success("✅ Arquivo KML gerado com sucesso!")
                
                # Botão para download
                with open(tmp_file_path, 'rb') as f:
                    kml_data = f.read()
                
                st.download_button(
                    label="📥 Download do Arquivo KML",
                    data=kml_data,
                    file_name=uploaded_file.name.replace('.xlsx', '.kml'),
                    mime="application/vnd.google-earth.kml+xml",
                    use_container_width=True
                )
                
                # Limpar arquivo temporário após o download
                os.unlink(tmp_file_path)
                
                # Mostrar estatísticas
                st.info(f"""
                **📊 Estatísticas:**
                - Total de pontos processados: {total_rows}
                - Pastas criadas: {len(folders)}
                - Nomes das pastas: {', '.join(folders.keys())}
                """)
                
            except Exception as e:
                st.error(f"❌ Ocorreu um erro ao processar o arquivo:\n{str(e)}")
    
    else:
        # Instruções quando não há arquivo
        st.info("""
        **📋 Instruções:**
        1. Faça upload de um arquivo Excel (.xlsx)
        2. Certifique-se de que o arquivo contenha as seguintes colunas:
           - **Coluna B**: Nome do ponto
           - **Coluna C**: Latitude
           - **Coluna D**: Longitude  
           - **Coluna F**: Nome da pasta (para organização)
           - **Coluna I**: Descrição do ponto
        3. Clique em "Gerar Arquivo KML" para converter
        """)
        
        # Exemplo de estrutura
        with st.expander("📝 Exemplo de estrutura do Excel"):
            st.markdown("""
            | A | B | C | D | E | F | G | H | I |
            |---|---|---|---|---|---|---|---|---|
            | ... | **Nome** | **Latitude** | **Longitude** | ... | **Pasta** | ... | ... | **Descrição** |
            | 1 | Ponto 1 | -23.5505 | -46.6333 | ... | São Paulo | ... | ... | Descrição do ponto 1 |
            | 2 | Ponto 2 | -22.9068 | -43.1729 | ... | Rio de Janeiro | ... | ... | Descrição do ponto 2 |
            | 3 | Ponto 3 | -15.7975 | -47.8919 | ... | Brasília | ... | ... | Descrição do ponto 3 |
            """)

if __name__ == "__main__":
    main()
