import pandas as pd
import xml.etree.ElementTree as ET

# Path para o arquivo XML
xml_file_path = r"C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\ncm\ATRIBUTOS_POR_NCM_2024_06_06.xml"

try:
    # Parse XML
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    # Lista para armazenar os dados do NCM 8443.99.80
    data_ncm = []

    # Extrair informações do NCM 8443.99.80
    for ncm in root.findall('.//ncm'):
        codigo_ncm = ncm.find('codigoNcm').text
        if codigo_ncm in ['8443.31.15', '8443.99.80', '8443.99.90', '8443.99.70', '8443.31.99',
                          '8443.99.60', '8443.32.40', '8443.32.99', '8479.89.99',
                          '8443.32.99', '8443.99.39', '8443.99.90', '4911.99.00',
                          '8443.32.36', '8443.31.14']:
            for atributo in ncm.findall('.//atributo'):
                codigo = atributo.find('codigo').text
                modalidade = atributo.find('modalidade').text
                obrigatorio = atributo.find('obrigatorio').text
                multivalorado = atributo.find('multivalorado').text
                data_inicio_vigencia = atributo.find('dataInicioVigencia').text
                data_ncm.append({'Código NCM': codigo_ncm,
                                 'Código Atributo': codigo,
                                 'Modalidade': modalidade,
                                 'Obrigatório': obrigatorio,
                                 'Multivalorado': multivalorado,
                                 'Data Início Vigência': data_inicio_vigencia})

    # Criar DataFrame do Pandas
    df = pd.DataFrame(data_ncm)

    # Caminho para a área de trabalho fornecido
    desktop_path = r"C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop"

    # Construir o caminho para o arquivo Excel
    excel_file_path = desktop_path + "\\ncm_3.xlsx"
    # Exportar para Excel
    df.to_excel(excel_file_path, index=False)

    print(f"NCM 8443.99.80 exportado para '{excel_file_path}' com sucesso!")
    
except FileNotFoundError:
    print("Arquivo XML não encontrado. Verifique o caminho do arquivo.")
except Exception as e:
    print("Ocorreu um erro ao processar o arquivo XML:", e)
