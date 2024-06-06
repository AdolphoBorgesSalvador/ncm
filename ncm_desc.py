import xml.etree.ElementTree as ET
import pandas as pd

# Carregar o XML
tree = ET.parse(r"C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\ncm\ATRIBUTOS_POR_NCM_2024_06_06.xml")
root = tree.getroot()

codigos_especificos = [
    'ATT_10278', 'ATT_10734', 'ATT_9241', 'ATT_10888', 'ATT_10156',
    'ATT_2556', 'ATT_8976', 'ATT_10598', 'ATT_9924', 'ATT_10909',
    'ATT_8645', 'ATT_1187', 'ATT_2568', 'ATT_2692', 'ATT_9101',
    'ATT_9764', 'ATT_10691'
]

# Listas para armazenar os dados
codigos = []
nomes = []
orientacoes = []
forma = []

# Iterar sobre cada elemento 'atributo'
for atributo in root.findall('detalhesAtributos/atributo'):
    codigo = atributo.find('codigo').text
    print(f'Checando atributo com código {codigo}')
    if codigo in codigos_especificos:
        print(f'Encontrado código {codigo}')
        nome = atributo.find('nome').text
        print(f'Nome: {nome}')
        orientacao_element = atributo.find('orientacaoPreenchimento')
        orientacao = orientacao_element.text if orientacao_element is not None else None
        print(f'Orientação de preenchimento: {orientacao}')
        forma = atributo.find('formaPreenchimento')
        print(f'forma: {forma}')

        # Adicionar os dados às listas
        codigos.append(codigo)
        nomes.append(nome)
        orientacoes.append(orientacao)
        forma.append(forma)

# Verificar se foram encontrados atributos
if not codigos:
    print('Nenhum atributo encontrado com os códigos específicos.')

# Criar um DataFrame com os dados
data = {'Codigo': codigos, 'Nome': nomes, 'Orientacao de Preenchimento': orientacoes, 'Forma': forma}
df = pd.DataFrame(data)

# Salvar para Excel
df.to_excel('dados_extraidos_4.xlsx', index=False)
