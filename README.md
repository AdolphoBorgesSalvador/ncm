# Projeto de Extração e Manipulação de Dados de NCM

Este projeto automatiza a extração de informações específicas de arquivos XML relacionados a NCM (Nomenclatura Comum do Mercosul). Ele foi desenvolvido para facilitar o trabalho com grandes volumes de dados, consolidando informações em arquivos Excel.

---

## Estrutura do Projeto

### Principais Scripts

1. **`ncm_desc.py`**
   - Extrai informações específicas de atributos dentro de um arquivo XML de NCM.
   - Procura por atributos específicos definidos em uma lista e coleta informações como nome, orientação de preenchimento e forma.
   - Gera um arquivo Excel (`dados_extraidos_4.xlsx`) com os dados extraídos.

2. **`ncm_extraidos.py`**
   - Focado na extração de informações específicas relacionadas a códigos NCM selecionados.
   - Extrai atributos e detalhes como modalidade, obrigatoriedade, multivalorado e data de início de vigência.
   - Exporta os dados processados para um arquivo Excel (`ncm_3.xlsx`).

---

## Funcionalidades

1. **Extração de Dados XML:**
   - Leitura e parsing de arquivos XML grandes.
   - Busca de atributos e NCMs específicos com base em filtros definidos.

2. **Automatização da Exportação:**
   - Criação de arquivos Excel organizados com os dados extraídos.
   - Otimização do processo para trabalhar com vários códigos de NCM.

3. **Tratamento de Exceções:**
   - Verificação de caminhos inválidos para evitar erros.
   - Tratamento de possíveis problemas durante a leitura do XML.

---

## Como Usar

### Requisitos

- **Python**: 3.x
- **Bibliotecas necessárias**:
  - `pandas`
  - `openpyxl`
  - `xml.etree.ElementTree` (padrão no Python)

##
