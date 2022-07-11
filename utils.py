import pandas as pd # Manipulação e análise de dados.
from bs4 import BeautifulSoup # Analisar documentos HTML e XML.

# Função que lê a página web pega os dados das tabelas e salva em ficheiros do tipo excel (xlsx)
def ConvertTableToExcel(filename, elements):

    # Necessário para fazer a escrita do ficheiro xlsx
    writer = pd.ExcelWriter(filename+".xlsx", engine = 'xlsxwriter')

    # Para cada tabela presente na página
    for idx, element in enumerate(elements):

        # Pegar o código HTML dentro do elemento selecionado + HTML do elemento selecionado
        html_content = element.get_attribute('outerHTML')

        # Convert o código HTML em estrutura de dado
        soup = BeautifulSoup(html_content, 'html.parser')

        # Localizar a tabela no meio desse código
        table = soup.find(name='table')

        # Converter a tabela em HTML para DataFrame
        df_full = pd.read_html(str(table))[0]

        # Escrever o que está no DataFrame para o ficheiro xlsx
        if 'Ano' in df_full.columns:
            df_full.to_excel(writer, sheet_name='{0} ano'.format(df_full['Ano'].values[0]), index=False, na_rep='-')
        else:
            df_full.to_excel(writer, sheet_name='{0}º ano'.format(idx+1), index=False, na_rep='-')

        # Ajuste automático da largura das colunas 
        for column in df_full:
            column_length = max(df_full[column].astype(str).map(len).max(), len(str(column)))
            col_idx = df_full.columns.get_loc(column)
            if 'Ano' in df_full.columns:
                writer.sheets['{0} ano'.format(df_full['Ano'].values[0])].set_column(col_idx, col_idx, column_length+2)
            else:
                writer.sheets['{0}º ano'.format(idx+1)].set_column(col_idx, col_idx, column_length+2)
            
    # Salvar as alterações no ficheiro xlsx
    writer.save()