import pandas as pd
import os
import glob
from datetime import datetime

def main():
    # Define o caminho da pasta Downloads do usuário atual
    pasta_downloads = os.path.join(os.environ['USERPROFILE'], 'Downloads')

    # Encontra todos os arquivos .txt na pasta Downloads
    caminho_arquivos = os.path.join(pasta_downloads, '*.txt')
    arquivos_txt = glob.glob(caminho_arquivos)

    # Lista para armazenar os DataFrames
    dataframes = []

    # Verifica se há arquivos .txt
    if not arquivos_txt:
        print("Nenhum arquivo .txt encontrado.")
        return

    # Cabeçalho das colunas
    colunas = ["Date", "User Name", "Time", "TCode", "Program", "Type", "Ctr", "SF", "ST", "TG", "DK", "Key", "Dynamic key", "Valid from", "Valid to", "Tax Rate", "Tax base", "O", "Txt", "Conv. 100", "FCP Rate", "FCP Base", "Part Exemp", "FCP Resale"]

    # Lê cada arquivo TXT e adiciona à lista de DataFrames
    for arquivo in arquivos_txt:
        try:
            # Lê o arquivo, ignorando as primeiras 4 linhas e sem usar a primeira coluna como índice
            df = pd.read_csv(arquivo, delimiter='|', encoding='utf-8', skiprows=4, header=None, usecols=range(1, 25))
            df.columns = colunas  # Definindo os nomes das colunas
            print(f"Arquivo lido com {len(df)} linhas e {len(df.columns)} colunas: {arquivo}")
            dataframes.append(df)
        except pd.errors.ParserError as e:
            print(f"Erro ao processar o arquivo {arquivo}: {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar o arquivo {arquivo}: {e}")

    # Concatena todos os DataFrames em um único DataFrame
    df_final = pd.concat(dataframes, ignore_index=True) if dataframes else pd.DataFrame()

    # Salva o DataFrame final em um arquivo Excel
    if not df_final.empty:
        agora = datetime.now()
        nome_arquivo = agora.strftime('%Y%m%d_%H%M%S') + '_dados_combinados.xlsx'
        caminho_excel = os.path.join(pasta_downloads, nome_arquivo)
        df_final.to_excel(caminho_excel, index=False)
        print(f"Arquivo Excel criado com sucesso: {caminho_excel}")
        print(f"Salvou {len(df_final)} linhas e {len(df_final.columns)} colunas.")
    else:
        print("DataFrame final está vazio após a concatenação. Nenhum arquivo foi salvo.")

if __name__ == "__main__":
    main()
