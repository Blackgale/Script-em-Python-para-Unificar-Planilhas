import pandas as pd
import glob
import os

# Obter lista de arquivos Excel
path = input("Digite o caminho do diretório onde estão suas planilhas: ")
extension = input("Digite a extensão dos arquivos Excel (por exemplo, xlsx): ")
all_files = glob.glob(os.path.join(path, f"*.{extension}"))

# Criar uma lista vazia para armazenar as planilhas
dfs = []

# Ler cada arquivo Excel e adicioná-lo à lista
for file in all_files:
    sheet_name = input(f"Digite o nome da planilha que deseja unificar do arquivo {file}: ")
    df = pd.read_excel(file, sheet_name=sheet_name)
    dfs.append(df)

# Unificar as planilhas usando a função concat do pandas
combined_df = pd.concat(dfs)

# Salvar a planilha unificada em um arquivo Excel
output_file = input("Digite o nome do arquivo Excel de saída: ")
combined_df.to_excel(output_file, index=False)

print("Planilhas unificadas com sucesso!")
