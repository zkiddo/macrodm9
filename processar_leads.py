import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

def selecionar_pasta():
    root = tk.Tk()
    root.withdraw()
    pasta = filedialog.askdirectory(title="Selecione a pasta com os arquivos CSV de resultados")
    return pasta

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()
    arquivo = filedialog.askopenfilename(title="Selecione a base CSV para mesclagem", filetypes=[("CSV Files", "*.csv")])
    return arquivo

def processar_arquivos(pasta_resultados, arquivo_base):
    # Listar todos os arquivos CSV na pasta
    arquivos = [f for f in os.listdir(pasta_resultados) if f.endswith('.csv')]
    
    # Criar uma lista para armazenar os DataFrames
    dfs = []

    # Ler e combinar todos os arquivos CSV da pasta
    for arquivo in arquivos:
        caminho_arquivo = os.path.join(pasta_resultados, arquivo)
        df = pd.read_csv(caminho_arquivo, dtype=str, encoding='latin1', sep=",", on_bad_lines="skip")
        dfs.append(df)

    # Se houver arquivos, combinar os DataFrames em um único DataFrame
    if dfs:
        df_resultados = pd.concat(dfs, ignore_index=True)
    else:
        print("Nenhum arquivo CSV encontrado na pasta!")
        return

    # Garantir que a coluna 'CNPJ' tenha exatamente 14 dígitos
    if 'CNPJ' in df_resultados.columns:
        df_resultados['CNPJ'] = df_resultados['CNPJ'].str.zfill(14)

    
    # Filtrar apenas leads não cadastrados
    if 'Status' in df_resultados.columns:
        df_resultados = df_resultados[df_resultados['Status'] == 'Lead Não Cadastrado']

    # Carregar a base CSV para mesclagem
    df_base = pd.read_csv(arquivo_base, dtype=str, encoding='latin1', sep=";", on_bad_lines="skip")
    
    # Garantir que a coluna 'CNPJ' na base tenha 14 dígitos
    if 'CNPJ' in df_base.columns:
        df_base['CNPJ'] = df_base['CNPJ'].str.zfill(14)

    cnpjs_resultados = set(df_resultados['CNPJ'])
    cnpjs_base = set(df_base['CNPJ'])
    interseccao = cnpjs_resultados.intersection(cnpjs_base)

    print(f"CNPJs em df_resultados: {len(cnpjs_resultados)}")
    print(f"CNPJs em df_base: {len(cnpjs_base)}")
    print(f"CNPJs em comum: {len(interseccao)}")

    
    # Mesclar dados usando CNPJ como chave primária
    df_final = df_resultados.merge(df_base, on='CNPJ', how='inner')


    # Verificar total antes da remoção
    print(f"Total de linhas antes da remoção de duplicados: {len(df_final)}")

# Verificar telefones duplicados antes
    duplicatas_antes = df_final[df_final.duplicated(subset=['TELEFONE'], keep=False)]
    print(f"Total de telefones duplicados antes da remoção: {len(duplicatas_antes)}")

# Listar os primeiros telefones duplicados
    print(duplicatas_antes.head(20))

# Normalizar espaços e caracteres estranhos antes da remoção
    df_final['TELEFONE'] = df_final['TELEFONE'].astype(str).str.strip()  # Remove espaços extras
    df_final['TELEFONE'] = df_final['TELEFONE'].str.replace(r'\D', '', regex=True)  # Remove traços, parênteses e espaços

# Remover duplicados mantendo apenas um CNPJ por telefone
    df_final = df_final.sort_values(by=['CNPJ'])
    df_final = df_final.drop_duplicates(subset=['TELEFONE'], keep='first')

# Verificar total depois da remoção
    print(f"Total de linhas depois da remoção de duplicados: {len(df_final)}")

# Verificar telefones duplicados depois
    duplicatas_depois = df_final[df_final.duplicated(subset=['TELEFONE'], keep=False)]
    print(f"Total de telefones duplicados depois da remoção: {len(duplicatas_depois)}")

# Se ainda houver duplicados, listar alguns para análise
    print(duplicatas_depois.head(20))



    # Salvar resultado final
    caminho_saida = os.path.join(pasta_resultados, 'Leads_Consolidados.xlsx')
    df_final.to_excel(r"C:\\Mailing\\resultado_final.xlsx", index=False)

    print(f"Arquivo salvo em: {caminho_saida}")

if __name__ == "__main__":
    pasta = selecionar_pasta()
    base_csv = selecionar_arquivo()
    processar_arquivos(pasta, base_csv)
