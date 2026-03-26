import pandas as pd
import os

def ler_planilha(caminho_arquivo):
    
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao == '.csv':
        print("Lendo arquivo CSV...")
        return pd.read_csv(caminho_arquivo)
    elif extensao in ['.xlsx', '.xls']:
        print("Lendo arquivo Excel...")
        return pd.read_excel(caminho_arquivo)
    else:
        raise ValueError(f"Formato de arquivo {extensao} não suportado.")
