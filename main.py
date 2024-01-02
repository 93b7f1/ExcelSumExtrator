import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog


def obter_segundo_valor_time(caminho_arquivo):
    try:
        if caminho_arquivo.lower().endswith('.csv'):
            df = pd.read_csv(caminho_arquivo, delimiter=',', quotechar='"', dtype={'Time': str})
        else:
            df = pd.read_excel(caminho_arquivo)
    except PermissionError as e:
        print(f"Ignorando arquivo '{caminho_arquivo}' devido a erro de permissão: {e}")
        return None
    except ValueError as e:
        print(f"Ignorando arquivo '{caminho_arquivo}' devido a um erro: {e}")
        return None

    df['Time'] = pd.to_datetime(df['Time'], dayfirst=True)

    if len(df['Time']) >= 2:
        segundo_valor_time = df['Time'].iloc[1]
        return segundo_valor_time.strftime('%d/%m/%Y %H:%M') if isinstance(segundo_valor_time, pd.Timestamp) else str(
            segundo_valor_time)
    else:
        return None
def obter_ultimo_valor_time(caminho_arquivo):
    try:
        if caminho_arquivo.lower().endswith('.csv'):
            df = pd.read_csv(caminho_arquivo, delimiter=',', quotechar='"', dtype={'Time': str})
        else:
            df = pd.read_excel(caminho_arquivo)
    except PermissionError as e:
        print(f"Ignorando arquivo '{caminho_arquivo}' devido a erro de permissão: {e}")
        return None
    except ValueError as e:
        print(f"Ignorando arquivo '{caminho_arquivo}' devido a um erro: {e}")
        return None

    if 'Time' in df.columns:
        df['Time'] = pd.to_datetime(df['Time'], dayfirst=True)
        ultimo_valor_time = df['Time'].iloc[-1]
        return ultimo_valor_time.strftime('%d/%m/%Y %H:%M') if isinstance(ultimo_valor_time, pd.Timestamp) else str(
            ultimo_valor_time)
    else:
        return None



def somar_colunas_excel(diretorio):
    arquivos = os.listdir(diretorio)
    arquivos_csv_xls_xlsx = [arquivo for arquivo in arquivos if arquivo.lower().endswith(('.csv', '.xls', '.xlsx')) and not arquivo.startswith('~$')]

    if not arquivos_csv_xls_xlsx:
        print("Nenhum arquivo CSV, Excel ou Excel antigo encontrado no diretório.")
        return None

    resultados_por_arquivo = {}

    for arquivo in arquivos_csv_xls_xlsx:
        caminho_arquivo = os.path.join(diretorio, arquivo)
        ultimo_valor_time = obter_ultimo_valor_time(caminho_arquivo)
        primeiro_valor = obter_segundo_valor_time(caminho_arquivo)

        try:
            if arquivo.lower().endswith('.csv'):
                df = pd.read_csv(caminho_arquivo, delimiter=',', quotechar='"', dtype={'Time': str})
            else:
                df = pd.read_excel(caminho_arquivo)
        except PermissionError as e:
            print(f"Ignorando arquivo '{arquivo}' devido a erro de permissão: {e}")
            continue

        if 'PRODUCE' in df.columns:
            df = df.drop(columns=['PRODUCE'])

        if 'Time' in df.columns:
            df = df.drop(columns=['Time'])

        df = df.apply(pd.to_numeric, errors='coerce').fillna(0)

        resultado = df.sum().to_dict()
        # resultado[' '] = ' '

        medias = df[df != 0].mean()
        for coluna, media in medias.items():
            if not pd.isna(media) and media != 0:
                resultado[f'Average{coluna}({arquivo})'] = media

        # resultado['  '] = '  '
        resultado['Start_Date'] = primeiro_valor
        resultado['End_Date'] = ultimo_valor_time

        resultado = {key: 0.0 if pd.isna(value) or value == '' else value for key, value in resultado.items()}

        resultados_por_arquivo[arquivo] = resultado

    return resultados_por_arquivo

def salvar_resultados_em_csv(resultados, nome_do_arquivo_saida):
    df_final = pd.DataFrame()
    for arquivo, resultado in resultados.items():
        df_titulo = pd.DataFrame([f' {arquivo}'])

        df_colunas = pd.DataFrame([resultado.keys()])
        df_resultado = pd.DataFrame([resultado.values()])

        df_final = pd.concat([df_final, df_titulo, df_colunas, df_resultado, pd.DataFrame([' '])], axis=0,
                             ignore_index=True)

    df_final = df_final.iloc[:-1]

    df_final = df_final.replace('', 0.0)

    df_final.to_csv(nome_do_arquivo_saida + '.csv', sep=';', index=False, header=False)



def selecionar_diretorio():
    diretorio_do_excel = filedialog.askdirectory()
    if diretorio_do_excel:
        entry_diretorio.delete(0, tk.END)
        entry_diretorio.insert(0, diretorio_do_excel)


def processar_excel():
    diretorio_do_excel = entry_diretorio.get()
    nome_do_arquivo_saida = entry_saida.get()

    resultados = somar_colunas_excel(diretorio_do_excel)

    if resultados is not None:
        salvar_resultados_em_csv(resultados, nome_do_arquivo_saida)
        resultado_label.config(text=f"O arquivo {nome_do_arquivo_saida}.csv foi salvo no mesmo diretório do programa")


root = tk.Tk()
root.title("Processar Arquivos Excel")

label_diretorio = tk.Label(root, text="Diretório dos Arquivos:")
entry_diretorio = tk.Entry(root, width=50)
button_selecionar_diretorio = tk.Button(root, text="Selecionar Diretório", command=selecionar_diretorio)

label_saida = tk.Label(root, text="Nome do Arquivo de Saída:")
entry_saida = tk.Entry(root, width=50)

button_processar = tk.Button(root, text="Processar", command=processar_excel)
resultado_label = tk.Label(root, text="")

label_diretorio.grid(row=0, column=0, padx=10, pady=10, sticky=tk.E)
entry_diretorio.grid(row=0, column=1, padx=10, pady=10, sticky=tk.W)
button_selecionar_diretorio.grid(row=0, column=2, padx=10, pady=10)

label_saida.grid(row=1, column=0, padx=10, pady=10, sticky=tk.E)
entry_saida.grid(row=1, column=1, padx=10, pady=10, sticky=tk.W)

button_processar.grid(row=2, column=1, pady=20)
resultado_label.grid(row=3, column=0, columnspan=3)

root.mainloop()
