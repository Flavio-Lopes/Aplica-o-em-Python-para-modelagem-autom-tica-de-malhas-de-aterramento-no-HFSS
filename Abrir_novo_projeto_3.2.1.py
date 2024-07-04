import subprocess
import time
import os
# Solicitar ao usuário que insira o caminho para o arquivo Python
caminho_arquivo = input("\nInsira o caminho para o arquivo Python (ex: C:\\caminho\\para\\seu\\nome_arquivo.py): \n").strip()

# Solicitar ao usuário que insira o caminho para o novo projeto do HFSS
caminho_projeto = input("\nInsira o caminho para o projeto do HFSS(ex: C:\\caminho\\para\\seu\\nome_projeto.aedt): \n").strip()
caminho_projeto_com_duplas = caminho_projeto.replace("\\","\\\\")

# Conteúdo do código HFSS a ser escrito no arquivo .py
codigo_hfss = f"""
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oProject = oDesktop.NewProject()
oProject.InsertDesign("HFSS", "HFSSDesign1", "HFSS Terminal Network", "")
oProject.SaveAs("{caminho_projeto_com_duplas}", True)
"""

# Escreve o código HFSS no arquivo Python
with open(caminho_arquivo, "w") as arquivo:
    arquivo.write(codigo_hfss)

#  Validar endereço do ansysedt
def obter_caminho_valido(mensagem):
    while True:
        caminho = input(mensagem).strip()
        if caminho.lower() == 'sair':
            exit()
        elif os.path.isfile(caminho):
            return caminho
        else:
            print(f"\nO caminho {caminho} não foi encontrado. Tente novamente ou digite 'sair' para sair.")

# Caminho completo para o ansysedt.exe ou ansysedtsv
caminho_ansysedt = obter_caminho_valido("\nInsira o caminho para o ansysedtf (ex: C:\\caminho\\para\\seu\\ansysedtsv.exe):\n")

# Comando para executar o arquivo Python no HFSS
comando = f'"{caminho_ansysedt}" -runscript "{caminho_arquivo}"'

# Executa o comando em segundo plano
subprocess.Popen(comando, shell=True)

# Aguarde alguns segundos antes de fechar o script Python
tempo_de_espera = 20  # ajuste conforme necessário
time.sleep(tempo_de_espera)
