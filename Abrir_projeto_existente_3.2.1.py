# Bibliotecas importantes para a aplicação
import subprocess
import time
import os
# Solicita ao usuário que insira o caminho para o arquivo Python
caminho_arquivo = input("\nInsira o caminho para o arquivo Python (ex: C:\\caminho\\para\\seu\\arquivo.py): ").strip()

# Solicita ao usuário que insira o caminho para o projeto do HFSS
caminho_projeto = input("\nInsira o caminho para o projeto do HFSS (ex: C:\\caminho\\para\\seu\\projeto.aedt): ").strip()

# Verificar se o projeto do HFSS existe
if not os.path.isfile(caminho_projeto):
    print(f"O projeto {caminho_projeto} não foi encontrado. Por favor, verifique o caminho e tente novamente.")
    exit()

# Converter barras inclinadas para barras invertidas nos caminhos devido a semântica do HFSS
caminho_projeto = caminho_projeto.replace("\\", "/")

# Conteúdo do código HFSS a ser escrito no arquivo .py, note que é mais simplificado comparado com a ação de criar uma aplicação e salvá-la em algum diretório
codigo_hfss = f"""
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oDesktop.OpenProject("{caminho_projeto}")
"""

# Escreve o código que será executado no HFSS no arquivo Python (.py)
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
# Solicita o caminho completo para o ansysedt.exe
caminho_ansysedt = obter_caminho_valido("\nInsira o caminho para o ansysedtf (ex: C:\\caminho\\para\\seu\\ansysedtsv.aedt):\n")
# Comando para executar o arquivo Python no HFSS
comando = f'"{caminho_ansysedt}" -runscript "{caminho_arquivo}"'

# Executa o comando em segundo plano, de maneira de para além de executar o arquivo, permite a abertura da interface gráfica do HFSS já com o projeto existente
subprocess.Popen(comando, shell=True)

# Tempo de atraso para executar todo o comando de abertura e carregamento do código no HFSS sem finalizar a compilação do código
tempo_de_espera = 30 # ajuste conforme necessário
time.sleep(tempo_de_espera)
