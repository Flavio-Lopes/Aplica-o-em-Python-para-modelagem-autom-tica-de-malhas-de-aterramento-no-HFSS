import subprocess
import time
import os

# Solicitar ao usuário que insira o caminho para o arquivo Python
caminho_arquivo = input("\nInsira o caminho para o arquivo Python (ex: C:\\caminho\\para\\seu\\arquivo.py): \n").strip()

# Verificar se o arquivo Python existe
if not os.path.isfile(caminho_arquivo):
    print(f"O arquivo {caminho_arquivo} não foi encontrado. Por favor, verifique o caminho e tente novamente.")
    exit()

# Solicitar ao usuário que insira o caminho para o projeto do HFSS
caminho_projeto = input("\nInsira o caminho para o projeto do HFSS (ex: C:\\caminho\\para\\seu\\projeto.aedt): \n").strip()

# Verificar se o projeto do HFSS existe
if not os.path.isfile(caminho_projeto):
    print(f"O projeto {caminho_projeto} não foi encontrado. Por favor, verifique o caminho e tente novamente.")
    exit()


# Obter apenas o nome do arquivo do projeto (sem a extensão .aedt)
nome_projeto = os.path.splitext(os.path.basename(caminho_projeto))[0] 

# Solicitar ao usuário que insira o nome do designs do projeto do HFSS
nome_design = input("\nInsira o nome do design do projeto: \n").strip()

# Converter barras inclinadas para barras invertidas nos caminhos
caminho_arquivo = caminho_arquivo.replace("\\", "/")
caminho_projeto = caminho_projeto.replace("\\", "/")

# Solicitar ao usuário que insira o comprimento da haste
comprimento_haste = float(input("\nInsira o comprimento da haste: \n").strip())

# Calcular os pontos de início e fim
ponto_inicio = -comprimento_haste / 2
ponto_fim = comprimento_haste / 2

# Conteúdo do código HFSS a ser escrito no arquivo
codigo_hfss = f"""
import ScriptEnv

ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oDesktop.OpenProject("{caminho_projeto}")
oProject = oDesktop.SetActiveProject("{nome_projeto}")
oDesign = oProject.SetActiveDesign("{nome_design}")
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty( 
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers", 
                "haste_cobre:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="			, "0meter",
                    "Y:="			, "{ponto_inicio}meter",
                    "Z:="			, "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="			, "0meter",
                    "Y:="			, "{ponto_fim}meter",
                    "Z:="			, "0meter"
                ]
            ]
        ]
    ]
)
oProject.Save()
"""

# Escreve o código HFSS no arquivo Python
with open(caminho_arquivo, "w") as arquivo:
    arquivo.write(codigo_hfss)

#  Valida o endereço do ansysedt
def obter_caminho_valido(mensagem):
    while True:
        caminho = input(mensagem).strip()
        if caminho.lower() == 'sair':
            exit()
        elif os.path.isfile(caminho):
            return caminho
        else:
            print(f"\nO caminho {caminho} não foi encontrado. Tente novamente ou digite 'sair' para sair.")


# Caminho completo para o ansysedt.exe
caminho_ansysedt = r"C:\Program Files\AnsysEM\Ansys Student\v232\Win64\ansysedtsv.exe"

# Comando para executar o arquivo Python no HFSS
comando = f'"{caminho_ansysedt}" -runscript "{caminho_arquivo}"'

# Executa o comando
subprocess.Popen(comando, shell=True)
# Aguarde alguns segundos antes de fechar o HFSS
tempo_de_espera = 30  # ajuste conforme necessário
time.sleep(tempo_de_espera)
