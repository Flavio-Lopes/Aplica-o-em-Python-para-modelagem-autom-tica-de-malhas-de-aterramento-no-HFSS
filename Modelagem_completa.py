import subprocess
import time
import os
import pandas as pd
import matplotlib.pyplot as plt
import psutil
import numpy as np

def verificar_HFSS_em_execucao():
    # Verifica se o HFSS está em execução
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == 'ansysedtsv.exe':
            return True
    return False

# Solicitar ao usuário que insira o caminho para o arquivo Python
caminho_arquivo = input("\nInsira o caminho para o arquivo Python (ex: C:\\caminho\\para\\seu\\arquivo.py): \n").strip()

# Solicitar ao usuário que insira o caminho para o projeto base do HFSS
caminho_projeto_base = input("\nInsira o caminho para o projeto base do HFSS (ex: C:\\caminho\\para\\seu\\projeto.aedt): \n").strip()

# Solicitar ao usuário que insira o caminho para o projeto base do HFSS
caminho_projeto_novo = input("\nInsira o caminho para o projeto novo do HFSS baseado em suas alterações (ex: C:\\caminho\\para\\seu\\projeto_novo.aedt): \n").strip()
# Caminho completo para o ansysedt.exe

#  Função para validar o endereço dado do ansysedt
def obter_caminho_valido(mensagem):
    while True:
        caminho = input(mensagem).strip()
        if caminho.lower() == 'sair':
            exit()
        elif os.path.isfile(caminho):
            return caminho
        else:
            print(f"\nO caminho {caminho} não foi encontrado. Tente novamente ou digite 'sair' para sair.")

caminho_ansysedt = obter_caminho_valido("\nInsira o caminho para o ansysedtf (ex: C:\\caminho\\para\\seu\\ansysedtsv.exe):\n")

# Verificar se o projeto do HFSS existe
if not os.path.isfile(caminho_projeto_base):
    print(f"O projeto {caminho_projeto_base} não foi encontrado. Por favor, verifique o caminho e tente novamente.")
    exit()


# Obter apenas o nome do arquivo do projeto (sem a extensão .aedt)

nome_projeto = os.path.splitext(os.path.basename(caminho_projeto_base))[0] 

# Solicitar ao usuário que insira o nome do designs do projeto do HFSS
#nome_design = input("\nInsira o nome do design do projeto: \n").strip()

# Converter barras inclinadas para barras invertidas nos caminhos
caminho_arquivo_altern = caminho_arquivo.replace("\\", "/")
caminho_projeto_base = caminho_projeto_base.replace("\\", "/")
caminho_projeto_novo = caminho_projeto_novo.replace("\\", "/")


def modelo_1():
    # Solicitar ao usuário que insira o comprimento da haste 2-6m 
    comprimento_haste = float(input("\nInsira o comprimento da malha quadrada: \n").strip())

    # Solicitar ao usuário a espessura das hastes de cobre, entre 10 -15 mm
    diametro_hastes = float(input("\nInsira o diâmetro das hastes entre 10-15 mm: \n").strip())

    # Solicitar ao usuário a resistividade do solo 
    resistividade = float(input("\nInsira a resistividade do solo: \n").strip())

    if comprimento_haste <= 12:
        tamanho_reticulo = comprimento_haste / 3
        tamanho_malha_base = comprimento_haste 
        n_cinturoes = 0
        qtd_reticulos = 2
    else:
        # Calcule o divisor necessário
        divisor = 1 + 2 * ((comprimento_haste - 12) // 24)
        while True:
            tamanho_malha_base = comprimento_haste / divisor
            tamanho_reticulo = tamanho_malha_base / 3
            if 1 <= tamanho_reticulo <= 4:
                break
            divisor += 2
        n_cinturoes = int((divisor - 1) // 2)
        qtd_reticulos = 2 + 6 * n_cinturoes  # Cada cinturão adiciona 6 retículos em cada direção

    # Adicionar a mensagem de progresso
    print(f"Está sendo gerada uma malha {comprimento_haste}m x {comprimento_haste}m com {qtd_reticulos}x{qtd_reticulos} retículos quadrados de {tamanho_reticulo:.2f}m de comprimento cada.") 

    # Conteúdo do código HFSS a ser escrito no arquivo
    codigo_hfss = f"""
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oDesktop.OpenProject("{caminho_projeto_base}")
oProject = oDesktop.SetActiveProject("teste_malha_unit")
oProject.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:ProjectVariableTab",
            [
                "NAME:PropServers", 
                "ProjectVariables"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:$C",
                    "Value:="        , "{0.001*diametro_hastes}"
                ]
            ]
        ]
    ])
oProject.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:ProjectVariableTab",
            [
                "NAME:PropServers", 
                "ProjectVariables"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:$L",
                    "Value:="        , "{tamanho_malha_base/2}"
                ]
            ]
        ]
    ])
oProject.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:ProjectVariableTab",
            [
                "NAME:PropServers", 
                "ProjectVariables"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:$cond",
                    "Value:="        , "{1/resistividade}"
                ]
            ]
        ]
    ])
oProject.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:ProjectVariableTab",
            [
                "NAME:PropServers",
                "ProjectVariables"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:$R",
                    "Value:="        , "{2*comprimento_haste}"
                ]
            ]
        ]
    ])  
oDesign = oProject.SetActiveDesign("Malha_unitaria")
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.SetWCS(
    [
        "NAME:SetWCS Parameter",
        "Working Coordinate System:=", "Global",
        "RegionDepCSOk:="    , False
    ])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateGroup(
    [
        "NAME:GroupParameter",
        "ParentGroupID:="    , "Malha",
        "Parts:="        , "Haste_centra_Y,Haste_central_X,HasteX1,HasteX2,HasteY1,HasteY2",
        "SubmodelInstances:="    , "",
        "Groups:="        , ""
    ])
"""
    contador_grupos = 1
    # Adicionar a lógica de replicação para múltiplos cinturões
    for n in range(1, n_cinturoes + 1):
        for i in range(-n, n + 1):
            for j in range(-n, n + 1):
                if (i == n or i == -n or j == n or j == -n) and (i != 0 or j != 0):  # Evitar a posição central já considerada e dentro do cinturão n
                    codigo_hfss += f"""
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_Group"
    ])
oEditor.Paste()
oEditor.Move(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_Group{contador_grupos}",
        "NewPartsModelFlag:="    , "Model"
    ], 
    [
        "NAME:TranslateParameters",
        "TranslateVectorX:="    , "{i}*2*$L",
        "TranslateVectorY:="    , "{j}*2*$L",
        "TranslateVectorZ:="    , "0meter"
    ])
"""
                    contador_grupos += 1 
    codigo_hfss += f'oProject.SaveAs("{caminho_projeto_novo}", True)'

    print("Executando o modelo 1...")
    # Escreve o código HFSS no arquivo Python
    with open(caminho_arquivo, "w") as arquivo:
        arquivo.write(codigo_hfss)

    # Comando para executar o arquivo Python no HFSS
    comando = f'"{caminho_ansysedt}" -runscript "{caminho_arquivo}"'

    # Executa o comando
    processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # Animação das reticências
    anima_reticencias = ['.', '..', '...']
    idx = 0
    while True:
        # Se o processo ainda estiver em execução, exiba a animação
        print(f"\rCarregando{anima_reticencias[idx % len(anima_reticencias)]}", end="", flush=True)
        idx += 1
        time.sleep(0.5)  # Intervalo de atualização da animação
        # Verifica se o arquivo de saída foi criado ou modificado
        if os.path.exists(caminho_projeto_novo):
            # Verifica a última vez que o arquivo foi modificado
            ultima_modificacao = os.path.getmtime(caminho_projeto_novo)
            if time.time() - ultima_modificacao < 10:  # Arquivo modificado nos últimos 10 segundos
                print("\nO HFSS terminou de executar o modelo, confira no HFSS.")
                break
      
         
def modelo_2():
    # Solicitar ao usuário que insira o comprimento da haste 2-6m 
    comprimento_haste = float(input("\nInsira o comprimento da malha quadrada: \n").strip())

    # Solicitar ao usuário a espessura das hastes de cobre, entre 10-15 mm
    diametro_hastes = float(input("\nInsira o diâmetro das hastes entre 10-15 mm: \n").strip())

    # Solicitar ao usuário a dimensão das hastes verticais
    comprimento_haste_vertical = float(input("\nInsira o comprimento das hastes verticais: \n").strip())

    # Solicitar ao usuário a resistividade do solo 
    resistividade = float(input("\nInsira a resistividade do solo: \n").strip())

    if comprimento_haste <= 12:
        tamanho_reticulo = comprimento_haste / 3
        tamanho_malha_base = comprimento_haste / 2
        n_cinturoes = 0
        qtd_reticulos = 2
    else:
        # Calcule o divisor necessário
        divisor = 1 + 2 * ((comprimento_haste - 12) // 24)
        while True:
            tamanho_malha_base = comprimento_haste / divisor
            tamanho_reticulo = tamanho_malha_base / 3
            if 1 <= tamanho_reticulo <= 4:
                break
            divisor += 2
        n_cinturoes = int((divisor - 1) // 2)
        qtd_reticulos = 2 + 6 * n_cinturoes  # Cada cinturão adiciona 6 retículos em cada direção
    # Adicionar a mensagem de progresso
    print(f"Está sendo gerada uma malha {comprimento_haste}m x {comprimento_haste}m com {qtd_reticulos}x{qtd_reticulos} retículos quadrados de {tamanho_reticulo:.2f}m de comprimento cada.")

    # Conteúdo do código HFSS a ser escrito no arquivo
    codigo_hfss = f"""
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oDesktop.OpenProject("{caminho_projeto_base}")
oProject = oDesktop.SetActiveProject("teste_malha_unit")
oProject.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:ProjectVariableTab",
            [
                "NAME:PropServers",
                "ProjectVariables"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:$C",
                    "Value:="        , "{0.001 * diametro_hastes}"
                ]
            ]
        ]
    ])
oProject.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:ProjectVariableTab",
            [
                "NAME:PropServers",
                "ProjectVariables"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:$L",
                    "Value:="        , "{tamanho_malha_base / 2}"
                ]
            ]
        ]
    ])
oProject.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:ProjectVariableTab",
            [
                "NAME:PropServers",
                "ProjectVariables"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:$cond",
                    "Value:="        , "{1 / resistividade}"
                ]
            ]
        ]
    ])
oProject.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:ProjectVariableTab",
            [
                "NAME:PropServers",
                "ProjectVariables"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:$Lz",
                    "Value:="        , "{comprimento_haste_vertical}"
                ]
            ]
        ]
    ])

oProject.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:ProjectVariableTab",
            [
                "NAME:PropServers",
                "ProjectVariables"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:$R",
                    "Value:="        , "{2*comprimento_haste}"
                ]
            ]
        ]
    ])    
oDesign = oProject.SetActiveDesign("Malha_unitaria")
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "HasteX2"
    ])
oEditor.Paste()
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DAttributeTab",
            [
                "NAME:PropServers",
                "HasteX3"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Name",
                    "Value:="        , "Haste_vertical"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DCmdTab",
            [
                "NAME:PropServers",
                "Haste_vertical:CreatePolyline:1"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Coordinate System",
                    "Value:="        , "Global"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers",
                "Haste_vertical:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="            , "$L*(2/3)",
                    "Y:="            , "0",
                    "Z:="            , "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="            , "$L*(2/3)",
                    "Y:="            , "0",
                    "Z:="            , "-$Lz"
                ]
            ]
        ]
    ])
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_vertical"
    ])
oEditor.Paste()
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DCmdTab",
            [
                "NAME:PropServers",
                "Haste_vertical1:CreatePolyline:1"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Coordinate System",
                    "Value:="        , "Global"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers",
                "Haste_vertical1:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="            , "$L*(2/3)",
                    "Y:="            , "$L*(2/3)",
                    "Z:="            , "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="            , "$L*(2/3)",
                    "Y:="            , "$L*(2/3)",
                    "Z:="            , "-$Lz"
                ]
            ]
        ]
    ])
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_vertical"
    ])
oEditor.Paste()
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DCmdTab",
            [
                "NAME:PropServers",
                "Haste_vertical2:CreatePolyline:1"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Coordinate System",
                    "Value:="        , "Global"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers",
                "Haste_vertical2:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="            , "$L*(2/3)",
                    "Y:="            , "-$L*(2/3)",
                    "Z:="            , "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="            , "$L*(2/3)",
                    "Y:="            , "-$L*(2/3)",
                    "Z:="            , "-$Lz"
                ]
            ]
        ]
    ])
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_vertical"
    ])
oEditor.Paste()
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DCmdTab",
            [
                "NAME:PropServers",
                "Haste_vertical3:CreatePolyline:1"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Coordinate System",
                    "Value:="        , "Global"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers",
                "Haste_vertical3:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="            , "-$L*(2/3)",
                    "Y:="            , "$L*(2/3)",
                    "Z:="            , "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="            , "-$L*(2/3)",
                    "Y:="            , "$L*(2/3)",
                    "Z:="            , "-$Lz"
                ]
            ]
        ]
    ])
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_vertical"
    ])
oEditor.Paste()
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DCmdTab",
            [
                "NAME:PropServers",
                "Haste_vertical4:CreatePolyline:1"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Coordinate System",
                    "Value:="        , "Global"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers",
                "Haste_vertical4:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="            , "-$L*(2/3)",
                    "Y:="            , "0",
                    "Z:="            , "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="            , "-$L*(2/3)",
                    "Y:="            , "0",
                    "Z:="            , "-$Lz"
                ]
            ]
        ]
    ])
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_vertical"
    ])
oEditor.Paste()
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DCmdTab",
            [
                "NAME:PropServers",
                "Haste_vertical5:CreatePolyline:1"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Coordinate System",
                    "Value:="        , "Global"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers",
                "Haste_vertical5:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="            , "0",
                    "Y:="            , "$L*(2/3)",
                    "Z:="            , "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="            , "0",
                    "Y:="            , "$L*(2/3)",
                    "Z:="            , "-$Lz"
                ]
            ]
        ]
    ])
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_vertical"
    ])
oEditor.Paste()
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DCmdTab",
            [
                "NAME:PropServers",
                "Haste_vertical6:CreatePolyline:1"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Coordinate System",
                    "Value:="        , "Global"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers",
                "Haste_vertical6:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="            , "-$L*(2/3)",
                    "Y:="            , "0",
                    "Z:="            , "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="            , "-$L*(2/3)",
                    "Y:="            , "0",
                    "Z:="            , "-$Lz"
                ]
            ]
        ]
    ])
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_vertical"
    ])
oEditor.Paste()
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DCmdTab",
            [
                "NAME:PropServers",
                "Haste_vertical7:CreatePolyline:1"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Coordinate System",
                    "Value:="        , "Global"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers",
                "Haste_vertical7:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="            , "-$L*(2/3)",
                    "Y:="            , "-$L*(2/3)",
                    "Z:="            , "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="            , "-$L*(2/3)",
                    "Y:="            , "-$L*(2/3)",
                    "Z:="            , "-$Lz"
                ]
            ]
        ]
    ])
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_vertical"
    ])
oEditor.Paste()
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DCmdTab",
            [
                "NAME:PropServers",
                "Haste_vertical8:CreatePolyline:1"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Coordinate System",
                    "Value:="        , "Global"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers",
                "Haste_vertical8:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="            , "0",
                    "Y:="            , "-$L*(2/3)",
                    "Z:="            , "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="            , "0",
                    "Y:="            , "-$L*(2/3)",
                    "Z:="            , "-$Lz"
                ]
            ]
        ]
    ])
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_vertical"
    ])
oEditor.Paste()
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DCmdTab",
            [
                "NAME:PropServers",
                "Haste_vertical9:CreatePolyline:1"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Coordinate System",
                    "Value:="        , "Global"
                ]
            ]
        ]
    ])
oEditor.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:Geometry3DPolylineTab",
            [
                "NAME:PropServers",
                "Haste_vertical9:CreatePolyline:1:Segment0"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:Point1",
                    "X:="            , "0",
                    "Y:="            , "0",
                    "Z:="            , "0meter"
                ],
                [
                    "NAME:Point2",
                    "X:="            , "0",
                    "Y:="            , "0",
                    "Z:="            , "-$Lz"
                ]
            ]
        ]
    ])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateGroup(
    [
        "NAME:GroupParameter",
        "ParentGroupID:="    , "Malha",
        "Parts:="        , "Haste_centra_Y,Haste_central_X,Haste_vertical,Haste_vertical1,Haste_vertical2,Haste_vertical3,Haste_vertical4,Haste_vertical5,Haste_vertical6,Haste_vertical7,Haste_vertical8,Haste_vertical9,HasteX1,HasteX2,HasteY1,HasteY2",
        "SubmodelInstances:="    , "",
        "Groups:="        , ""
    ])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_centra_Y,Haste_central_X,Haste_vertical,Haste_vertical1,Haste_vertical2,Haste_vertical3,Haste_vertical4,Haste_vertical5,Haste_vertical6,Haste_vertical7,Haste_vertical8,HasteX1,HasteX2,HasteY1,HasteY2"
	])
oEditor.SetWCS(
    [
        "NAME:SetWCS Parameter",
        "Working Coordinate System:=", "Global",
        "RegionDepCSOk:="    , False
    ])
"""
    contador_grupos = 1
    # Adicionar a lógica de replicação para múltiplos cinturões
    for n in range(1, n_cinturoes + 1):
        for i in range(-n, n + 1):
            for j in range(-n, n + 1):
                if (i == n or i == -n or j == n or j == -n) and (i != 0 or j != 0):  # Evitar a posição central já considerada e dentro do cinturão n
                    codigo_hfss += f"""
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_Group"
    ])
oEditor.Paste()
oEditor.Move(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_Group{contador_grupos}",
        "NewPartsModelFlag:="    , "Model"
    ], 
    [
        "NAME:TranslateParameters",
        "TranslateVectorX:="    , "{i}*2*$L",
        "TranslateVectorY:="    , "{j}*2*$L",
        "TranslateVectorZ:="    , "0meter"
    ])
"""
                    contador_grupos += 1 
    codigo_hfss += f'oProject.SaveAs("{caminho_projeto_novo}", True)'

    print("Executando o modelo 2...")
    # Escreve o código HFSS no arquivo Python
    with open(caminho_arquivo, "w") as arquivo:
        arquivo.write(codigo_hfss)

    # Comando para executar o arquivo Python no HFSS
    comando = f'"{caminho_ansysedt}" -runscript "{caminho_arquivo}"'

    # Executa o comando
    processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # Animação das reticências
    anima_reticencias = ['.', '..', '...']
    idx = 0
    while True:
        # Se o processo ainda estiver em execução, exiba a animação
        print(f"\rCarregando{anima_reticencias[idx % len(anima_reticencias)]}", end="", flush=True)
        idx += 1
        time.sleep(0.5)  # Intervalo de atualização da animação
        # Verifica se o arquivo de saída foi criado ou modificado
        if os.path.exists(caminho_projeto_novo):
            # Verifica a última vez que o arquivo foi modificado
            ultima_modificacao = os.path.getmtime(caminho_projeto_novo)
            if time.time() - ultima_modificacao < 10:  # Arquivo modificado nos últimos 10 segundos
                print("\nO HFSS terminou de executar o modelo, confira no HFSS.")
                break
      
def modelo_3():
       # Solicitar ao usuário que insira o comprimento da haste 2-6m 
    comprimento_haste = float(input("\nInsira o comprimento da malha quadrada: \n").strip())

   # Solicitar ao usuário a espessura das hastes de cobre, entre 10 -15 mm
    diametro_hastes = float(input("\nInsira o diâmetro das hastes entre 10-15 mm: \n").strip())
    
	#Solicitar ao usuário a dimensão das hastes verticais
    comprimento_haste_vertical = float(input("\nInsira o comprimento das hastes verticais inclinadas: \n").strip())
     
	 # Solicitar ao usuário a resistividade do solo 
    resistividade = float(input("\nInsira a resistividade do solo: \n").strip()) 
    if comprimento_haste <= 12:
        tamanho_reticulo = comprimento_haste / 3
        tamanho_malha_base = comprimento_haste / 2
        n_cinturoes = 0
        qtd_reticulos = 2
    else:
        # Calcule o divisor necessário
        divisor = 1 + 2 * ((comprimento_haste - 12) // 24)
        while True:
            tamanho_malha_base = comprimento_haste / divisor
            tamanho_reticulo = tamanho_malha_base / 3
            if 1 <= tamanho_reticulo <= 4:
                break
            divisor += 2
        n_cinturoes = int((divisor - 1) // 2)
        qtd_reticulos = 2 + 6 * n_cinturoes  # Cada cinturão adiciona 6 retículos em cada direção
    # Adicionar a mensagem de progresso
    print(f"Está sendo gerada uma malha {comprimento_haste}m x {comprimento_haste}m com {qtd_reticulos}x{qtd_reticulos} retículos quadrados de {tamanho_reticulo:.2f}m de comprimento cada.")

   # Conteúdo do código HFSS a ser escrito no arquivo
    codigo_hfss = f"""
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oDesktop.OpenProject("{caminho_projeto_base}")
oProject = oDesktop.SetActiveProject("teste_malha_unit")
oProject.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:ProjectVariableTab",
			[
				"NAME:PropServers", 
				"ProjectVariables"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:$C",
					"Value:="		, "{0.001*diametro_hastes}"
				]
			]
		]
	])
oProject.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:ProjectVariableTab",
			[
				"NAME:PropServers", 
				"ProjectVariables"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:$L",
					"Value:="		, "{tamanho_malha_base/2}"
				]
			]
		]
	])
oProject.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:ProjectVariableTab",
			[
				"NAME:PropServers", 
				"ProjectVariables"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:$cond",
					"Value:="		, "{1/resistividade}"
				]
			]
		]
	])
oProject.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:ProjectVariableTab",
			[
				"NAME:PropServers", 
				"ProjectVariables"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:$Lz",
					"Value:="		, "{comprimento_haste_vertical}"
				]
			]
		]
	])
oProject.ChangeProperty(
    [
        "NAME:AllTabs",
        [
            "NAME:ProjectVariableTab",
            [
                "NAME:PropServers",
                "ProjectVariables"
            ],
            [
                "NAME:ChangedProps",
                [
                    "NAME:$R",
                    "Value:="        , "{2*comprimento_haste}"
                ]
            ]
        ]
    ]) 
oDesign = oProject.SetActiveDesign("Malha_unitaria")
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_centra_Y"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DAttributeTab",
			[
				"NAME:PropServers", 
				"Haste_centra_Y1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Name",
					"Value:="		, "Haste_diagon1"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon1:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon1:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "0",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon1:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "0",
					"Y:="			, "-$Lz*(((2)^(1/2))/2)-$L*(2/3)",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon2:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon2:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "$L*(2/3)",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon2:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "$L*(2/3)",
					"Y:="			, "-$Lz*(((2)^(1/2))/2)-$L*(2/3)",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon3:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon3:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon3:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "-$Lz*(((2)^(1/2))/2)-$L*(2/3)",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon4:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon4:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "$L*(2/3)",
					"Y:="			, "0",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon4:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "$Lz*(((2)^(1/2))/2)+$L*(2/3)",
					"Y:="			, "0",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon5:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon5:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "0",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon5:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "-$Lz*(((2)^(1/2))/2)-$L*(2/3)",
					"Y:="			, "0",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon6:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon6:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "0",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon6:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "0",
					"Y:="			, "$Lz*(((2)^(1/2))/2)+$L*(2/3)",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon7:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon7:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "$L*(2/3)",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon7:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "$L*(2/3)",
					"Y:="			, "$Lz*(((2)^(1/2))/2)+$L*(2/3)",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon8:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon8:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon8:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "$Lz*(((2)^(1/2))/2)+$L*(2/3)",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon9:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon9:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "$L*(2/3)",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon9:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "$Lz*(((2)^(1/2))/2)+$L*(2/3)",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon10:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon10:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "$L*(2/3)",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon10:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "$Lz*(((2)^(1/2))/2)+$L*(2/3)",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon11:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon11:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon11:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "-$Lz*(((2)^(1/2))/2)-$L*(2/3)",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon12:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon12:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "0meter"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon12:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "-$Lz*(((2)^(1/2))/2)-$L*(2/3)",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "-$Lz*(((2)^(1/2))/2)"
				]
			]
		]
	])

oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_diagon1"
	])
oEditor.Paste()
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DCmdTab",
			[
				"NAME:PropServers", 
				"Haste_diagon13:CreatePolyline:1"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Coordinate System",
					"Value:="		, "Global"
				]
			]
		]
	])

oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon13:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "0",
					"Y:="			, "0",
					"Z:="			, "0"
				]
			]
		]
	])
oEditor.ChangeProperty(
	[
		"NAME:AllTabs",
		[
			"NAME:Geometry3DPolylineTab",
			[
				"NAME:PropServers", 
				"Haste_diagon13:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point2",
					"X:="			, "0",
					"Y:="			, "0",
					"Z:="			, "-$Lz"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.SetWCS(
	[
		"NAME:SetWCS Parameter",
		"Working Coordinate System:=", "Global",
		"RegionDepCSOk:="	, False
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateGroup(
	[
		"NAME:GroupParameter",
		"ParentGroupID:="	, "Malha",
		"Parts:="		, "Haste_centra_Y,Haste_central_X,Haste_diagon1,Haste_diagon2,Haste_diagon3,Haste_diagon4,Haste_diagon5,Haste_diagon6,Haste_diagon7,Haste_diagon8,Haste_diagon9,Haste_diagon10,Haste_diagon11,Haste_diagon12,Haste_diagon13,HasteX1,HasteX2,HasteY1,HasteY2",
		"SubmodelInstances:="	, "",
		"Groups:="		, ""
	])
"""
    contador_grupos = 1
    # Adicionar a lógica de replicação para múltiplos cinturões
    for n in range(1, n_cinturoes + 1):
        for i in range(-n, n + 1):
            for j in range(-n, n + 1):
                if (i == n or i == -n or j == n or j == -n) and (i != 0 or j != 0):  # Evitar a posição central já considerada e dentro do cinturão n
                    codigo_hfss += f"""
oEditor.Copy(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_Group"
    ])
oEditor.Paste()
oEditor.Move(
    [
        "NAME:Selections",
        "Selections:="        , "Haste_Group{contador_grupos}",
        "NewPartsModelFlag:="    , "Model"
    ], 
    [
        "NAME:TranslateParameters",
        "TranslateVectorX:="    , "{i}*2*$L",
        "TranslateVectorY:="    , "{j}*2*$L",
        "TranslateVectorZ:="    , "0meter"
    ])
"""
                    contador_grupos += 1 
    codigo_hfss += f'oProject.SaveAs("{caminho_projeto_novo}", True)'
    
    print("Executando o modelo 3...")
    # Escreve o código HFSS no arquivo Python
    with open(caminho_arquivo, "w") as arquivo:
        arquivo.write(codigo_hfss)

    # Comando para executar o arquivo Python no HFSS
    comando = f'"{caminho_ansysedt}" -runscript "{caminho_arquivo}"'

    # Executa o comando
    processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

# Animação das reticências
    anima_reticencias = ['.', '..', '...']
    idx = 0
    while True:
        # Se o processo ainda estiver em execução, exiba a animação
        print(f"\rCarregando{anima_reticencias[idx % len(anima_reticencias)]}", end="", flush=True)
        idx += 1
        time.sleep(0.5)  # Intervalo de atualização da animação
    # Verifica se o arquivo de saída foi criado ou modificado
        if os.path.exists(caminho_projeto_novo):
        # Verifica a última vez que o arquivo foi modificado
            ultima_modificacao = os.path.getmtime(caminho_projeto_novo)
            if time.time() - ultima_modificacao < 10:  # Arquivo modificado nos últimos 10 segundos
                print("\nO HFSS terminou de executar o modelo, confira no HFSS.\n")
                break

def lightning(tpp, tmm, tend, peak, resol, diretorio):
    # Pontos de tempo (em ms)
    t = np.linspace(0, tend, resol)
    
    # Variáveis auxiliares
    tp = tpp
    tm = tmm
    
    # Cálculo dos parâmetros
    alpha = 1 / tp
    beta = 1 / tm
    A = 1 / (np.exp(-alpha * (np.log(beta) - np.log(alpha)) / (beta - alpha)) - 
             np.exp(-beta * (np.log(beta) - np.log(alpha)) / (beta - alpha)))
    I = A * peak * (np.exp(-alpha * t) - np.exp(-beta * t))
    
    # Cálculo do erro
    max_idx = np.argmax(I)
    error = np.sum(np.abs(I[max_idx:-1] - I[max_idx + 1:])) / (len(I) - max_idx)
    
    # Encontrando os pontos
    max_i = np.max(I)
    max_v = np.argmax(I)
    max_t = t[max_v]
    
    med_v_candidates = np.where(np.abs(I - peak / 2) < error)[0]
    if len(med_v_candidates) == 0:
        med_v = max_v  # Usando o valor máximo como fallback
    else:
        med_v = med_v_candidates[-1]
    
    med_t = t[med_v]
    kk = 0
    jj = 0
    
    while (np.abs(max_t - tpp) > resol / len(t) or np.abs(med_t - tmm) > resol / len(t)) and kk < np.math.comb(resol, 2) and jj != 10000:
        # Método da bisseção
        if np.abs(max_t - tpp) > resol / len(t):
            if max_t - tpp > 0:
                tp *= 0.9995
            else:
                tp *= 1.0005
        if np.abs(med_t - tmm) > resol / len(t):
            if med_t - tmm > 0:
                tm *= 0.9995
            else:
                tm *= 1.0005
        
        # Cálculo dos parâmetros
        alpha = 1 / tp
        beta = 1 / tm
        A = 1 / (np.exp(-alpha * (np.log(beta) - np.log(alpha)) / (beta - alpha)) - 
                 np.exp(-beta * (np.log(beta) - np.log(alpha)) / (beta - alpha)))
        I = A * peak * (np.exp(-alpha * t) - np.exp(-beta * t))
        
        # Cálculo do erro
        max_idx = np.argmax(I)
        error = np.sum(np.abs(I[max_idx:-1] - I[max_idx + 1:])) / (len(I) - max_idx)
        
        # Encontrando os pontos
        max_i = np.max(I)
        max_v = np.argmax(I)
        max_t = t[max_v]
        
        med_v_candidates = np.where(np.abs(I - peak / 2) < error)[0]
        if len(med_v_candidates) == 0:
            med_v = max_v  # Usando o valor máximo como fallback
        else:
            med_v = med_v_candidates[-1]
        
        # Finalizando o loop para repetições de med_t e max_t
        if t[med_v] == med_t and t[max_v] == max_t and 0.8 * tmm < t[med_v] < 1.2 * tmm and 0.8 * tpp < t[max_v] < 1.2 * tpp:
            jj = 10000
        
        kk += 1
        med_t = t[med_v]
    
    # Fator de correção do pico
    I = (peak / np.max(I)) * I
    e = np.array([np.abs(tpp - max_t) / tpp, np.abs(tmm - med_t) / tmm])
    emax = np.max(e)
    
    K = A * (peak / np.max(I))

    # Plotando o gráfico
    plt.figure(figsize=(10, 6))
    plt.plot(t, I, label='Corrente de Descarga')
    plt.xlabel('Tempo (microseconds)')
    plt.ylabel('Corrente (kA)')
    plt.title('Modelo de Dupla Exponencial de Descarga Atmosférica')
    plt.legend()
    plt.grid(True)
    plt.show()
    
    # Formatando o nome do arquivo
    tpp_int = int(tpp)  # Pegando apenas a parte inteira de tpp
    file_name = f'k{int(peak)}_{tpp_int}_{int(tmm)}u.tab'
    file_path = os.path.join(diretorio, file_name)
    np.savetxt(file_path, np.column_stack((t*1e-6, I)), delimiter='\t', header='Time (µs)\tCurrent (kA)', comments='')

    return K, alpha, beta, emax, kk, file_path, file_name
def define_input1():
    diretorio_raio=os.path.dirname(caminho_projeto_novo)
    nome_projeto_novo = os.path.splitext(os.path.basename(caminho_projeto_novo))[0] 
    caminho_arquivo_2 = os.path.join(diretorio_raio, "raio3.py") 
    peak = float(input("\nInsira o valor de pico da corrente (em kA): ").strip())
    codigo_hfss_export_descarga = f"""
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oDesktop.OpenProject("{caminho_projeto_novo}")
oProject = oDesktop.SetActiveProject("{nome_projeto_novo}")
oDesign = oProject.SetActiveDesign("Malha_unitaria")
oDesign.ExportDataset("Lightning8_1", "{diretorio_raio}/Raio.8_25.tab")
oModule = oDesign.GetModule("BoundarySetup")
oModule.EditCurrent("Current1", 
	[
		"NAME:Current1",
		[
			"NAME:Direction",
			"Coordinate System:="	, "plano do solo",
			"Start:="		, ["0meter","0meter","6.42meter"],
			"End:="			, ["0meter","0meter","6.4meter"]
		],
		"TimeProfile:="		, "Dataset",
		"HfssFrequency:="	, "1GHz",
		"MinFreq:="		, "0Hz",
		"MaxFreq:="		, "400kHz",
		"Magnitude:="		, "{peak}A",
		"Dataset:="		, "Lightning8_1",
		"DatasetMaxFreq:="	, "400kHz"
	])
oModule = oDesign.GetModule("AnalysisSetup")
oModule.EditSetup("Setup1", 
	[
		"NAME:Setup1",
		"Frequency:="		, "5GHz",
		"MaxDeltaE:="		, 0.1,
		"MaximumPasses:="	, 20,
		"UseImplicitSolver:="	, True,
		"IsEnabled:="		, True,
		[
			"NAME:MeshLink",
			"ImportMesh:="		, False
		],
		"BasisOrder:="		, -1,
		[
			"NAME:Transient",
			"ListsForFields:="	, ["solo"],
			"UseSaveCount:="	, 0,
			"UseSaveStart:="	, 1,
			"SaveStart:="		, "0s",
			"UseSaveDelta:="	, 1,
			"SaveDelta:="		, "500ns",
			"SaveRadFields:="	, 0,
			"SaveFDRadFields:="	, 0,
			"UseAutoTermination:="	, 0,
			"TerminateOnMaximum:="	, 1,
			"UseMaxTime:="		, 1,
			"MaxTime:="		, "50us"
		]
	])
oProject.Save()
"""	
        # Escreve o código HFSS no arquivo Python
    with open(caminho_arquivo_2, "w") as arquivo:
        arquivo.write(codigo_hfss_export_descarga)

	# Comando para executar o arquivo Python no HFSS
    comando = f'"{caminho_ansysedt}" -runscript "{caminho_arquivo_2}"'
    processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # Animação das reticências
    anima_reticencias = ['.', '..', '...']
    idx = 0
    while True:
        # Se o processo ainda estiver em execução, exiba a animação
        print(f"\rCarregando{anima_reticencias[idx % len(anima_reticencias)]}", end="", flush=True)
        idx += 1
        time.sleep(0.5)  # Intervalo de atualização da animação
	# Verifica se o arquivo de saída foi criado ou modificado
        if os.path.exists(caminho_projeto_novo):
		# Verifica a última vez que o arquivo foi modificado
            ultima_modificacao = os.path.getmtime(caminho_projeto_novo)
            if time.time() - ultima_modificacao < 10:  # Arquivo modificado nos últimos 10 segundos
                print("\nSegue o gráfico da descarga atmosférica.\n")
                break

    caminho_raio = os.path.join(diretorio_raio, "Raio.8_25.tab")
	# Ler o arquivo .tab

    df = pd.read_csv(caminho_raio, delimiter='\t')
	# Remover as aspas duplas dos nomes das colunas
    print(df.columns)
	# Criar um gráfico XY
    plt.figure(figsize=(10,6))
    plt.plot(df['X ']*1000000, df['Y']*peak, marker='o')  # Agora você pode usar 'X' e 'Y' diretamente
    plt.title('Descarga 8/25')
    plt.xlabel('Tempo (us)')
    plt.ylabel('Corrente (A)')
    plt.grid(True)
    plt.show() 

def define_input2():
    diretorio_raio=os.path.dirname(caminho_projeto_novo)
    nome_projeto_novo = os.path.splitext(os.path.basename(caminho_projeto_novo))[0] 
    caminho_arquivo_2 = os.path.join(diretorio_raio, "raio4.py")
    peak = float(input("\nInsira o valor de pico da corrente (em kA): ").strip())
    codigo_hfss_export_descarga = f"""
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oDesktop.OpenProject("{caminho_projeto_novo}")
oProject = oDesktop.SetActiveProject("{nome_projeto_novo}")
oDesign = oProject.SetActiveDesign("Malha_unitaria")
oDesign.ExportDataset("Lightning1_2_1", "{diretorio_raio}/Raio.1.2_50.tab")
oModule = oDesign.GetModule("BoundarySetup")
oModule.EditCurrent("Current1", 
	[
		"NAME:Current1",
		[
			"NAME:Direction",
			"Coordinate System:="	, "plano do solo",
			"Start:="		, ["0meter","0meter","6.42meter"],
			"End:="			, ["0meter","0meter","6.4meter"]
		],
		"TimeProfile:="		, "Dataset",
		"HfssFrequency:="	, "1GHz",
		"MinFreq:="		, "0Hz",
		"MaxFreq:="		, "400kHz",
		"Magnitude:="		, "{peak}A",
		"Dataset:="		, "Lightning1_2_1",
		"DatasetMaxFreq:="	, "400kHz"
	])
oModule = oDesign.GetModule("AnalysisSetup")
oModule.EditSetup("Setup1", 
	[
		"NAME:Setup1",
		"Frequency:="		, "5GHz",
		"MaxDeltaE:="		, 0.1,
		"MaximumPasses:="	, 20,
		"UseImplicitSolver:="	, True,
		"IsEnabled:="		, True,
		[
			"NAME:MeshLink",
			"ImportMesh:="		, False
		],
		"BasisOrder:="		, -1,
		[
			"NAME:Transient",
			"ListsForFields:="	, ["solo"],
			"UseSaveCount:="	, 0,
			"UseSaveStart:="	, 1,
			"SaveStart:="		, "0s",
			"UseSaveDelta:="	, 1,
			"SaveDelta:="		, "500ns",
			"SaveRadFields:="	, 0,
			"SaveFDRadFields:="	, 0,
			"UseAutoTermination:="	, 0,
			"TerminateOnMaximum:="	, 1,
			"UseMaxTime:="		, 1,
			"MaxTime:="		, "100us"
		]
	])
oProject.Save()
"""	
		# Escreve o código HFSS no arquivo Python
    with open(caminho_arquivo_2, "w") as arquivo:
         arquivo.write(codigo_hfss_export_descarga)

	# Comando para executar o arquivo Python no HFSS
    comando = f'"{caminho_ansysedt}" -runscript "{caminho_arquivo_2}"'
    processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

	# Animação das reticências
    anima_reticencias = ['.', '..', '...']
    idx = 0
    while True:
		# Se o processo ainda estiver em execução, exiba a animação
        print(f"\rCarregando{anima_reticencias[idx % len(anima_reticencias)]}", end="", flush=True)
        idx += 1
        time.sleep(0.5)  # Intervalo de atualização da animação
	# Verifica se o arquivo de saída foi criado ou modificado
        if os.path.exists(caminho_projeto_novo):
		# Verifica a última vez que o arquivo foi modificado
            ultima_modificacao = os.path.getmtime(caminho_projeto_novo)
            if time.time() - ultima_modificacao < 10:  # Arquivo modificado nos últimos 10 segundos
                print("\nSegue o gráfico da descarga atmosférica.\n")
                break

    caminho_raio = os.path.join(diretorio_raio, "Raio.1.2_50.tab")
	# Ler o arquivo .tab

    df = pd.read_csv(caminho_raio, delimiter='\t')
    # Remover as aspas duplas dos nomes das colunas
    print(df.columns)
    # Criar um gráfico XY
    plt.figure(figsize=(10,6))
    plt.plot(df['X ']*1000000, df['Y']*peak, marker='o')  # Agora você pode usar 'X' e 'Y' diretamente
    plt.title('Descarga 1.25/50')
    plt.xlabel('Tempo (us)')
    plt.ylabel('Corrente (A)')
    plt.grid(True)
    plt.show() 

def define_input3():
    diretorio_raio = os.path.dirname(caminho_projeto_novo)
    while True:
        tpp = float(input("\nInsira o tempo de pico da corrente (em microsegundos): ").strip())
        tmm = float(input("\nInsira o tempo de meia-onda da corrente (em microsegundos): ").strip())
        tend = float(input("\nInsira o tempo final da simulação (em microsegundos): ").strip())
        peak = float(input("\nInsira o valor de pico da corrente (em kA): ").strip())
        resol = int(input("\nInsira a resolução (número de pontos): ").strip())
        
        K, alpha, beta, emax, kk, file_path, file_name = lightning(tpp, tmm, tend, peak, resol, diretorio_raio)
        file_path= file_path.replace("\\", "/")
        #print(f"K: {K}, alpha: {alpha}, beta: {beta}, emax: {emax}, kk: {kk}")
        print("Confira o gráfico gerado.")
    
        while True:
            salvar = input("\nDeseja salvar este modelo? (s/n): ").strip().lower()
            if salvar == 's':
                nome_projeto_novo = os.path.splitext(os.path.basename(caminho_projeto_novo))[0]
                caminho_arquivo_2 = os.path.join(diretorio_raio, f"raio_criado.py")
                codigo_hfss_export_descarga = f"""
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oDesktop.OpenProject("{caminho_projeto_novo}")
oProject = oDesktop.SetActiveProject("{nome_projeto_novo}")
oDesign = oProject.SetActiveDesign("Malha_unitaria")
oDesign.ImportDataset("{file_path}")
oModule = oDesign.GetModule("BoundarySetup")
oModule.EditCurrent("Current1", 
    [
        "NAME:Current1",
        [
            "NAME:Direction",
            "Coordinate System:="    , "Global",
            "Start:="        , ["0meter","0meter","7.02meter"],
            "End:="            , ["0meter","0meter","7meter"]
        ],
        "TimeProfile:="        , "Dataset",
        "HfssFrequency:="    , "1GHz",
        "MinFreq:="        , "0Hz",
        "MaxFreq:="        , "200MHz",
        "Magnitude:="        , "1A",
        "Dataset:="        , "{os.path.splitext(file_name)[0]}",
        "DatasetMaxFreq:="    , "200MHz"
    ])
oModule = oDesign.GetModule("AnalysisSetup")
oModule.EditSetup("Setup1", 
	[
		"NAME:Setup1",
		"Frequency:="		, "5GHz",
		"MaxDeltaE:="		, 0.1,
		"MaximumPasses:="	, 20,
		"UseImplicitSolver:="	, True,
		"IsEnabled:="		, True,
		[
			"NAME:MeshLink",
			"ImportMesh:="		, False
		],
		"BasisOrder:="		, -1,
		[
			"NAME:Transient",
			"ListsForFields:="	, ["solo"],
			"UseSaveCount:="	, 0,
			"UseSaveStart:="	, 1,
			"SaveStart:="		, "0s",
			"UseSaveDelta:="	, 1,
			"SaveDelta:="		, "500ns",
			"SaveRadFields:="	, 0,
			"SaveFDRadFields:="	, 0,
			"UseAutoTermination:="	, 0,
			"TerminateOnMaximum:="	, 1,
			"UseMaxTime:="		, 1,
			"MaxTime:="		, "{tend}us"
		]
	])
oProject.Save()
"""    
                # Escreve o código HFSS no arquivo Python
                with open(caminho_arquivo_2, "w") as arquivo:
                    arquivo.write(codigo_hfss_export_descarga)

                # Comando para executar o arquivo Python no HFSS
                comando = f'"{caminho_ansysedt}" -runscript "{caminho_arquivo_2}"'
                processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

                # Animação das reticências
                anima_reticencias = ['.', '..', '...']
                idx = 0
                while True:
                    # Se o processo ainda estiver em execução, exiba a animação
                    print(f"\rCarregando{anima_reticencias[idx % len(anima_reticencias)]}", end="", flush=True)
                    idx += 1
                    time.sleep(0.5)  # Intervalo de atualização da animação
                    # Verifica se o arquivo de saída foi criado ou modificado
                    if os.path.exists(caminho_projeto_novo):
                        # Verifica a última vez que o arquivo foi modificado
                        ultima_modificacao = os.path.getmtime(caminho_projeto_novo)
                        if time.time() - ultima_modificacao < 10:  # Arquivo modificado nos últimos 10 segundos
                            print("\nO HFSS terminou de executar o modelo, confira no HFSS.\n")
                            break
                return  # Retorna à função principal
            elif salvar == 'n':
                break
            else:
                print("Entrada inválida. Por favor, insira 's' para salvar ou 'n' para recusar.")

def simular_projeto(caminho_projeto_novo):
    # Código para gerar o arquivo simular.py e executar a simulação no HFSS
    diretorio = os.path.dirname(caminho_projeto_novo)
    nome_projeto_novo = os.path.splitext(os.path.basename(caminho_projeto_novo))[0]
    caminho_arquivo_simular = os.path.join(diretorio, "simular.py")
    caminho_log_simulacao = os.path.join(diretorio, "log_simulacao.txt")
    
    # Limpa o conteúdo do arquivo de log
    with open(caminho_log_simulacao, "w") as log_file:
        log_file.write("")
    
    codigo_hfss_simulacao = f"""
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oDesktop.OpenProject("{caminho_projeto_novo}")
oProject = oDesktop.SetActiveProject("{nome_projeto_novo}")
oDesign = oProject.SetActiveDesign("Malha_unitaria")
with open("{caminho_log_simulacao}", "a") as log_file:
    log_file.write("Simulacao Iniciada\\n")
oDesign.AnalyzeAll()
"""
    
    with open(caminho_arquivo_simular, "w") as arquivo:
        arquivo.write(codigo_hfss_simulacao)

    comando = f'"{caminho_ansysedt}" -runscript "{caminho_arquivo_simular}"'
    processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # Animação das reticências
    anima_reticencias = ['.', '..', '...']
    idx = 0
    simulacao_iniciada = False
    print("Carregando a simulação...")
    while not simulacao_iniciada:
        # Verifica o arquivo de log para detectar quando a simulação inicia
        with open(caminho_log_simulacao, "r") as log_file:
            logs = log_file.readlines()
            for log in logs:
                if "Simulacao Iniciada" in log:
                    simulacao_iniciada = True
                    break
        if simulacao_iniciada:
            print("\nA simulação iniciou, confira no HFSS.\n")
            break
        # Exiba a animação enquanto aguarda
        print(f"\rCarregando{anima_reticencias[idx % len(anima_reticencias)]}", end="", flush=True)
        idx += 1
        time.sleep(0.5)  # Intervalo de atualização da animação

def escolhe_modelo():
    while True:
        # Solicitar o modelo da malha 
        print("\nInsira o modelo desejado (1, 2 ou 3) ou digite 'sair' para sair da aplicação.\n Modelo 1: Malha horizontal.\n Modelo 2: Malha com hastes verticais.\n Modelo 3: Malha com hastes verticais inclinadas.")
        modelo_malha = input().strip()

        if modelo_malha == '1':
            modelo_1()
        elif modelo_malha == '2':
            modelo_2()
        elif modelo_malha == '3':
            modelo_3()
        elif modelo_malha.lower() == 'sair':
            print("Saindo da aplicação...")
            break
        else:
            print("Entrada inválida. Por favor, insira 1, 2, 3 ou 'sair'.")
            continue

        # O usuário irá escolher se deseja alterar o modelo, ou continuar a aplicação
        while True:
            print("\nSe deseja continuar a aplicação, digite: 1\nSe deseja alterar o modelo, digite: 2 ")
            opcao = input().strip()

            if opcao == '1':
                while True:
                    print("\nPara o input da aplicação, digite 1 para uma descarga atmosférica 8/20us, \ndigite 2 para uma descarga 1.2/50us ou \ndigite 3 para gerar seu próprio modelo.")
                    modelo_input = input().strip()
                    if modelo_input == '1':
                        define_input1()
                    elif modelo_input == '2':
                        define_input2()
                    elif modelo_input == '3':
                        define_input3()
                    else:
                        print("Entrada inválida. Por favor, insira 1, 2 ou 3")
                        continue

                    while True:
                        print("\nSe deseja continuar com esse modelo de descarga, digite 1. Se não, digite 2)")
                        continuar = input().strip()
                        if continuar == '1':
                            while verificar_HFSS_em_execucao():
                                print("O HFSS está em execução. Por favor, feche o programa para prosseguir.")
                                time.sleep(5)
                            break  # Sai do loop "continuar" para prosseguir
                        elif continuar == '2':
                            break  # Sai do loop "continuar" para voltar ao loop "modelo_input"
                        else:
                            print("Entrada inválida. Por favor, insira 1 ou 2")

                    if continuar == '2':
                        continue  # Volta ao loop "modelo_input"

                    while True:
                        print("\nDeseja simular o projeto, alterar o modelo ou sair? (simular/alterar/sair)")
                        escolha_simulacao = input().strip().lower()
                        if escolha_simulacao == 'simular':
                            simular_projeto(caminho_projeto_novo)
                            print("\nEncerrada a aplicação!")
                            return  # Finaliza a aplicação após a simulação
                        elif escolha_simulacao == 'sair':
                            print("\nEncerrada a aplicação!")
                            return  # Finaliza a aplicação sem simular                       
                        elif escolha_simulacao == 'alterar':
                            break  # Sai do loop "escolha_simulacao" para voltar à escolha do modelo
                        else:
                            print("Entrada inválida. Por favor, insira 'simular' ou 'alterar'.")
                    if escolha_simulacao == 'alterar':
                        break
                if escolha_simulacao == 'alterar':
                    break
            elif opcao == '2':
                break  # Sai do loop "opcao" para voltar à escolha do modelo
            else:
                print("Entrada inválida. Por favor, insira 1 ou 2.")
                continue

    
if __name__ == "__main__":
    escolhe_modelo()