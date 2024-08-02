import subprocess
import time
import os
import sys
import psutil

# Solicitar ao usuário que insira o caminho para o arquivo Python
caminho_arquivo = input("\nInsira o caminho para o arquivo Python (ex: C:\\caminho\\para\\seu\\arquivo.py): \n").strip()

# Solicitar ao usuário que insira o caminho para o projeto base do HFSS
caminho_projeto_base = input("\nInsira o caminho para o projeto base do HFSS (ex: C:\\caminho\\para\\seu\\projeto.aedt): \n").strip()

# Solicitar ao usuário que insira o caminho para o projeto base do HFSS
caminho_projeto_novo = input("\nInsira o caminho para o projeto novo do HFSS baseado em suas alterações (ex: C:\\caminho\\para\\seu\\projeto): \n").strip()

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
					"Value:="		, "{comprimento_haste/2}"
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
oProject.SaveAs("{caminho_projeto_novo}", True)
"""
    print("Executando o modelo 1...")
    # Escreve o código HFSS no arquivo Python
    with open(caminho_arquivo, "w") as arquivo:
        arquivo.write(codigo_hfss)

    # Caminho completo para o ansysedt.exe
    caminho_ansysedt = r"C:\Program Files\AnsysEM\Ansys Student\v232\Win64\ansysedtsv.exe"

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
        
   # Aguarde alguns segundos antes de fechar o HFSS
    tempo_de_espera = 30  # ajuste conforme necessário
    time.sleep(tempo_de_espera)    

def modelo_2():
       # Solicitar ao usuário que insira o comprimento da haste 2-6m 
    comprimento_haste = float(input("\nInsira o comprimento da malha quadrada: \n").strip())

   # Solicitar ao usuário a espessura das hastes de cobre, entre 10 -15 mm
    diametro_hastes = float(input("\nInsira o diâmetro das hastes entre 10-15 mm: \n").strip())
    
	#Solicitar ao usuário a dimensão das hastes verticais
    comprimento_haste_vertical = float(input("\nInsira o comprimento das hastes verticais: \n").strip())
     
	 # Solicitar ao usuário a resistividade do solo 
    resistividade = float(input("\nInsira a resistividade do solo: \n").strip()) 
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
					"Value:="		, "{comprimento_haste/2}"
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
oDesign = oProject.SetActiveDesign("Malha_unitaria")
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "HasteX2"
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
				"HasteX3"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Name",
					"Value:="		, "Haste_vertical"
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
				"Haste_vertical:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "$L*(2/3)",
					"Y:="			, "0",
					"Z:="			, "0meter"
				],
				[
					"NAME:Point2",
					"X:="			, "$L*(2/3)",
					"Y:="			, "0",
					"Z:="			, "-$Lz"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_vertical"
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
				"Haste_vertical1:CreatePolyline:1"
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
				"Haste_vertical1:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "$L*(2/3)",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "0meter"
				],
				[
					"NAME:Point2",
					"X:="			, "$L*(2/3)",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "-$Lz"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_vertical"
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
				"Haste_vertical2:CreatePolyline:1"
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
				"Haste_vertical2:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "$L*(2/3)",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "0meter"
				],
				[
					"NAME:Point2",
					"X:="			, "$L*(2/3)",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "-$Lz"
				]
			]
		]
	])
oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_vertical"
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
				"Haste_vertical3:CreatePolyline:1"
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
				"Haste_vertical3:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "0meter"
				],
				[
					"NAME:Point2",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "-$Lz"
				]
			]
		]
	])
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_vertical"
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
				"Haste_vertical4:CreatePolyline:1"
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
				"Haste_vertical4:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "0",
					"Z:="			, "0meter"
				],
				[
					"NAME:Point2",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "0",
					"Z:="			, "-$Lz"
				]
			]
		]
	])
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_vertical"
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
				"Haste_vertical5:CreatePolyline:1"
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
				"Haste_vertical5:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "0",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "0meter"
				],
				[
					"NAME:Point2",
					"X:="			, "0",
					"Y:="			, "$L*(2/3)",
					"Z:="			, "-$Lz"
				]
			]
		]
	])
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_vertical"
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
				"Haste_vertical6:CreatePolyline:1"
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
				"Haste_vertical6:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "0",
					"Z:="			, "0meter"
				],
				[
					"NAME:Point2",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "0",
					"Z:="			, "-$Lz"
				]
			]
		]
	])
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_vertical"
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
				"Haste_vertical7:CreatePolyline:1"
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
				"Haste_vertical7:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "0meter"
				],
				[
					"NAME:Point2",
					"X:="			, "-$L*(2/3)",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "-$Lz"
				]
			]
		]
	])
oEditor.Copy(
	[
		"NAME:Selections",
		"Selections:="		, "Haste_vertical"
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
				"Haste_vertical8:CreatePolyline:1"
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
				"Haste_vertical8:CreatePolyline:1:Segment0"
			],
			[
				"NAME:ChangedProps",
				[
					"NAME:Point1",
					"X:="			, "0",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "0meter"
				],
				[
					"NAME:Point2",
					"X:="			, "0",
					"Y:="			, "-$L*(2/3)",
					"Z:="			, "-$Lz"
				]
			]
		]
	])
oProject.SaveAs("{caminho_projeto_novo}", True)
"""
    print("Executando o modelo 2...")
    # Escreve o código HFSS no arquivo Python
    with open(caminho_arquivo, "w") as arquivo:
        arquivo.write(codigo_hfss)

    # Caminho completo para o ansysedt.exe
    caminho_ansysedt = r"C:\Program Files\AnsysEM\Ansys Student\v232\Win64\ansysedtsv.exe"

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
        
   # Aguarde alguns segundos antes de fechar o HFSS
    tempo_de_espera = 30  # ajuste conforme necessário
    time.sleep(tempo_de_espera)   
    
def modelo_3():
       # Solicitar ao usuário que insira o comprimento da haste 2-6m 
    comprimento_haste = float(input("\nInsira o comprimento da malha quadrada: \n").strip())

   # Solicitar ao usuário a espessura das hastes de cobre, entre 10 -15 mm
    diametro_hastes = float(input("\nInsira o diâmetro das hastes entre 10-15 mm: \n").strip())
    
	#Solicitar ao usuário a dimensão das hastes verticais
    comprimento_haste_vertical = float(input("\nInsira o comprimento das hastes verticais inclinadas: \n").strip())
     
	 # Solicitar ao usuário a resistividade do solo 
    resistividade = float(input("\nInsira a resistividade do solo: \n").strip()) 
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
					"Value:="		, "{comprimento_haste/2}"
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
oProject.SaveAs("{caminho_projeto_novo}", True)
"""
    print("Executando o modelo 3...")
    # Escreve o código HFSS no arquivo Python
    with open(caminho_arquivo, "w") as arquivo:
        arquivo.write(codigo_hfss)

    # Caminho completo para o ansysedt.exe
    caminho_ansysedt = r"C:\Program Files\AnsysEM\Ansys Student\v232\Win64\ansysedtsv.exe"

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
        
   # Aguarde alguns segundos antes de fechar o HFSS
    tempo_de_espera = 30  # ajuste conforme necessário
    time.sleep(tempo_de_espera)

# Solicitar o modelo da malha 
while True:
    print("\nInsira o modelo desejado (1, 2 ou 3) ou digite 'sair' para sair da aplicação.\n Modelo 1:Malha horizontal.\n Modelo 2:Malha com hastes verticais.\n Modelo 3:Malha com hastes verticia inclinadas.")
    modelo_malha = input().strip()

    if modelo_malha == '1':
        modelo_1()
        break
    elif modelo_malha == '2':
        modelo_2()
        break
    elif modelo_malha == '3':
        modelo_3()
        break
    elif modelo_malha.lower() == 'sair':
        print("Saindo da aplicação...")
        break
    else:
        print("Entrada inválida. Por favor, insira 1, 2, 3 ou 'sair'.")



