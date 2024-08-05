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
caminho_projeto_novo = input("\nInsira o caminho para o projeto novo do HFSS baseado em suas alterações (ex: C:\\caminho\\para\\seu\\projeto): \n").strip()

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

# Solicita o caminho completo para o ansysedt.exe e o valida
caminho_ansysedt = obter_caminho_valido("\nInsira o caminho para o ansysedtf (ex: C:\\caminho\\para\\seu\\ansysedtsv.aedt):\n")

# Verificar se o projeto do HFSS existe
if not os.path.isfile(caminho_projeto_base):
    print(f"O projeto {caminho_projeto_base} não foi encontrado. Por favor, verifique o caminho e tente novamente.")
    exit()

# Obter apenas o nome do arquivo do projeto (sem a extensão .aedt)

nome_projeto = os.path.splitext(os.path.basename(caminho_projeto_base))[0] 

# Converter barras inclinadas para barras invertidas nos caminhos
caminho_arquivo_altern = caminho_arquivo.replace("\\", "/")
caminho_projeto_base = caminho_projeto_base.replace("\\", "/")
caminho_projeto_novo = caminho_projeto_novo.replace("\\", "/")



def modelo_1():
  ....
def modelo_2():
  ....

def modelo_3():
  .....

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
    tpp_int = int(tpp)  # Salvando apenas a parte inteira de tpp
    file_name = f'k{int(peak)}_{tpp_int}_{int(tmm)}u.tab'
    file_path = os.path.join(diretorio, file_name)
    np.savetxt(file_path, np.column_stack((t * 1e-6, I)), delimiter='\t', header='Time (µs)\tCurrent (kA)', comments='')

    return K, alpha, beta, emax, kk, file_path, file_name
def define_input1():
    diretorio_raio=os.path.dirname(caminho_projeto_novo)
    nome_projeto_novo = os.path.splitext(os.path.basename(caminho_projeto_novo))[0] 
    caminho_arquivo_2 = os.path.join(diretorio_raio, "raio3.py") 
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
            "Coordinate System:="   , "plano do solo",
            "Start:="       , ["0meter","0meter","6.42meter"],
            "End:="         , ["0meter","0meter","6.4meter"]
        ],
        "TimeProfile:="     , "Dataset",
        "HfssFrequency:="   , "1GHz",
        "MinFreq:="     , "0Hz",
        "MaxFreq:="     , "400kHz",
        "Magnitude:="       , "1A",
        "Dataset:="     , "Lightning8_1",
        "DatasetMaxFreq:="  , "400kHz"
    ])
oModule = oDesign.GetModule("AnalysisSetup")
oModule.EditSetup("Setup1", 
    [
        "NAME:Setup1",
        "Frequency:="       , "5GHz",
        "MaxDeltaE:="       , 0.1,
        "MaximumPasses:="   , 20,
        "UseImplicitSolver:="   , True,
        "IsEnabled:="       , True,
        [
            "NAME:MeshLink",
            "ImportMesh:="      , False
        ],
        "BasisOrder:="      , -1,
        [
            "NAME:Transient",
            "ListsForFields:="  , ["solo"],
            "UseSaveCount:="    , 0,
            "UseSaveStart:="    , 1,
            "SaveStart:="       , "0s",
            "UseSaveDelta:="    , 1,
            "SaveDelta:="       , "500ns",
            "SaveRadFields:="   , 0,
            "SaveFDRadFields:=" , 0,
            "UseAutoTermination:="  , 0,
            "TerminateOnMaximum:="  , 1,
            "UseMaxTime:="      , 1,
            "MaxTime:="     , "50us"
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
    plt.plot(df['X '], df['Y'], marker='o')  # Agora você pode usar 'X' e 'Y' diretamente
    plt.title('Descarga 8/25')
    plt.xlabel('Tempo (us)')
    plt.ylabel('Corrente (A)')
    plt.grid(True)
    plt.show() 

def define_input2():
    diretorio_raio=os.path.dirname(caminho_projeto_novo)
    nome_projeto_novo = os.path.splitext(os.path.basename(caminho_projeto_novo))[0] 
    caminho_arquivo_2 = os.path.join(diretorio_raio, "raio4.py")
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
            "Coordinate System:="   , "plano do solo",
            "Start:="       , ["0meter","0meter","6.42meter"],
            "End:="         , ["0meter","0meter","6.4meter"]
        ],
        "TimeProfile:="     , "Dataset",
        "HfssFrequency:="   , "1GHz",
        "MinFreq:="     , "0Hz",
        "MaxFreq:="     , "400kHz",
        "Magnitude:="       , "1A",
        "Dataset:="     , "Lightning1_2_1",
        "DatasetMaxFreq:="  , "400kHz"
    ])
oModule = oDesign.GetModule("AnalysisSetup")
oModule.EditSetup("Setup1", 
    [
        "NAME:Setup1",
        "Frequency:="       , "5GHz",
        "MaxDeltaE:="       , 0.1,
        "MaximumPasses:="   , 20,
        "UseImplicitSolver:="   , True,
        "IsEnabled:="       , True,
        [
            "NAME:MeshLink",
            "ImportMesh:="      , False
        ],
        "BasisOrder:="      , -1,
        [
            "NAME:Transient",
            "ListsForFields:="  , ["solo"],
            "UseSaveCount:="    , 0,
            "UseSaveStart:="    , 1,
            "SaveStart:="       , "0s",
            "UseSaveDelta:="    , 1,
            "SaveDelta:="       , "500ns",
            "SaveRadFields:="   , 0,
            "SaveFDRadFields:=" , 0,
            "UseAutoTermination:="  , 0,
            "TerminateOnMaximum:="  , 1,
            "UseMaxTime:="      , 1,
            "MaxTime:="     , "100us"
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
    plt.plot(df['X '], df['Y'], marker='o')  # Agora você pode usar 'X' e 'Y' diretamente
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
        "Dataset:="        , "{os.path.splitext(file_name)[0]}1",
        "DatasetMaxFreq:="    , "200MHz"
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

def escolhe_modelo():
    continuar_aplicacao = True

    # Solicitar o modelo da malha 
    while continuar_aplicacao:
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
            break  # Garante que o loop saia quando 'sair' for escolhido
            
        else:
            print("Entrada inválida. Por favor, insira 1, 2, 3 ou 'sair'.")
            continue
        
        # O usuário irá escolher se deseja alterar o modelo, ou continuar a aplicação
        while True:
            print("\nSe deseja continuar a aplicação, digite: 1\nSe deseja alterar o modelo, digite: 2 ")
            opcao = input().strip()
            while verificar_HFSS_em_execucao():
                print("O HFSS está em execução. Por favor, feche o programa para prosseguir.")
                time.sleep(5)  # Aguarda 5 segundos antes de verificar novamente   
            
            if opcao == '1':  
                while True:   
                    print("\nPara o input da aplicação, digite 1 para uma descarga atmosférica 8/20us, digite 2 para uma descarga 1.2/50us ou digite 3 para gerar seu próprio modelo.")
                    modelo_input = input().strip()
                    if modelo_input == '1':
                        define_input1()
                        while True:
                            print("\nSe deseja continuar com essa descarga, digite 1. Se não, digite 2)")
                            continuar = input().strip()
                            if continuar == '1':
                                while verificar_HFSS_em_execucao():
                                    print("O HFSS está em execução. Por favor, feche o programa para prosseguir.")
                                    time.sleep(5)  # Aguarda 5 segundos antes de verificar novamente
                                continuar_aplicacao = False  # Sair do loop principal
                                break
                            elif continuar == '2':
                                break  # Retorna à escolha do input
                            else:
                                print("Entrada inválida. Por favor, insira 1 ou 2")
                        if continuar == '1':
                            break
                    
                    elif modelo_input == '2':
                        define_input2()
                        while True:
                            print("\nSe deseja continuar com essa descarga, digite 1. Se não, digite 2)")
                            continuar = input().strip()
                            if continuar == '1':
                                while verificar_HFSS_em_execucao():
                                    print("O HFSS está em execução. Por favor, feche o programa para prosseguir.")
                                    time.sleep(5)  # Aguarda 5 segundos antes de verificar novamente
                                continuar_aplicacao = False  # Sair do loop principal
                                break
                            elif continuar == '2':
                                break  # Retorna à escolha do input
                            else:
                                print("Entrada inválida. Por favor, insira 1 ou 2")
                        if continuar == '1':
                            break
                    elif modelo_input == '3':
                        define_input3()
                        while True:
                            print("\nSe deseja continuar com esse modelo, digite 1. Se não, digite 2)")
                            continuar = input().strip()
                            if continuar == '1':
                                while verificar_HFSS_em_execucao():
                                    print("O HFSS está em execução. Por favor, feche o programa para prosseguir.")
                                    time.sleep(5)  # Aguarda 5 segundos antes de verificar novamente
                                continuar_aplicacao = False  # Sair do loop principal
                                break
                            elif continuar == '2':
                                break  # Retorna à escolha do input
                            else:
                                print("Entrada inválida. Por favor, insira 1 ou 2")
                        if continuar == '1':
                            break                        
                    else:
                        print("Entrada inválida. Por favor, insira 1, 2 ou 3")
                                            
            elif opcao == '2': 
                break  # Sai do loop atual para permitir a re-seleção do modelo
            else:
                print("Entrada inválida. Por favor, insira 1 ou 2.")
                continue

            if continuar == '1':  # Verifica 'continuar' após o loop interno
                break        

    print("Saindo do loop...")   
            
if __name__ == "__main__":
    escolhe_modelo()       

