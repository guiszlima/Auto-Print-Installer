import json
import subprocess 
def get_json_data(json_file):
    try:
        with open(json_file, 'r', encoding='utf-8') as file:
            data = json.load(file)
            return data 
     
    except Exception as e:
        return f"Ocorreu um erro: {e}"

def instala_impressora(nomePrinter,nomeDriver,ipAdress,caminhoParaDriver,caminhoServidorDriver):


        nomePrinterFormat = f""" "{nomePrinter}" """ 

        nomeDriverFormat = f""" "{nomeDriver}" """ 

        ipAdressFormat =  f""" "{ipAdress}" """ 

        portName = f""" "IP_{ipAdress}" """ 

        protocolName = """ "LPR" """ 

        caminhoParaDriverFormat = f""" "{caminhoParaDriver}" """ 
        
        destino = r"C:\\"

        comandoCopiarFile = f'Copy-Item "{caminhoServidorDriver}" -Destination "{destino}" -Recurse -Force'

        

        comandoInstalarDriver = f"printui /ia /m {nomeDriverFormat} /f {caminhoParaDriverFormat} /h x64" 

        comandoInstalarPort = f"Add-PrinterPort  -Name  {portName} -LprHostAddress {ipAdressFormat}  -LprQueueName {protocolName} " 
        comandoInstalarPrinter = f"Add-Printer -Name {nomePrinterFormat} -DriverName {nomeDriverFormat} -PortName {portName}" 
       
        subprocess.run(["powershell", "-Command", comandoCopiarFile ], shell=True) 

        subprocess.run(["powershell", "-Command", comandoInstalarDriver ], shell=True) 

        subprocess.run(["powershell", "-Command", comandoInstalarPort ], shell=True) 

        subprocess.run(["powershell", "-Command", comandoInstalarPrinter ], shell=True) 

json_file = 'registros/impressoras.json'
data = get_json_data(json_file)

impressoras_array = []

while True:
    edificios_array = []
    c = 0
    for predio,key in data.items():
        edificios_array.append(predio)
    
        print(c,predio)
        c+= 1
    escolhaPredio = int(input('Qual predio se encontra? Escolha um número:\n'))
    if escolhaPredio > len(edificios_array):
        print("Prédio Invalido")
    else:
        break
impressoras =  data[edificios_array[escolhaPredio]]
for impressora , keys in impressoras.items():
    impressoras_array.append(impressora)
    print(impressora)


while True:
    broke = False
    
    for impressora, keys in impressoras.items():
        escolhaImpressora = input('Qual Impressora deseja escolher?\n')
        if escolhaImpressora  in impressoras_array:
            broke = True
            break
        else:
            print("Impressora Não Existe") 
    if broke == True:
        break        
for impressora, keys in impressoras.items():
    if impressora == escolhaImpressora:
        print("Instalando Impressora...")
        
        if isinstance(keys, dict):  # Verifica se keys é um dicionário
            nomePrinter = keys.get('Nome_Impressora', '')
            nomeDriver = keys.get('Nome_Driver', '')
            ipAdress = keys.get('Endereco_IP', '')
            caminhoParaDriver = keys.get('Caminho_Driver', '')
            caminhoServidorDriver = keys.get('Caminho_Driver_Servidor', '')
            

            instala_impressora(nomePrinter, nomeDriver, ipAdress, caminhoParaDriver, caminhoServidorDriver)
            print("Impressora Instalada com Sucesso!")
        else:
            print("Impressora não encontrada.")
        break







