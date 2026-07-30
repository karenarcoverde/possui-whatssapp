import requests
import pandas as pd
import time


headers = {
    "Client-Token": CLIENT_TOKEN
}

# Caminho do arquivo original (leitura e gravação no mesmo local)
nome_arquivo_excel = "contatos.xlsx" 

try:
    df = pd.read_excel(nome_arquivo_excel)
    
    if 'telefone' not in df.columns:
        print("Erro: A planilha precisa ter uma coluna chamada exatamente 'telefone'.")
        exit()
        
    print(f"Planilha carregada! Iniciando validação de {len(df)} números...\n")

    resultados_whatsapp = []

    for indice, linha in df.iterrows():
        numero_celular = str(linha['telefone']).strip()
        numero_celular = "".join(filter(str.isdigit, numero_celular))
        
        if not numero_celular:
            resultados_whatsapp.append("Número Inválido/Vazio")
            continue

        url = f"https://api.z-api.io/instances/{INSTANCIA_ID}/token/{CLIENT_TOKEN}/phone-exists/{numero_celular}"

        try:
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                data = response.json()
                
                if data.get("exists") is True:
                    print(f"[{indice + 1}] {numero_celular}: POSSUI WhatsApp")
                    resultados_whatsapp.append("Sim")
                else:
                    print(f"[{indice + 1}] {numero_celular}: NÃO possui WhatsApp")
                    resultados_whatsapp.append("Não")
                    
            else:
                print(f"[{indice + 1}] Erro na API para o número {numero_celular}. Status: {response.status_code}")
                resultados_whatsapp.append("Erro na API")

        except Exception as e:
            print(f"[{indice + 1}] Erro de conexão para o número {numero_celular}: {e}")
            resultados_whatsapp.append("Erro de Conexão")
        
        time.sleep(0.5)

    # Adiciona a nova coluna no DataFrame carregado
    df['possui_whatsapp'] = resultados_whatsapp

    # SALVAMENTO NO MESMO ARQUIVO: Sobrescreve o arquivo contatos.xlsx original
    df.to_excel(nome_arquivo_excel, index=False)
    print(f"\nProcesso concluído com sucesso! O arquivo '{nome_arquivo_excel}' foi atualizado com a nova coluna.")

except FileNotFoundError:
    print(f"Erro: O arquivo '{nome_arquivo_excel}' não foi encontrado na pasta atual.")
except Exception as e:
    print(f"Ocorreu um erro geral ao processar a planilha: {e}")
