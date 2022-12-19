from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from google_api_functions import *
from datetime import datetime 
import pandas as pd
import traceback
import json
import time
import os

def carregarParametros():
	with open("parametros.json", "r") as infile:
		parametros = json.load(infile)
	return parametros

def apagarCSVs():
	os.chdir(r'C:\\Users\\'+ user_name +'\\Downloads')
	try:
		nomesDosArquivos = [nomesDosArquivos for nomesDosArquivos in os.listdir() if ('.csv' in nomesDosArquivos) and ('.part' not in nomesDosArquivos)]
		if debug_mode:
			print('Pasta de download limpa.')
		for arquivo in nomesDosArquivos:
			os.remove(arquivo)
		os.chdir(diretorio_robo)
	except IndexError:
		pass
	except Exception as e:
		time.sleep(1)

def baixarArquivoGestaoDePacotes(xpathcompleto, arquivoSolicitado):
	apagarCSVs()
	driver.get('https://envios.mercadolivre.com.br/logistics/management-packages')
	nome_do_arquivo = 'C:\\Users\\' + os.getlogin() + '\\Downloads\\' + 'logistics_packages_' + '-'.join([time.strftime("%d"),time.strftime("%m"),time.strftime("%Y")]) + '.csv'
	if debug_mode:
		print(nome_do_arquivo)
	# Baixa o arquivo
	while True:
		time.sleep(5)
		try:
			time.sleep(1)
			driver.find_element(By.XPATH,'/html/body/main/div/div[2]/div/div/div[1]/button/span').click()
			
			# Loop para clicar no botão do arquivo
			for i in range(20):
				time.sleep(1)
				try:
					emTransito = [emTransito.text for emTransito in driver.find_elements(By.CLASS_NAME, "summarize-column")[2].find_elements(By.CLASS_NAME,'status-card')]
					emTransito = [emTransito_.split('\n')[0] for emTransito_ in emTransito]
					
					naEstacao = [naEstacao.text for naEstacao in driver.find_elements(By.CLASS_NAME, "summarize-column")[1].find_elements(By.CLASS_NAME,'status-card')]
					naEstacao = [naEstacao_.split('\n')[0] for naEstacao_ in naEstacao]
					
					naEstacao2 = [naEstacao2.text for naEstacao2 in driver.find_elements(By.CLASS_NAME, "summarize-column")[1].find_elements(By.CLASS_NAME,'status-card--stale status-card')]
					naEstacao2 = [naEstacao2.split('\n')[0] for naEstacao2 in naEstacao2]
					
					if (arquivoSolicitado not in emTransito) and (arquivoSolicitado not in naEstacao) and (arquivoSolicitado not in naEstacao2):

						columns = ['ID do envio', 'Status do envio', 'Subtatus do envio',
       'Valor declarado', 'Destination Facility Type',
       'Destination Facility ID', 'Rota', 'Promessa de entrega', 'Altura',
       'Largura', 'Comprimento', 'Peso', 'Volume', 'Nome e sobrenome',
       'Telefone', 'Tipo de Endereço', 'Endereço', 'Rua', 'Número',
       'Referências', 'Cidade', 'State', 'bairro', 'Codigo Postal', 'Origem']

						return pd.DataFrame(columns=columns)
					botaodoarquivo = driver.find_element(By.XPATH,xpathcompleto)
					textodobotaodoarquivo = botaodoarquivo.text.split('\n')[0]
					time.sleep(3)
					botaodoarquivo.click()
					if debug_mode:
						print('Loop para clicar no botão do arquivo')
					break
				except Exception as e:
					if debug_mode:
						print(traceback.format_exc())
					pass
			
			# Loop para clicar no botão de baixar o arquivo
			for i in range(20):
				time.sleep(1)
				try:
					botaobaixar = driver.find_element(By.XPATH,'/html/body/main/div/div[2]/div/div/div[2]/div[2]/button')
					time.sleep(3)
					botaobaixar.click()
					break
				except Exception as e:
					if debug_mode:
						print(traceback.format_exc())
					pass
			break	
		except:
			funcaoPrincipal()
	
	# Carrega e trata o arquivo
	while True:
		time.sleep(1)
		try:
			os.chdir(f'C:\\Users\\{user_name}\\Downloads')
			for i in range(20):
				time.sleep(1)
				try:
					nomeArquivoGestaoDePacotes = [nomesDosArquivos for nomesDosArquivos in os.listdir() if ('logistics_packages' in nomesDosArquivos) and ('.part' not in nomesDosArquivos)][0]
					break
				except Exception as e:
					if debug_mode:
						print(traceback.format_exc())
					if i == 19:
						raise IndexError
					pass
			arquivoGestaoDePacotes = pd.read_csv(nomeArquivoGestaoDePacotes)
			os.remove(nomeArquivoGestaoDePacotes)
			os.chdir(diretorio_robo)
			arquivoGestaoDePacotes['ID do envio'] = pd.to_numeric(arquivoGestaoDePacotes['ID do envio'], errors='coerce')
			arquivoGestaoDePacotes = arquivoGestaoDePacotes.loc[~ (arquivoGestaoDePacotes['ID do envio'].isna())]
			arquivoGestaoDePacotes['ID do envio'] = arquivoGestaoDePacotes['ID do envio'].astype('str').str[:11]
			arquivoGestaoDePacotes = arquivoGestaoDePacotes.fillna('')
			print(f'Arquivo "{textodobotaodoarquivo}" carregado')
			return arquivoGestaoDePacotes
		except Exception as e:
			if debug_mode:
				print(traceback.format_exc())
				print(e)
			raise NotImplementedError

def funcaoPrincipal():
	while True:
		try:
			apagarCSVs()

			emRotaDeEntrega = baixarArquivoGestaoDePacotes('//*[@id="packages-management"]/div[1]/div/div[2]/div/div[3]/div[3]','Em rota de entrega').values.tolist() 
			entregue = baixarArquivoGestaoDePacotes('//*[@id="packages-management"]/div[1]/div/div[2]/div/div[3]/div[1]','Entregues').values.tolist()
			falhaDeEntrega = baixarArquivoGestaoDePacotes('//*[@id="packages-management"]/div[1]/div/div[2]/div/div[3]/div[2]','Falha de entrega').values.tolist()		
			solucaoDeProblemas = baixarArquivoGestaoDePacotes('/html/body/main/div/div[2]/div/div/div[1]/div/div[2]/div/div[2]/div[5]','Para solução de problemas').values.tolist()		
			paraDevolucao = baixarArquivoGestaoDePacotes('/html/body/main/div/div[2]/div/div/div[1]/div/div[2]/div/div[2]/div[3]','Para devolução').values.tolist()		
			agora = time.strftime('%H:%M')
			
			ID_PLANILHA_BASE_COCKPIT = "1x3t-0JsNwN38FajdWNWlN9Z_cEbjz-BQqnchy-KjmWQ"
			
			limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'FALHA DE ENTREGA!A2:Y')
			update_values(ID_PLANILHA_BASE_COCKPIT,'FALHA DE ENTREGA!A2','USER_ENTERED',falhaDeEntrega)
			
			limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'EM ROTA DE ENTREGA!A2:Y')
			update_values(ID_PLANILHA_BASE_COCKPIT,'EM ROTA DE ENTREGA!A2','USER_ENTERED',emRotaDeEntrega)

			limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'ENTREGUES!A2:Y')
			update_values(ID_PLANILHA_BASE_COCKPIT,'ENTREGUES!A2','USER_ENTERED',entregue)
			
			limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'SOLUÇÃO DE PROBLEMAS!A2:Y')
			update_values(ID_PLANILHA_BASE_COCKPIT,'ENTREGUES!A2','USER_ENTERED',solucaoDeProblemas)
			
			limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'DEVOLUÇÃO!A2:Y')
			update_values(ID_PLANILHA_BASE_COCKPIT,'ENTREGUES!A2','USER_ENTERED',paraDevolucao)
			
			print(f'Última atualização: {agora}')
			
			print('Pausa para acompanhamento...')
			# time.sleep(int(carregarParametros()["delayacompanhamento"])*60)
			time.sleep(5*60)

		except Exception as e:
			if debug_mode:
				print(e)
				print(traceback.format_exc())
			pass

diretorio_robo = os.getcwd()
user_name = os.getlogin()
debug_mode = False

print('Abrindo driver Firefox')
# profile_path = r'C:\Users\vdiassob\AppData\Roaming\Mozilla\Firefox\Profiles\eituekku.robo'
options = Options()
# options.add_argument("-profile")
# options.add_argument(profile_path)
options.binary_location = carregarParametros()["caminhonavegador"]
driver = webdriver.Firefox(options=options)
driver.get('https://envios.mercadolivre.com.br/logistics/routing/planification/download')

input('Após logar no logistics, pressione ENTER para continuar...\n')

while True:
	try:
		funcaoPrincipal()
	except:
		pass
