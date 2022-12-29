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
	os.chdir(f'{diretorio_robo}\\Downloads')
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
		print(traceback.format_exc())

def baixarArquivoGestaoDePacotes(xpathcompleto, arquivoSolicitado):
	apagarCSVs()
	driver.get('https://envios.mercadolivre.com.br/logistics/management-packages')
	# Baixa o arquivo
	while True:
		time.sleep(5)
		try:
			for i in range(10):
				time.sleep(1)
				driver.find_element(By.CLASS_NAME,'summarize-button').click()
				break
			
			# Loop para clicar no botão do arquivo
			
			if debug_mode:
				print('Loop para clicar no botão do arquivo')
			
			for i in range(20):
				time.sleep(1)
				try:
					botoes_gestaodepacotes = [i for i in driver.find_elements(By.CLASS_NAME, 'status-card__title') if i.text == arquivoSolicitado]
					if len(botoes_gestaodepacotes) < 1:
						raise NotImplementedError
					else:
						botoes_gestaodepacotes[0].click()
						textodobotaodoarquivo = botoes_gestaodepacotes[0].text
						break
				except Exception as e:
					if debug_mode:
						print(traceback.format_exc())
					textodobotaodoarquivo = 'Indisponível'
					pass

			if textodobotaodoarquivo == 'Indisponível':

				columns = ['ID do envio', 'Status do envio', 'Subtatus do envio',
'Valor declarado', 'Destination Facility Type',
'Destination Facility ID', 'Rota', 'Promessa de entrega', 'Altura',
'Largura', 'Comprimento', 'Peso', 'Volume', 'Nome e sobrenome',
'Telefone', 'Tipo de Endereço', 'Endereço', 'Rua', 'Número',
'Referências', 'Cidade', 'State', 'bairro', 'Codigo Postal', 'Origem']
			
				return pd.DataFrame(columns=columns)

			# Loop para clicar no botão de baixar o arquivo
			for i in range(20):
				time.sleep(1)
				try:
					botaobaixar = driver.find_element(By.CLASS_NAME,'downloader')
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
			os.chdir(f'{diretorio_robo}\\Downloads')
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
			apagarCSVs()
			# os.remove(nomeArquivoGestaoDePacotes)
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
			
			ID_PLANILHA_BASE_COCKPIT = carregarParametros()["ID_PLANILHA_BASE_COCKPIT"]
			
			for i in range(5):
				try:
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
					break
				except:
					if debug_mode:
						print(traceback.format_exc())
					pass
				driver.switch_to.window(driver.window_handles[1])
				
				for i in range(5):
					time.sleep(1)
					try:
						driver.find_elements(By.ID, 'more-options-header-menu-button')[0].click()
						break
					except:
						pass
				for i in range(5):
					time.sleep(1)
					try:
						driver.find_elements(By.ID, 'header-refresh-button')[0].click()
						break
					except:
						pass	
				
				driver.switch_to.window(driver.window_handles[0])
			
			print(f'Última atualização: {agora}')
			
			print('Pausa para acompanhamento...')
			time.sleep(int(carregarParametros()["delayAcompanhamentoExpedicao"])*60)

		except Exception as e:
			if debug_mode:
				print(e)
				print(traceback.format_exc())
			pass

def verificarPastaDownloads():
	if os.path.isdir(os.getcwd()+'\\Downloads'):
		return True
	else:
		os.mkdir(os.getcwd()+'\\Downloads')


verificarPastaDownloads()
diretorio_robo = os.getcwd()
user_name = os.getlogin()
debug_mode = True

print('Abrindo driver Firefox')
profile_path = carregarParametros()["perfilFirefox"]
options = Options()
options.add_argument("-profile")
options.add_argument(profile_path)
options.binary_location = carregarParametros()["caminhonavegador"]
driver = webdriver.Firefox(options=options)
driver.get('https://envios.mercadolivre.com.br/logistics/routing/planification/download')
driver.execute_script("window.open('');")
driver.switch_to.window(driver.window_handles[1])
driver.get('https://datastudio.google.com/reporting/96427ae2-bf8a-4c8d-9260-35d32154d663/page/TuM7C')
driver.switch_to.window(driver.window_handles[0])

input('Após logar no logistics, pressione ENTER para continuar...\n')

while True:
	try:
		funcaoPrincipal()
	except:
		pass
