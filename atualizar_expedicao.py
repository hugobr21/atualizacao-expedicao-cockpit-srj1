from google_api_functions import *
# from atualizacao_goleiro import *
from datetime import datetime 
import pandas as pd
import json
import time
import os

def carregarParametros():
	with open("parametros.json", "r") as infile:
		parametros = json.load(infile)
	return parametros

def importarBasesDeRoteirizacao(nome_planilha,ciclo):
	try:
		try:
			planilhaBase = pd.read_excel([nomeDoArquivo for nomeDoArquivo in os.listdir() if nome_planilha in nomeDoArquivo][0],sheet_name='Planilha1')
		except:
			planilhaBase = pd.read_excel([nomeDoArquivo for nomeDoArquivo in os.listdir() if nome_planilha in nomeDoArquivo][0],sheet_name='Plan1')
		planilhaBase['Ciclo'] = ciclo
		planilhaBase['Shipment'] = planilhaBase['Shipment'].astype('str')
		return planilhaBase[['Shipment','Ciclo','Rota']]
	except:
		try:
			planilhaBase = pd.read_excel(nome_planilha + '.xlsm',sheet_name='Planilha1')
			planilhaBase['Ciclo'] = ciclo
			planilhaBase['Shipment'] = planilhaBase['Shipment'].astype('str')
			return planilhaBase[['Shipment','Ciclo','Rota']]
		except:
			return pd.DataFrame({'Shipment':[],'Rota':[],'Ciclo':[]})

def importarPlanification():
	try:
		planification = pd.read_csv([nomeDoArquivo for nomeDoArquivo in os.listdir() if 'planification' in nomeDoArquivo][0])
		planification['Shipment'] = planification['Shipment'].astype('str')
		return planification
	except:
		print('Um erro ocorreu ao importar planification.')

def atualizarBase(id_planilha,aba_range,dados_da_base):
	limpar_celulas(id_planilha,aba_range)
	update_values(id_planilha,aba_range, 'USER_ENTERED', dados_da_base)

def apagarCSVs():
	os.chdir(r'C:\\Users\\'+ user_name +'\\Downloads')
	try:
		nomesDosArquivos = [nomesDosArquivos for nomesDosArquivos in os.listdir() if ('.csv' in nomesDosArquivos) and ('.part' not in nomesDosArquivos)]
		print('Pasta de download limpa.')
		for arquivo in nomesDosArquivos:
			os.remove(arquivo)
		os.chdir(diretorio_robo)
	except IndexError:
		pass
	except Exception as e:
		time.sleep(1)

def baixar_planification():
	driver.get('https://envios.mercadolivre.com.br/logistics/routing/planification/download')
	nome_do_arquivo = 'C:\\Users\\' + os.getlogin() + '\\Downloads\\' + 'planification_' + '_'.join([time.strftime("%d"),time.strftime("%m"),time.strftime("%Y")]) + '.csv'
	# print(nome_do_arquivo)
	contador = 0
	while True:
		time.sleep(5)
		try:
			driver.find_element(By.XPATH,'//*[@id="routing-downloads"]').click()
			botao_baixar_planification = driver.find_element(By.XPATH,'/html/body/main/div/div/div/div/div/div[7]/button')
			botao_baixar_planification.click()
			# print('clicar p baixar')
			break	
		except:
			if contador >= 10:
				print('#1 Reiniciando página e download')
				driver.get("https://envios.mercadolivre.com.br/logistics/routing/planification/download")
				time.sleep(6)
				# funcao_principal()
			print('erro ao baixar arquivo')
			contador = contador + 1
			pass
	contador = 0
	while True:
		time.sleep(1)
		try:
			os.chdir(f'C:\\Users\\{user_name}\\Downloads')
			planification = pd.read_csv([nomesDosArquivos for nomesDosArquivos in os.listdir() if ('planification' in nomesDosArquivos) and ('.part' not in nomesDosArquivos)][0])
			os.chdir(diretorio_robo)
			planification['Shipment'] = pd.to_numeric(planification['Shipment'], errors='coerce')
			planification = planification.loc[~ (planification['Status'].isna())]
			planification = planification.loc[~ (planification['Shipment'].isna())]
			planification['Shipment'] = planification['Shipment'].astype('str').str[:11]
			os.remove(nome_do_arquivo)
			print('Arquivo carregado')
			return planification
		except Exception as e:
			print(e)
			contador = contador + 1
			if contador >=100:
				try:
					os.remove(nome_do_arquivo)
				except:
					pass
				print('#2 Reiniciando página e download')
				driver.get("https://envios.mercadolivre.com.br/logistics/routing/planification/download")
				time.sleep(6)
				raise KeyError
			# print('erro ao carregar arquivo')
			pass

def funcao_principal(planilha_base_dash_svc,
baseroteirizacao_aba_range,
baseDeRoteirizacaoConsolidada,
planification_aba_range,
planification_roteirizacao):

	# Sobe base de roteirizacao
	atualizarBase(planilha_base_dash_svc, baseroteirizacao_aba_range, baseDeRoteirizacaoConsolidada)

	# Sobe planification
	atualizarBase(planilha_base_dash_svc, planification_aba_range, planification_roteirizacao)

	print('Base atualizada!')
	
def carregarBaseEtiquetada():
	data_agora = datetime(datetime.now().year,datetime.now().month,datetime.now().day,00,00,00).strftime("_%d_%m_%Y")
	try:
		base_atstation = pd.read_excel(f'mudanca_de_status_etiquetagem{data_agora}.xlsx')
	except:
		base_atstation = pd.DataFrame({'Shipment':[],'Mudança de Status':[],'Hora':[]})
		base_atstation.to_excel(f'mudanca_de_status_etiquetagem{data_agora}.xlsx', index=False)
	base_atstation['Shipment'] = base_atstation['Shipment'].astype('str')
	return base_atstation

def carregarBaseSorteado():
	data_agora = datetime(datetime.now().year,datetime.now().month,datetime.now().day,00,00,00).strftime("_%d_%m_%Y")
	try:
		base_sorteado = pd.read_excel(f'mudanca_de_status_sorteado{data_agora}.xlsx')
	except:
		base_sorteado = pd.DataFrame({'Shipment':[],'Mudança de Status':[],'Hora':[]})
		base_sorteado.to_excel(f'mudanca_de_status_sorteado{data_agora}.xlsx', index=False)
	base_sorteado['Shipment'] = base_sorteado['Shipment'].astype('str')
	return base_sorteado

def baixarEmRotaDeEntrega():
	driver.get('https://envios.mercadolivre.com.br/logistics/management-packages')
	nome_do_arquivo = 'C:\\Users\\' + os.getlogin() + '\\Downloads\\' + 'logistics_packages_' + '-'.join([time.strftime("%d"),time.strftime("%m"),time.strftime("%Y")]) + '.csv'
	print(nome_do_arquivo)
	contador = 0
	# Baixa o arquivo
	while True:
		time.sleep(5)
		try:
			time.sleep(1)
			driver.find_element(By.XPATH,'//*[@id="packages-management"]/div[1]/button').click()
			time.sleep(1)
			driver.find_element(By.XPATH,'//*[@id="packages-management"]/div[1]/div/div[2]/div/div[3]/div[3]').click()
			time.sleep(1)
			driver.find_element(By.XPATH,'/html/body/main/div/div[2]/div/div/div[2]/div[2]/button').click()
			# botao_baixar_emrotadeentrega = driver.find_element(By.XPATH,'/html/body/main/div/div/div/div/div/div[7]/button')
			# botao_baixar_emrotadeentrega.click()
			# # print('clicar p baixar')
			break	
		except:
			if contador >= 10:
				print('#1 Reiniciando página e download')
				driver.get("https://envios.mercadolivre.com.br/logistics/management-packages")
				time.sleep(6)
				# funcao_principal()
			print('erro ao baixar arquivo')
			contador = contador + 1
			pass
	contador = 0
	# Carrega e trata o arquivo
	while True:
		time.sleep(1)
		try:
			os.chdir(f'C:\\Users\\{user_name}\\Downloads')
			emrotadeentrega = pd.read_csv([nomesDosArquivos for nomesDosArquivos in os.listdir() if ('logistics_packages' in nomesDosArquivos) and ('.part' not in nomesDosArquivos)][0])
			os.chdir(diretorio_robo)
			emrotadeentrega['ID do envio'] = pd.to_numeric(emrotadeentrega['ID do envio'], errors='coerce')
			emrotadeentrega = emrotadeentrega.loc[~ (emrotadeentrega['ID do envio'].isna())]
			emrotadeentrega['ID do envio'] = emrotadeentrega['ID do envio'].astype('str').str[:11]
			emrotadeentrega = emrotadeentrega.fillna('')
			os.remove(nome_do_arquivo)
			print('Arquivo em rota de entrega carregado')
			return emrotadeentrega
		except Exception as e:
			print(e)
			contador = contador + 1
			if contador >=100:
				try:
					os.remove(nome_do_arquivo)
				except:
					pass
				print('#2 Reiniciando página e download')
				driver.get("https://envios.mercadolivre.com.br/logistics/management-packages")
				time.sleep(6)
				raise KeyError
			# print('erro ao carregar arquivo')
			pass

def baixarEntregues():
	driver.get('https://envios.mercadolivre.com.br/logistics/management-packages')
	nome_do_arquivo = 'C:\\Users\\' + os.getlogin() + '\\Downloads\\' + 'logistics_packages_' + '-'.join([time.strftime("%d"),time.strftime("%m"),time.strftime("%Y")]) + '.csv'
	print(nome_do_arquivo)
	contador = 0
	# Baixa o arquivo
	while True:
		time.sleep(5)
		try:
			time.sleep(1)
			driver.find_element(By.XPATH,'//*[@id="packages-management"]/div[1]/button').click()
			time.sleep(1)
			driver.find_element(By.XPATH,'//*[@id="packages-management"]/div[1]/div/div[2]/div/div[3]/div[1]').click()
			time.sleep(1)
			driver.find_element(By.XPATH,'/html/body/main/div/div[2]/div/div/div[2]/div[2]/button').click()
			break	
		except:
			if contador >= 10:
				print('#1 Reiniciando página e download')
				driver.get("https://envios.mercadolivre.com.br/logistics/management-packages")
				time.sleep(6)
			print('erro ao baixar arquivo')
			contador = contador + 1
			pass
	contador = 0
	# Carrega e trata o arquivo
	while True:
		time.sleep(1)
		try:
			os.chdir(f'C:\\Users\\{user_name}\\Downloads')
			entregues = pd.read_csv([nomesDosArquivos for nomesDosArquivos in os.listdir() if ('logistics_packages' in nomesDosArquivos) and ('.part' not in nomesDosArquivos)][0])
			os.chdir(diretorio_robo)
			entregues['ID do envio'] = pd.to_numeric(entregues['ID do envio'], errors='coerce')
			entregues = entregues.loc[~ (entregues['ID do envio'].isna())]
			entregues['ID do envio'] = entregues['ID do envio'].astype('str').str[:11]
			entregues = entregues.fillna('')
			os.remove(nome_do_arquivo)
			print('Arquivo baixar entregues carregado')
			return entregues
		except Exception as e:
			print(e)
			contador = contador + 1
			if contador >=100:
				try:
					os.remove(nome_do_arquivo)
				except:
					pass
				print('#2 Reiniciando página e download')
				driver.get("https://envios.mercadolivre.com.br/logistics/management-packages")
				time.sleep(6)
				# raise KeyError
			# print('erro ao carregar arquivo')
			pass

def baixarFalhaDeEntrega():
	driver.get('https://envios.mercadolivre.com.br/logistics/management-packages')
	nome_do_arquivo = 'C:\\Users\\' + os.getlogin() + '\\Downloads\\' + 'logistics_packages_' + '-'.join([time.strftime("%d"),time.strftime("%m"),time.strftime("%Y")]) + '.csv'
	print(nome_do_arquivo)
	contador = 0
	# Baixa o arquivo
	while True:
		time.sleep(5)
		try:
			time.sleep(1)
			driver.find_element(By.XPATH,'//*[@id="packages-management"]/div[1]/button').click()
			time.sleep(1)
			driver.find_element(By.XPATH,'//*[@id="packages-management"]/div[1]/div/div[2]/div/div[3]/div[2]').click()
			time.sleep(1)
			driver.find_element(By.XPATH,'/html/body/main/div/div[2]/div/div/div[2]/div[2]/button').click()
			break	
		except:
			if contador >= 10:
				print('#1 Reiniciando página e download - baixarfalhadeentrega')
				driver.get("https://envios.mercadolivre.com.br/logistics/management-packages")
				time.sleep(6)
			print('erro ao baixar arquivo')
			contador = contador + 1
			pass
	contador = 0
	# Carrega e trata o arquivo
	while True:
		time.sleep(1)
		try:
			os.chdir(f'C:\\Users\\{user_name}\\Downloads')
			falhaDeEntrega = pd.read_csv([nomesDosArquivos for nomesDosArquivos in os.listdir() if ('logistics_packages' in nomesDosArquivos) and ('.part' not in nomesDosArquivos)][0])
			os.chdir(diretorio_robo)
			falhaDeEntrega['ID do envio'] = pd.to_numeric(falhaDeEntrega['ID do envio'], errors='coerce')
			falhaDeEntrega = falhaDeEntrega.loc[~ (falhaDeEntrega['ID do envio'].isna())]
			falhaDeEntrega['ID do envio'] = falhaDeEntrega['ID do envio'].astype('str').str[:11]
			falhaDeEntrega = falhaDeEntrega.fillna('')
			os.remove(nome_do_arquivo)
			print(f'Arquivo falha de entrega carregado - baixarfalhadeentrega')
			return falhaDeEntrega
		except Exception as e:
			print(e)
			contador = contador + 1
			if contador >=100:
				try:
					os.remove(nome_do_arquivo)
				except:
					pass
				print('#2 Reiniciando página e download - baixarfalhadeentrega')
				driver.get("https://envios.mercadolivre.com.br/logistics/management-packages")
				time.sleep(6)
				# raise KeyError
			# print('erro ao carregar arquivo')
			pass

def funcaoPrincipal():
	while True:
		try:
			apagarCSVs()

			emRotaDeEntrega = baixarEmRotaDeEntrega().values.tolist()
			entregue = baixarEntregues().values.tolist()
			falhaDeEntrega = baixarFalhaDeEntrega().values.tolist()		
			agora = time.strftime('%H:%M')
			
			ID_PLANILHA_BASE_COCKPIT = '1x3t-0JsNwN38FajdWNWlN9Z_cEbjz-BQqnchy-KjmWQ'
			
			limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'FALHA DE ENTREGA!A2:Y')
			update_values(ID_PLANILHA_BASE_COCKPIT,'FALHA DE ENTREGA!A2','USER_ENTERED',falhaDeEntrega)
			
			limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'EM ROTA DE ENTREGA!A2:Y')
			update_values(ID_PLANILHA_BASE_COCKPIT,'EM ROTA DE ENTREGA!A2','USER_ENTERED',emRotaDeEntrega)

			limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'ENTREGUES!A2:Y')
			update_values(ID_PLANILHA_BASE_COCKPIT,'ENTREGUES!A2','USER_ENTERED',entregue)
			


			print(f'Última atualização: {agora}')
			# exit()
			print('Pausa para acompanhamento...')
			time.sleep(int(carregarParametros()["delayacompanhamento"])*60)

		except Exception as e:
			if debug_mode:
				print(e)
			pass

diretorio_robo = os.getcwd()
user_name = os.getlogin()
debug_mode = True

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
