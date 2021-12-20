from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import warnings
import pandas as pd
from os import listdir
from os.path import isfile, join
from pandas.core.frame import DataFrame
from selenium.webdriver.common.by import By
import openpyxl


def selecionar_cliente():
  '''
  Função para selecionar a lista de contas e o cliente a ser analisado
  '''
  clientes = {'havan': ['024101','024102','024081','024103','024086','024104','024082','024105','024083','024106','024087','024107','024085','024108','024084','024109','024096','024110','024097','024111','024098','024112','024099','024113','024100','024114','024309','024310','024269','024270','024271','024272','024289','024290','024291','024292','024293','024294','024295','024296','024297','024298','024299','024300','024301','024302','024303','024304','024331','024332','024333','024334','024335','024336','024212','024213','024337','024338','024339','024340','024341','024342','024343','024344','024347','024348','024353','024354','024355','024356','024357','024358','024359','024360','024361','024362','024363','024364','024365','024366','024959','024960','024961','024962','024963','024964','024965','024966','024967','024968','024969','024970','024971','024972','024973','024974','024975','024976','024977','024978','024979','024980','024981','024982','024983','024984','024985','024986','024987','024988','024989','024990','024991','024992','024995','024996','024997','024998','024999','025000','025001','025002','025003','025004','025005','025006','025007','025008','025009','025010','025011','025012','025015','025016','025017','025018','024279','024280','024273','024274','024275','024276','024307','024308','024281','024282','024305','024306','024277','024278','024568','024569','024570','024571','024572','024573','024574','024575','024576','024577','024579','024578','024580','024581','024582','024583','024584','024585','024586','024587','024588','024589','024590','024591','024592','024593','024594','024595','024596','024597','023278','024531','025291','025292','025293','025294','025295','025296','025297','025298','025299','025300','025301','025302','025303','025304','025305','025306','025307','025308','025309','025310','025311','025312','025313','025314','025315','025316','025317','025318','025319','025320','025321','025322','025323','025324','025325','025326','025327','025328','025329','025330','025331','025332','025333','025334','025335','025336','025337','025338','025339','025340','025341','025342','025343','025344','025345','025346','025347','025348','025349','025350','025351','025352','025353','025354','025355','025356','025358','025359','025360','025361','025362','025363','025364','025365','025366','025367','025368','025369','025370','025371','025372','025373','025374','025375','025376','025377','025378','025379','025380','025381','025382','025383','025384','025385','025386','025387','025388','025389','025390','025391','024307','024308','024285','024286','024287','024288','024993','024994','025013','025014','025396','025397','025392','025393','025737','025738','025739','025740','025741','025742','025743','025744','025745','025746','025747','025748','025749','025750','025751','025752','025753','025754','026063','026064','026065','026066','026067','026068','026344','026345'], 
              'EDITORA JURIDICA DA BAHIA': ['020794'], 
              'olist': ['026876', '026878', '026892', '020868', '026860', '026862', '026864', '026866', '026882', '026904', '026890', '026894', '026896', '026898', '026908', '026910', '026914', '026918', '026920', '028319', '026863', '026865', '026913', '026917', '028318'],
              'onofre': ['015052', '014950', '020865','015965','015964'],
              'LESEN': ['016137', '021379'],
              'MVX': ['015908','014935'],
              'riachuelo': ['022958','025784','027985','027986','028384','028385','027886','027887','028339','028340','028341','028342','028343','028344','028345','028346','028347','028348','028607','028608','028609','028610','028611','028612','028613','028614','028674']}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    

  print("Lista de clientes: ")
  print(list(clientes.keys()))
  
  sit = 0
  while sit == 0:
    cliente_desejado = input("Favor informar o cliente: ")
    try:
      lista_contas = clientes[cliente_desejado]
      sit = 1
    except:
      print('Cliente inválido')

  return cliente_desejado,lista_contas

def adicionar_cliente():
  nome_novo_cliente =  input("Insira o nome do cliente a ser adicionado: ")
  lista_contas_novo_cliente = input("informe a lista de contas, separadas por vírgula").split(',')

def juntar_relatorios(caminho):
  arquivos = [caminho+f for f in listdir(caminho) if isfile(join(caminho, f))]

  print("-----UNINDO RELATÓRIOS-----")

  dfs = [pd.read_excel(arquivo, header=1, dtype=object) for arquivo in arquivos]

  df_demon_list = []
  df_desc_list = []
  for df in dfs:
    if "Remessa" in df.columns:
      df_demon_list.append(df)
    else:
      df_desc_list.append(df)

  df_demon = pd.concat(df_demon_list)
  df_demon = df_demon.fillna('')

  df_desc = pd.concat(df_desc_list)
  df_desc = df_desc.fillna('')

  df_desc.Pedido = df_desc.Pedido.str.replace('', ' ')

  destino_demon = "Demonstrativo.xlsx"
  destino_desc = "Descritivo.xlsx"

  #salvar demonstrativo
  df_demon.to_excel(destino_demon, index=False, startrow=1)
  #alterar celula 1
  srcfile = openpyxl.load_workbook(destino_demon, read_only=False)
  srcfile['Sheet1']['A1'] = str("Relatorio TOP")
  srcfile.save(destino_demon)
  print("DEMONSTRATIVO OK")

  #Fazendo o mesmo para o descritivo
  df_desc.to_excel(destino_desc, index=False, startrow=1)
  srcfile = openpyxl.load_workbook(destino_desc, read_only=False)
  srcfile['Sheet1']['A1'] = str("Relatorio TOP")
  srcfile.save(destino_desc)
  print("DESCRITIVO OK")

def extrair_relatorios(usuario, senha, cliente, lista_contas, quinzenas):
  '''
  Função para extrair os relatorios do sistema fraction web.
  '''
  quinzenas = quinzenas
  lista_contas = lista_contas
  
  warnings.filterwarnings("ignore",category=DeprecationWarning)


  options = webdriver.ChromeOptions()
  options.add_argument('--no-sandbox')
  options.add_argument('--disable-dev-shm-usage')
  options.add_argument("--window-size=1920,1080")
  options.add_argument("--start-maximized")
  options.add_argument('--headless')
  nome_pasta = "relatorios_"+cliente
  pasta_destino = rf"C:\Users\daniel.watanabe\Documents\Script\sistema\{nome_pasta}"
  #params = {'behavior': 'allow', 'downloadPath': r'C:/Users/daniel.watanabe/Documents/Script/sistema/'+pasta_destino}


  params = {'behavior': 'allow', 'downloadPath': rf"{pasta_destino}"}

  # open it, go to a website, and get results
  wd = webdriver.Chrome(options=options)
  wd.execute_cdp_cmd('Page.setDownloadBehavior', params)

  #INPUTS
  razao_social = cliente
  len_contas = len(lista_contas)*len(quinzenas)

  #SITE FRACTIONWEB
  wd.get("http://www.jadlog.com.br/FractionWeb/login.jad")

  #LOGIN
  usuario = wd.find_element_by_id('id_usuario').send_keys(usuario)
  senha = wd.find_element_by_id('id_senha').send_keys(senha)
  wd.find_element_by_id('botao_login').click()

  count = 0
  for quinzena in quinzenas:
    for conta in lista_contas:
      count+=1
      wd.get('http://www.jadlog.com.br/FractionWeb/pages/folhaApoio/folha.jad')
      

      #CLICA EM CORRENTISTA E/D
      WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='opt_tipo_boleto_1']/div[2]"))).click()
      wd.find_element_by_xpath("//*[@id='opt_4']").click()

      #CLICA EM ADICIONAR
      WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='btn_adicionar']"))).click()

      #Escreve a razão social
      WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='form_correntista:pessoa_nome_fantasia']"))).send_keys(razao_social)
      time.sleep(0.25)

      #CLICA PARA BUSCAR
      wd.find_element_by_xpath("//*[@id='form_correntista:btn_busca_pessoa']").click()
      time.sleep(1)


      #ENVIAR CONTA
      element = WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='form_correntista:data_table_correntista:j_idt189:filter']")))
      element.send_keys(conta)
      time.sleep(1)
      
      try:
        wd.find_element_by_xpath("//*[@id='form_correntista:data_table_correntista:0:j_idt185']").click() #CLICA NO PRIMEIRO MATCH
        time.sleep(0.5)

        #DEFINE A QUINZENA
        wd.find_element(By.ID, "j_idt252_input").send_keys(quinzena)

        #WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='tipo_correntista']/tbody/tr[1]/td[1]/div/div[2]"))).click()
        time.sleep(0.5)

        #BAIXAR Demonstrativo
        wd.find_element_by_xpath("//a[contains(@onclick,'id_exportar_excel_populado')]").click() #BAIXAR RELATORIO
        WebDriverWait(wd, 100).until(EC.invisibility_of_element_located((By.ID, "id_bloqueio_folha"))) #Espera a telinha de loading desaparecer
        #time.sleep(4)

        #Clica em descritivo
        wd.find_element_by_xpath("//*[@id='tipo_correntista']/tbody/tr[2]/td[2]/div").click() 
        time.sleep(0.5)

        #baixar descritivo
        wd.find_element_by_xpath("//a[contains(@onclick,'id_exportar_excel_populado')]").click() #BAIXAR RELATORIO
        WebDriverWait(wd, 100).until(EC.invisibility_of_element_located((By.ID, "id_bloqueio_folha"))) #Espera a telinha de loading desaparecer

        
        print(f"Conta:{conta}   Quinzena: {quinzena} {count}/{len_contas} OK")
        

      except Exception as e:
        print(f"Erro na extração da conta: {conta} e quinzena: {quinzena}.\nerro:{e}")
      

    time.sleep(3)
  wd.close()

def filtrar_relatorios(caminho_demonstrativo,caminho_descritivo,caminho_critica):

  #Definir caminhos dos arquivos a serem filtrados
  caminho_demonstrativo = caminho_demonstrativo
  caminho_descritivo = caminho_descritivo
  caminho_critica = caminho_critica

  #Abrir os arquivos e transformar em data frames
  df_demonstrativo = pd.read_excel(caminho_demonstrativo, header=1, dtype=object)
  df_descritivo = pd.read_excel(caminho_descritivo, header=1, dtype=object)
  df_critica = pd.read_excel(caminho_critica, header=1, dtype=object)

  #Filtrar tanto demonstrativo quanto descritivo
  df_demonstrativo_final = df_demonstrativo[df_demonstrativo["Remessa"].isin(df_critica["CTE"])]
  df_descritivo_final = df_descritivo[df_descritivo["Cte"].isin(df_critica["CTE"])]

  #Substituir NA por vazio
  df_demonstrativo_final = df_demonstrativo_final.fillna('')
  df_descritivo_final = df_descritivo_final.fillna('')

  #Substituir vazio por espaço do descritivo final
  df_descritivo_final.Pedido = df_descritivo_final.Pedido.str.replace('', ' ')



  #Salvar em excel
  df_demonstrativo_final.to_excel(f"Demonstrativo_filtrado.xlsx", index=False, startrow=1)
  print("-----DEMONSTRATIVO OK-----")
  df_descritivo_final.to_excel(f"Descritivo_filtrado.xlsx", index=False, startrow=1)
  print("-----DESCRITIVO OK-----")

def buscar_ctes(usuario, senha, lista_ctes, output ='ctes.xlsx'):

  options = webdriver.ChromeOptions()
  options.add_argument('--no-sandbox')
  options.add_argument('--disable-dev-shm-usage')
  options.add_argument("--window-size=1920,1080")
  options.add_argument("--start-maximized")
  options.add_experimental_option('excludeSwitches', ['enable-logging'])
  options.add_argument('--headless')
  

  wd = webdriver.Chrome(options=options)
  wd.get("http://www.jadlog.com.br/FractionWeb/login.jad")

  #LOGIN
  WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.ID, 'id_usuario'))).send_keys(usuario)
  WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.ID, 'id_senha'))).send_keys(senha)
  WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.ID, 'botao_login'))).click()

  num_ctes = str(len(lista_ctes))

  dados_final = {}
  counter = 1
  for cte in lista_ctes:
    try:
      #site para buscar CTE
      wd.get("http://www.jadlog.com.br/FractionWeb/jad/pesquisar?execution=e1s1")
      #Inserir CTE e buscar
      WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.ID, "frmPesquisa:cte"))).send_keys(cte)
      WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.ID, "frmPesquisa:id_enviar"))).click()

      #Pegar tabela Informações da Remessa
      WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='j_idt175:j_idt287_content']/table[1]/tbody/tr")))

      #Pegar linhas
      linhas = wd.find_elements(By.XPATH, "//*[@id='j_idt175:j_idt287_content']/table[1]/tbody/tr")
      num_linhas = len(linhas)

      #colunas = WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='j_idt175:j_idt287_content']/table[1]/tbody/tr/td")))
      colunas = len(wd.find_elements(By.XPATH, "//*[@id='j_idt175:j_idt287_content']/table[1]/tbody/tr[1]/td"))
      dados = {}

      

      #pegar dados de uf e cep
      dados["UF Origem"] = wd.find_element(By.XPATH, "//*[@id='j_idt175:j_idt196_content']/table/tbody/tr[3]/td[2]/label").text.split("/")[1].strip()
      dados["Cep Origem"] = wd.find_element(By.XPATH, "//*[@id='j_idt175:j_idt196_content']/table/tbody/tr[5]/td[2]/label").text
      dados["UF Destino"] = wd.find_element(By.XPATH, "//*[@id='j_idt175:j_idt215_content']/table/tbody/tr[3]/td[2]/label").text.split("/")[1].strip()[:2]
      dados["Cep Destino"] = wd.find_element(By.XPATH, "//*[@id='j_idt175:j_idt215_content']/table/tbody/tr[5]/td[2]/label").text

      #pegar dacte
      dados["Dacte"] = wd.find_element(By.XPATH, "//*[@id='j_idt175:j_idt287_content']/table[2]/tbody/tr/td[2]/label[2]").text
      dados["Tomador"] = wd.find_element(By.XPATH, "//*[@id='j_idt175']/table[1]/tbody/tr/td[2]/label[2]").text
      for l in range(1,num_linhas + 1):
          for c in range(1,colunas+1):
              try:
                  dado_lista = wd.find_element(By.XPATH, "//*[@id='j_idt175:j_idt287_content']/table[1]/tbody/tr["+str(l)+"]/td["+str(c)+"]").text.split(":")
                  dados[dado_lista[0]] = dado_lista[1].strip()
              except:
                  pass
      try:
        dados["ESTORNADO"] = wd.find_element(By.XPATH, "//*[@id='j_idt175:j_idt183_header']/span").text
      except:
        dados["ESTORNADO"] = 'NAO'
      
      #Pegar CO Emissor
      
      dados["CO Emissor"] = wd.find_element(By.XPATH, "//*[@id='j_idt175:trackHistorico_data']/tr[1]/td[4]/span").text

      WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.ID, "j_idt175:j_idt347_toggler"))).click()
      
      #Pegar tabela componentes do frete
      WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='j_idt175:trackComponeteFrete_data']/tr")))

      #Pegar linhas
      linhas = wd.find_elements(By.XPATH, "//*[@id='j_idt175:trackComponeteFrete_data']/tr")
      num_linhas = len(linhas)
      colunas = len(wd.find_elements(By.XPATH, "//*[@id='j_idt175:trackComponeteFrete_data']/tr[1]/td"))

      #Pegar todos os componentes de frete
      lista_temp = []
      time.sleep(0.5)
      for a in range(1,num_linhas + 1):
          for b in range(1,colunas+1):
              dado = wd.find_element(By.XPATH, "//*[@id='j_idt175:trackComponeteFrete_data']/tr["+str(a)+"]/td["+str(b)+"]").text
              lista_temp.append(dado)

      lista_chaves = []
      lista_valores = []
      for item in lista_temp:
          if lista_temp.index(item) % 2 == 0:
            if item not in lista_chaves:
              lista_chaves.append(item)
            else:
              lista_chaves.append(item+"_2")
          else:
              lista_valores.append(item)
              
      for item in lista_chaves:
          dados["COMPONENTES;"+item] = lista_valores[lista_chaves.index(item)]

  
      dados_final[cte] = dados
      print(cte + f" {counter}/{num_ctes}   OK")
      counter += 1
    except:
      print(cte + f" {counter}/{num_ctes}   ERRO")
      counter += 1
      pass
  #fechar pagina
  wd.close()

  #criar dataframe com dicionario
  df_completo = DataFrame.from_dict(dados_final, orient='index')

  #Alterar tipos das colunas numericas
  colunas_float = ['ADVALOREM', 'Peso', 'Peso Taxado', 'COMPONENTES;FRETE PRINCIPAL', 'COMPONENTES;TAXA ENTREGA', 'COMPONENTES;COLETA', 'COMPONENTES;GRIS', 'COMPONENTES;GRIS_2', 'COMPONENTES;TAXA EXTRACENT', 'Valor Serviço', 'Valor Merc.', 'TX_NAO_MEC']
  for coluna in colunas_float:
      try:
          df_completo[coluna] = df_completo[coluna].str.replace(',','.').astype(float)
      except:
          pass


  df_completo.index.names = ['Cte']

  df_final = df_completo.drop(['Descrição','Lista','Recebedor','Entrega'],axis='columns')
  df_final = df_final.rename(columns={'Peso': 'Peso Real'})
  df_final.to_excel(output)
  print("-----FINALIZADO-----")

def interface():
  '''
  Função de interface para download e junção de relatórios
  '''

  #Iniciar interface e printar opções
  print("----------SISTEMA EXTRAÇÃO COMERCIAL----------\n")
  print('''O que você deseja fazer:
        1- Baixar relatorios de cliente cadastrado
        2- Baixar relatório de cliente não cadastrado
        3- Unir relatórios descritivos e demonstrativos
        4- Filtrar remessas da crítica
        5- Buscar CTEs Específicos''')
  
  sit = 0
  while sit == 0:
    acao_usuario = input("Selecionar opção: ")

    if acao_usuario in ['1','2']:
      usuario = input('\nUsuário Fraction: ')
      senha = input('Senha Fraction: ')

      #Baixar relatórios de cliente cadastrado
      if acao_usuario == '1':
        sit = 1
        retorno_select = selecionar_cliente()
        cliente = retorno_select[0]
        lista_contas = retorno_select[1]
      
      #Baixar relatórios de cliente não cadastrado
      elif acao_usuario == '2':
        sit = 1
        cliente = input("\nQue cliente deseja buscar? ")
        lista_contas = input("\nFavor informar a lista de contas para extrair relatórios (separar por ',', sem dígitos): ").split(',')

      quinzenas = input("Favor informar a(s) quinzena(s) separadas por ',' no formato DD/MM/YY: ").split(",")
      print("\n-----EXTRAINDO RELATÓRIOS-----\n")
      extrair_relatorios(usuario, senha, cliente, lista_contas, quinzenas)
      print("\n-----RELATÓRIOS EXTRAÍDOS-----\n")

      juntar_ask = input("Deseja unir os relatórios baixados? [s/n]")
      if juntar_ask == 's':
        caminho = "relatorios_"+cliente+"/"
        juntar_relatorios(caminho)
        print("\nJUNÇÃO REALIZADA\n")

    #Juntar relatorios de um diretório, separando demonstrativo e descritivo
    elif acao_usuario == '3':
      sit = 1
      caminho = input("\nFavor informar o nome da pasta com os relatórios a juntar: ")
      juntar_relatorios(caminho)
    
    #filtrar apenas linhas que estão na crítica, deve informar um arquivo com a coluna Cte da critica.
    elif acao_usuario == '4':
      sit = 1
      caminho_demonstrativo = input("Favor informar o caminho para o arquivo do demonstrativo: ")
      caminho_descritivo = input("Favor informar o caminho para o arquivo do descritivo: ")
      caminho_critica = input("Favor informar o caminho para o arquivo da crítica\n(Precisa ter a coluna Cte, na linha 2): ")
      filtrar_relatorios(caminho_demonstrativo,caminho_descritivo,caminho_critica)
    
    elif acao_usuario == '5':
      sit = 1
      usuario = input('\nUsuário Fraction: ')
      senha = input('Senha Fraction: ')
      lista_ctes = input("Favor informar uma lista de CTEs, separados por vírgula: ").split(",")
      lista_ctes = [cte.strip() for cte in lista_ctes]
      #lista_ctes = ['18176300074263','18176300074195','18176300064506','18176300056964','18160100185917','18000700049680','18000700049527','18000700049501','18000700049249','18000700041045','18176300051800','18171900016707','18108300188027']
      print("-----EXTRAINDO INFORMAÇÕES DE CTES-----")
      buscar_ctes(usuario,senha,lista_ctes)

    else:
      print("OPÇÃO INVÁLIDA")

interface()