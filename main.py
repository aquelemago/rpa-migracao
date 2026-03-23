from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.support.ui import Select
import time
import os
import glob
import re
import unicodedata
import shutil
import socket
import requests
import json
import traceback
import zipfile

def carregar_config():
    if not os.path.exists("configs.json"):
        print("Arquivo configs.json nao encontrado!")
        exit(1)
    with open("configs.json", "r", encoding="utf-8") as f:
        return json.load(f)

CONFIG = carregar_config()
DRIVE = CONFIG.get("drive_letter", "C").upper().strip().replace(":", "")
PASTA_DOWNLOAD = f"{DRIVE}:\\"

# ==========================
# CONFIGURACOES FIXAS
# ==========================
TIMEOUT_PADRAO = 30
TIMEOUT_DOWNLOAD = 600
MAX_TENTATIVAS = 5
PAUSA_ENTRE_ACOES = 2
PAUSA_APOS_REFRESH = 3
ARQUIVO_CHECKPOINT = "checkpoint.json"

ambientes = CONFIG.get("ambientes", {})
ambiente_escolhido = None
for nome, dados in ambientes.items():
    if dados.get("ativo") == True:
        ambiente_escolhido = dados
        print(f"Ambiente ativo: {nome}")
        break

if ambiente_escolhido is None:
    print("Nenhum ambiente ativo no configs.json.")
    exit(1)

URL_LOGIN = ambiente_escolhido["url_login"]
URL_LISTAGEM = ambiente_escolhido["url_listagem"]
LOGIN = ambiente_escolhido["credenciais"]["login"]
SENHA = ambiente_escolhido["credenciais"]["senha"]
UNIDADES_ALVO = [str(u) for u in ambiente_escolhido.get("unidades", [])]

def carregar_checkpoint():
    if os.path.exists(ARQUIVO_CHECKPOINT):
        with open(ARQUIVO_CHECKPOINT, "r", encoding="utf-8") as f:
            try:
                dados = json.load(f)
                if isinstance(dados, list):
                    return {str(k): {"pagina": 9999, "acumulado": 0} for k in dados}
                return dados
            except: return {}
    return {}

def salvar_checkpoint(processadas):
    with open(ARQUIVO_CHECKPOINT, "w", encoding="utf-8") as f:
        json.dump(processadas, f, indent=2)

def normalizar_nome_unidade(nome):
    if not nome or nome.strip() == ";;": return None
    nome = unicodedata.normalize('NFKD', nome).encode('ASCII', 'ignore').decode('ASCII')
    nome = re.sub(r'[\\/*?:"<>|]', "", nome)
    return nome.strip()

def contar_pdfs_nos_zips(pasta):
    if not os.path.exists(pasta): return 0
    total = 0
    for f in os.listdir(pasta):
        if f.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(os.path.join(pasta, f), 'r') as z:
                    total += len([name for name in z.namelist() if name.lower().endswith('.pdf')])
            except: pass
    return total

def aguardar_download(pasta, arquivos_antes, timeout=600):
    inicio = time.time()
    print(f"Aguardando arquivo ZIP aparecer em {pasta}...")
    
    while time.time() - inicio < timeout:
        arquivos_agora = set(glob.glob(os.path.join(pasta, "*.zip")))
        novos = arquivos_agora - arquivos_antes
        
        if novos:
            # Encontrou um novo ZIP. O Chrome so cria o .zip final apos concluir o .crdownload
            candidato = list(novos)[0]
            print(f"Detectado novo arquivo: {os.path.basename(candidato)}. Verificando integridade...")
            
            try:
                # Pequena espera para garantir que o SO liberou o arquivo apos o renomeio do Chrome
                time.sleep(1)
                tamanho_ini = os.path.getsize(candidato)
                if tamanho_ini > 0:
                    print(f"Download concluido com sucesso: {os.path.basename(candidato)}")
                    return candidato
            except Exception as e:
                print(f"Aguardando estabilizacao do arquivo... ({e})")
        
        time.sleep(1)
    
    raise TimeoutException("Tempo excedido aguardando o surgimento do arquivo ZIP no disco.")


def aguardar_carregamento(driver):
    try:
        wait = WebDriverWait(driver, 20)
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "loading-overlay")))
        wait.until(EC.invisibility_of_element_located((By.ID, "loading-bar")))
        time.sleep(2)
    except: pass

def obter_info_paginacao(driver):
    try:
        wait = WebDriverWait(driver, 15)
        info_element = wait.until(lambda d: d.find_element(By.ID, "DataTables_Table_0_info") if d.find_element(By.ID, "DataTables_Table_0_info").text.strip() else False)
        info = info_element.text
        match = re.search(r'(\d+)\s+até\s+([\d.,]+)\s+de\s+([\d.,]+)', info)
        if match:
            de = int(match.group(1).replace('.', '').replace(',', ''))
            ate = int(match.group(2).replace('.', '').replace(',', ''))
            total = int(match.group(3).replace('.', '').replace(',', ''))
            return de, ate, total
    except: pass
    return None, None, None

def verificar_tabela_vazia(driver):
    aguardar_carregamento(driver)
    try:
        seletores_vazio = ["//td[@class='dataTables_empty']", "//*[contains(text(), 'Não foram encontrados resultados')]"]
        for sel in seletores_vazio:
            if driver.find_elements(By.XPATH, sel): return True
        return False
    except: return False

def executar_com_retry(driver, funcao, *args, **kwargs):
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            return funcao(*args, **kwargs)
        except Exception as e:
            print(f"Tentativa {tentativa}/{MAX_TENTATIVAS} falhou: {e}")
            if tentativa == MAX_TENTATIVAS: raise
            try:
                driver.refresh()
                time.sleep(PAUSA_APOS_REFRESH)
                if "/login" in driver.current_url: fazer_login(driver)
                navegar_para_listagem(driver)
                reabrir_busca_avancada_e_modal(driver)
            except: pass

def fazer_login(driver):
    print("Iniciando Login...")
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    driver.get(URL_LOGIN)
    try: wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Entendi')]"))).click()
    except: pass
    wait.until(EC.visibility_of_element_located((By.ID, "login"))).send_keys(LOGIN)
    wait.until(EC.visibility_of_element_located((By.ID, "inputPassword1"))).send_keys(SENHA)
    driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.XPATH, "//button[@type='submit']"))))
    wait.until(EC.url_contains("/lms/#/portal"))
    time.sleep(3)

def navegar_para_listagem(driver):
    print("Navegando para listagem de certificados...")
    for tentativa in range(3):
        driver.get(URL_LISTAGEM)
        time.sleep(5)
        
        # Se cair no dashboard por erro de sessao ou redirecionamento
        if "dashboard" in driver.current_url:
            for s in ["//a[contains(text(),'Admin')]", "//span[contains(text(),'Admin')]"]:
                try: WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, s))).click(); time.sleep(2); break
                except: continue
            for s in ["//a[contains(text(),'Usuários Certificados')]", "//a[contains(@href,'certificado')]"]:
                try: WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, s))).click(); time.sleep(3); break
                except: continue
        
        # Validacao: Verificamos se um elemento essencial da pagina carregou (tabela ou botao de busca)
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//table | //button[@idaut='botao_busca_avancada']"))
            )
            print("Pagina de listagem carregada com sucesso.")
            return
        except TimeoutException:
            print(f"Pagina parece em branco (Tentativa {tentativa+1}/3). Atualizando...")
            driver.refresh()
            time.sleep(5)
    
    raise Exception("Falha ao carregar a pagina de listagem apos varias tentativas.")

def expandir_arvore(driver):
    print("Expandindo arvore organizacional...")
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    wait.until(EC.presence_of_element_located((By.XPATH, "//i[contains(@class,'fa-plus-square-o')]")))
    while True:
        icones = driver.find_elements(By.XPATH, "//i[contains(@class,'fa-plus-square-o')]")
        if not icones: break
        for icone in icones:
            try:
                el = icone.find_element(By.XPATH, "./ancestor::*[self::a or self::span or self::li][1]")
                driver.execute_script("arguments[0].click();", el)
                time.sleep(0.3)
            except: continue
        time.sleep(1)

def obter_caminho_hierarquico(driver, checkbox):
    caminho = []
    try:
        # Localiza o <li> que contem este checkbox
        li_atual = checkbox.find_element(By.XPATH, "./ancestor::li[1]")
        while li_atual:
            nome = None
            # Seletores possiveis para o texto da unidade
            for sel in [".//a[starts-with(@id, 'link_child_')]", ".//span[@ng-if='!hasChildren(unidade)']", ".//span[contains(@class, 'ng-binding')]"]:
                try:
                    el = li_atual.find_element(By.XPATH, sel)
                    if el.is_displayed() and el.text.strip():
                        nome = el.text.strip()
                        break
                except: continue
            
            if nome:
                norm = normalizar_nome_unidade(nome)
                if norm: caminho.insert(0, norm)
            
            # Sobe para o <li> pai (ancestor imediato)
            try:
                li_atual = li_atual.find_element(By.XPATH, "ancestor::li[1]")
                # Limite de seguranca para evitar loops infinitos em arvores mal formadas
                if len(caminho) > 20: break 
            except:
                li_atual = None
                break
    except: pass
    return caminho if caminho else ["Unidade_Desconhecida"]

def obter_todos_ids_descendentes(driver, checkbox_raiz):
    li_raiz = checkbox_raiz.find_element(By.XPATH, "./ancestor::li[1]")
    checkboxes = li_raiz.find_elements(By.XPATH, ".//input[starts-with(@id, 'input_selected_unit_')]")
    ids = []
    for cb in checkboxes:
        match = re.search(r'_(\d+)', cb.get_attribute("id"))
        if match: ids.append(match.group(1))
    return ids

def selecionar_unidade_e_buscar(driver, id_unidade):
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.ID, f"input_selected_unit_{id_unidade}"))) )
    time.sleep(1)
    driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.ID, "btn_select_unit"))) )
    time.sleep(1)
    driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Buscar']"))) )
    time.sleep(3)

def definir_quantidade_por_pagina(driver):
    print("Configurando exibicao de 50 registros por pagina...")
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    try:
        select_qtd = wait.until(EC.presence_of_element_located((By.XPATH, "//select[option[@value='50']]")))
        Select(select_qtd).select_by_value("50")
        time.sleep(2)
        aguardar_carregamento(driver)
    except: pass

def processar_paginas_da_unidade(driver, nome_unidade, pasta_destino, id_unidade, processadas_global):
    pagina_atual = 1
    cp = processadas_global.get(str(id_unidade), {"pagina": 0, "acumulado": 0})
    ultima_pag_concluida = cp.get("pagina", 0)
    total_confirmado = cp.get("acumulado", 0)
    total_esperado_final = None

    while True:
        aguardar_carregamento(driver)
        de, ate, total = obter_info_paginacao(driver)
        is_last_page = (ate == total) if ate and total else False
        
        if total_esperado_final is None and total is not None:
            total_esperado_final = total
            print(f"Meta Final da Unidade {id_unidade}: {total_esperado_final} registros.")
        
        if verificar_tabela_vazia(driver): break

        if pagina_atual <= ultima_pag_concluida:
            print(f"Pagina {pagina_atual} ja auditada. Saltando...")
            total_confirmado = ate
            if not avancar_pagina(driver): break
            pagina_atual += 1
            continue

        esperado_nesta_p = ate - de + 1
        print(f"Processando Pagina {pagina_atual} (Meta: {esperado_nesta_p} registros).")

        sucesso_pagina = False
        for tentativa_pag in range(1, 4):
            total_antes = contar_pdfs_nos_zips(pasta_destino)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            
            checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox' and @ng-model='check.value']")
            if len(checkboxes) < esperado_nesta_p:
                time.sleep(4)
                checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox' and @ng-model='check.value']")

            selecionados = 0
            for cb in checkboxes:
                try:
                    if not cb.is_selected(): driver.execute_script("arguments[0].click();", cb)
                    selecionados += 1
                except: pass
            
            print(f"Selecionados {selecionados} registros para baixar.")
            status = baixar_zip_unidade(driver, pagina_atual, pasta_destino, is_last_page)
            
            if status == "RETRY":
                print(f"Reiniciando selecao na pagina {pagina_atual}...")
                for cb in checkboxes:
                    try:
                        if cb.is_selected(): driver.execute_script("arguments[0].click();", cb)
                    except: pass
                time.sleep(PAUSA_ENTRE_ACOES)
                continue

            total_depois = contar_pdfs_nos_zips(pasta_destino)
            obtidos = total_depois - total_antes
            
            if obtidos >= selecionados:
                print(f"Pagina {pagina_atual} confirmada: {obtidos} arquivos salvos.")
                total_confirmado += obtidos
                sucesso_pagina = True
                for cb in checkboxes:
                    try:
                        if cb.is_selected(): driver.execute_script("arguments[0].click();", cb)
                    except: pass
                break
            else:
                print(f"Erro na Pagina {pagina_atual}: Esperava {selecionados}, obteve {obtidos}. Tentando novamente...")
                for cb in checkboxes:
                    try:
                        if cb.is_selected(): driver.execute_script("arguments[0].click();", cb)
                    except: pass
                time.sleep(5)

        if not sucesso_pagina: raise Exception(f"Falha na integridade da Pagina {pagina_atual}.")

        processadas_global[str(id_unidade)] = {"pagina": pagina_atual, "acumulado": total_confirmado}
        salvar_checkpoint(processadas_global)
        
        if not avancar_pagina(driver): break
        pagina_atual += 1

    total_fisico = contar_pdfs_nos_zips(pasta_destino)
    if total_esperado_final is not None and total_fisico < total_esperado_final:
        print(f"Divergencia Final: HD {total_fisico}, Sistema {total_esperado_final}.")
        raise Exception("Unidade incompleta.")
    
    print(f"Unidade {id_unidade} concluida: {total_fisico} arquivos.")
    processadas_global[str(id_unidade)] = {"pagina": 9999, "acumulado": total_fisico}
    salvar_checkpoint(processadas_global)

def avancar_pagina(driver):
    info_ant = ""
    try: info_ant = driver.find_element(By.ID, "DataTables_Table_0_info").text
    except: pass
    seletores = ["//li[contains(@class,'next') and not(contains(@class,'disabled'))]/a", "//*[@id='DataTables_Table_0_paginate']//li[@class='next']/a"]
    btn = None
    for s in seletores:
        try:
            el = driver.find_elements(By.XPATH, s)
            if el and el[0].is_displayed(): btn = el[0]; break
        except: continue
    if btn:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            time.sleep(1); driver.execute_script("arguments[0].click();", btn)
            if info_ant:
                try: WebDriverWait(driver, 20).until(lambda d: d.find_element(By.ID, "DataTables_Table_0_info").text != info_ant)
                except: time.sleep(5)
            aguardar_carregamento(driver); return True
        except: return False
    return False

def baixar_zip_unidade(driver, pagina_atual, pasta_destino, is_last_page):
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    
    # 1. Clicar em 'Baixar Certificados' para abrir o modal
    print("Abrindo modal de processamento...")
    btn_gerar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Baixar Certificados')]")))
    driver.execute_script("arguments[0].click();", btn_gerar)
    
    # VERIFICACAO SOLICITADA: Total < 50 e nao e ultima pagina
    try:
        # Reduzimos o tempo de espera para a verificação ser rápida (evita o "demorando")
        wait_curto = WebDriverWait(driver, 5)
        total_element = wait_curto.until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'TOTAL PROCESSADO:')]")))
        texto_total = total_element.text
        match = re.search(r'de\s+(\d+)', texto_total)
        if match:
            total_modal = int(match.group(1))
            print(f"Total detectado no modal: {total_modal}")
            if total_modal < 50 and not is_last_page:
                print(f"Aviso: Total {total_modal} < 50 e nao e a ultima pagina. Fechando e tentando novamente...")
                try:
                    # Se abriu em nova aba (conforme o prompt sugere "fechar aba")
                    if len(driver.window_handles) > 1:
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                    else:
                        # Se for modal (conforme 1.html), clicamos em cancelar
                        btn_cancelar = driver.find_element(By.XPATH, "//button[contains(text(),'Cancelar')]")
                        driver.execute_script("arguments[0].click();", btn_cancelar)
                        WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".modal-backdrop, .modal")))
                except: pass
                return "RETRY"
    except Exception as e:
        print(f"Modal detectado, mas info de total ainda nao visivel ou erro: {e}. Prosseguindo com clique no Iniciar...")

    # 2. Clicar em 'Iniciar' no modal (Acao solicitada: "ele deve clicar no botao iniciar")
    print("Clicando no botao Iniciar...")
    btn_confirmar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Iniciar')]")))
    driver.execute_script("arguments[0].click();", btn_confirmar)
    
    # 3. Aguardar o processamento e o surgimento do botao 'Download do ZIP'
    print("Processando certificados... Aguardando conclusao.")
    btn_download = None
    inicio_proc = time.time()
    while time.time() - inicio_proc < TIMEOUT_DOWNLOAD:
        try:
            # Tenta encontrar o botao pelo texto exato conforme o HTML fornecido
            botoes = driver.find_elements(By.XPATH, "//button[contains(text(),'Download do ZIP')]")
            if botoes and botoes[0].is_displayed():
                btn_download = botoes[0]
                break
        except: pass
        time.sleep(2)
    
    if not btn_download:
        raise TimeoutException("O botao 'Download do ZIP' nao apareceu no tempo limite.")

    # 4. Iniciar o download
    print("Botao de download encontrado! Iniciando transferencia...")
    antes = set(glob.glob(os.path.join(PASTA_DOWNLOAD, "*.zip")))
    driver.execute_script("arguments[0].click();", btn_download)
    
    # 5. Aguardar o arquivo no disco (agora ja configurado para baixar direto no drive)
    arquivo = aguardar_download(PASTA_DOWNLOAD, antes)
    
    # 6. Mover/Renomear para a pasta da unidade
    os.makedirs(pasta_destino, exist_ok=True)
    nome_zip = f"pagina {pagina_atual}.zip"
    destino_final = os.path.join(pasta_destino, nome_zip)
    
    # Se ja existir (devido a retry), remove o antigo
    if os.path.exists(destino_final): os.remove(destino_final)
    shutil.move(arquivo, destino_final)
    
    # 7. Fechar o modal
    print("Limpando modal e prosseguindo...")
    try:
        btn_cancelar = driver.find_element(By.XPATH, "//button[contains(text(),'Cancelar')]")
        driver.execute_script("arguments[0].click();", btn_cancelar)
        # Aguarda o modal sumir para nao atrapalhar o proximo clique
        WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".modal-backdrop, .modal")))
    except: pass
    time.sleep(1)

def reabrir_busca_avancada_e_modal(driver):
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
    driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[idaut='botao_busca_avancada']"))))
    time.sleep(2)
    wait.until(EC.element_to_be_clickable((By.ID, "search_unidades"))).click()
    time.sleep(2)
    expandir_arvore(driver)

def main():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--incognito")
    # Desabilita o "Download Bubble" que força o "Salvar como" no modo incognito em versoes novas do Chrome
    options.add_argument("--disable-features=DownloadBubble,DownloadBubbleV2")
    
    # Configura o Chrome para baixar diretamente na pasta do drive alvo
    # Garante que o caminho nao termine com barra invertida para evitar erros de escape no Chrome
    pasta_ajustada = PASTA_DOWNLOAD.rstrip("\\")
    prefs = {
        "profile.default_content_setting_values.automatic_downloads": 1,
        "download.default_directory": pasta_ajustada,
        "savefile.default_directory": pasta_ajustada,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    # --- FIX PARA DOWNLOAD EM MODO INCOGNITO (CHROME 117+) ---
    # O comando CDP Page.setDownloadBehavior força o browser a permitir downloads sem prompt "Salvar como"
    # Mesmo em janelas anonimas/incognito.
    params = {
        "behavior": "allow",
        "downloadPath": os.path.abspath(PASTA_DOWNLOAD)
    }
    driver.execute_cdp_cmd("Page.setDownloadBehavior", params)
    
    processadas_global = carregar_checkpoint()

    try:
        fazer_login(driver)
        navegar_para_listagem(driver)
        reabrir_busca_avancada_e_modal(driver)

        if not UNIDADES_ALVO or "XXXX" in UNIDADES_ALVO:
            elementos = driver.find_elements(By.XPATH, "//input[starts-with(@id, 'input_selected_unit_')]")
            ids_alvo = [re.search(r'_(\d+)', el.get_attribute("id")).group(1) for el in elementos if re.search(r'_(\d+)', el.get_attribute("id"))]
        else: ids_alvo = UNIDADES_ALVO
        
        # Pasta raiz onde os resultados serao organizados
        pasta_raiz = os.path.join(PASTA_DOWNLOAD, "resultado do bot")
        os.makedirs(pasta_raiz, exist_ok=True)

        for id_raiz in ids_alvo:
            try:
                cb_raiz = WebDriverWait(driver, TIMEOUT_PADRAO).until(EC.presence_of_element_located((By.ID, f"input_selected_unit_{id_raiz}")))
                ids_desc = obter_todos_ids_descendentes(driver, cb_raiz)
            except: continue

            for id_u in ids_desc:
                if processadas_global.get(str(id_u), {}).get("pagina") == 9999: continue
                print(f"\n--- Iniciando Unidade: {id_u} ---")
                try:
                    def rodar():
                        nonlocal driver, id_u, processadas_global, pasta_raiz
                        cb = WebDriverWait(driver, TIMEOUT_PADRAO).until(EC.presence_of_element_located((By.ID, f"input_selected_unit_{id_u}")))
                        li = cb.find_element(By.XPATH, "./ancestor::li[1]")
                        if li.find_elements(By.XPATH, ".//i[contains(@class,'fa-plus-square-o')]") or li.find_elements(By.XPATH, ".//a[starts-with(@id, 'link_child_')]"):
                            processadas_global[str(id_u)] = {"pagina": 9999, "acumulado": 0}
                            salvar_checkpoint(processadas_global); return

                        caminho = obter_caminho_hierarquico(driver, cb)
                        pasta_u = os.path.join(pasta_raiz, *caminho)
                        print(f"Caminho: {' > '.join(caminho)}")
                        
                        selecionar_unidade_e_buscar(driver, id_u)
                        if not verificar_tabela_vazia(driver):
                            definir_quantidade_por_pagina(driver)
                            processar_paginas_da_unidade(driver, caminho[-1], pasta_u, id_u, processadas_global)
                        else:
                            processadas_global[str(id_u)] = {"pagina": 9999, "acumulado": 0}
                            salvar_checkpoint(processadas_global)

                    executar_com_retry(driver, rodar)
                    driver.refresh(); navegar_para_listagem(driver); reabrir_busca_avancada_e_modal(driver)
                except Exception as e:
                    print(f"Erro Unidade {id_u}: {e}")
                    driver.refresh(); navegar_para_listagem(driver); reabrir_busca_avancada_e_modal(driver)

        if os.path.exists(pasta_raiz) and os.listdir(pasta_raiz):
            shutil.make_archive(os.path.join(PASTA_DOWNLOAD, "Resultado_Final_Bot"), 'zip', pasta_raiz)
            print("\nPROCESSO CONCLUIDO!")

    except Exception:
        traceback.print_exc()
        time.sleep(60)
    finally:
        driver.quit()

if __name__ == "__main__": main()