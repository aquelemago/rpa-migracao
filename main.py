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
import json
import traceback
import zipfile

def carregar_config():
    if not os.path.exists("configs.json"):
        print("[ERRO] Arquivo configs.json nao encontrado!")
        exit(1)
    with open("configs.json", "r", encoding="utf-8") as f:
        return json.load(f)

CONFIG = carregar_config()
DRIVE = CONFIG.get("drive_letter", "C").upper().strip().replace(":", "")
PASTA_DOWNLOAD = os.path.join(f"{DRIVE}:\\", "RPA_MIGRACAO_TEMP")

# ==========================
# CONFIGURACOES FIXAS
# ==========================
TIMEOUT_PADRAO = 30
TIMEOUT_DOWNLOAD = 600
PAUSA_APOS_REFRESH = 3
ARQUIVO_CHECKPOINT = "checkpoint.json"

# Garante que a pasta temporária de downloads existe e está limpa
os.makedirs(PASTA_DOWNLOAD, exist_ok=True)

ambientes = CONFIG.get("ambientes", {})
ambiente_escolhido = None
for nome, dados in ambientes.items():
    if dados.get("ativo") == True:
        ambiente_escolhido = dados
        print(f"[LOG] Ambiente ativo selecionado: {nome}")
        break

if ambiente_escolhido is None:
    print("[ERRO] Nenhum ambiente ativo encontrado no configs.json.")
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
        caminho_f = os.path.join(pasta, f)
        if f.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(caminho_f, 'r') as z:
                    total += len([name for name in z.namelist() if name.lower().endswith('.pdf')])
            except: pass
        elif f.lower().endswith('.pdf'):
            total += 1
    return total

def aguardar_download(pasta, arquivos_antes, timeout=600):
    inicio = time.time()
    print(f"[LOG] Monitorando novos arquivos em: {pasta}")
    
    ultimo_log = time.time()
    while time.time() - inicio < timeout:
        # Pega todos os arquivos na pasta, exceto temporários do Chrome
        arquivos_agora = set([os.path.join(pasta, f) for f in os.listdir(pasta) 
                             if not f.endswith(".crdownload") and not f.endswith(".tmp")])
        
        novos = arquivos_agora - arquivos_antes
        
        if novos:
            # Encontrou algo novo que não é temporário
            candidato = list(novos)[0]
            ext = os.path.splitext(candidato)[1].lower()
            print(f"[LOG] Novo arquivo detectado: {os.path.basename(candidato)} (Extensão: {ext})")
            
            # Aguarda o arquivo estabilizar (tamanho parar de crescer e não estar travado)
            tentativas_estabilizacao = 0
            while tentativas_estabilizacao < 10:
                try:
                    tamanho_1 = os.path.getsize(candidato)
                    time.sleep(2)
                    tamanho_2 = os.path.getsize(candidato)
                    if tamanho_1 == tamanho_2 and tamanho_1 > 0:
                        print(f"[OK] Arquivo estabilizado: {os.path.basename(candidato)} ({tamanho_1} bytes)")
                        return candidato
                except Exception as e:
                    print(f"[AVISO] Aguardando liberação do arquivo pelo SO... ({e})")
                
                tentativas_estabilizacao += 1
                time.sleep(1)
        
        # Log de progresso
        if time.time() - ultimo_log > 15:
            crdownloads = [f for f in os.listdir(pasta) if f.endswith(".crdownload")]
            if crdownloads:
                print(f"[STATUS] Download em curso: {crdownloads[0]}")
            else:
                print(f"[STATUS] Aguardando início do download... (Tempo decorrido: {int(time.time() - inicio)}s)")
            ultimo_log = time.time()

        time.sleep(1)
    
    raise TimeoutException(f"Tempo excedido ({timeout}s) aguardando download em {pasta}")


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
        # Regex corrigido para aceitar separadores (. ou ,) em todos os números
        match = re.search(r'([\d.,]+)\s+até\s+([\d.,]+)\s+de\s+([\d.,]+)', info)
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
    tentativa = 1
    while True:
        try:
            return funcao(*args, **kwargs)
        except Exception as e:
            print(f"[AVISO] Tentativa {tentativa} falhou: {e}")
            tentativa += 1
            try:
                driver.refresh()
                time.sleep(PAUSA_APOS_REFRESH)
                if "/login" in driver.current_url: fazer_login(driver)
                navegar_para_listagem(driver)
                reabrir_busca_avancada_e_modal(driver)
            except: pass

def fazer_login(driver):
    print("[LOG] Iniciando processo de login...")
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
    print("[LOG] Navegando para a listagem de certificados...")
    while True:
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
            print("[LOG] Pagina de listagem carregada com sucesso.")
            return
        except TimeoutException:
            print("[LOG] Pagina em branco ou nao carregada. Atualizando...")
            driver.refresh()
            time.sleep(5)

def expandir_arvore(driver):
    print("[LOG] Expandindo arvore organizacional...")
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
    print(f"[LOG] Selecionando unidade {id_unidade} e disparando busca...")
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.ID, f"input_selected_unit_{id_unidade}"))) )
    time.sleep(1)
    driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.ID, "btn_select_unit"))) )
    time.sleep(1)
    driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Buscar']"))) )
    time.sleep(3)

def definir_quantidade_por_pagina(driver):
    print("[LOG] Configurando exibicao para 10 registros por pagina...")
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    try:
        select_qtd = wait.until(EC.presence_of_element_located((By.XPATH, "//select[option[@value='10']]")))
        Select(select_qtd).select_by_value("10")
        time.sleep(2)
        aguardar_carregamento(driver)
        
        # GARANTIA: Clicar em Buscar para o portal processar o limite de 10
        print("[LOG] Refazendo busca para confirmar limite de 10 registros...")
        btn_busca = driver.find_element(By.XPATH, "//button[@title='Buscar']")
        driver.execute_script("arguments[0].click();", btn_busca)
        time.sleep(4)
        aguardar_carregamento(driver)
    except Exception as e:
        print(f"[AVISO] Nao conseguiu definir 10 por pagina: {e}")

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
            print(f"[STATUS] Meta Final da Unidade {id_unidade}: {total_esperado_final} registros.")
        
        if verificar_tabela_vazia(driver): break

        if pagina_atual <= ultima_pag_concluida:
            print(f"[INFO] Pagina {pagina_atual} ja auditada. Avancando...")
            total_confirmado = ate
            if not avancar_pagina(driver): break
            pagina_atual += 1
            continue

        esperado_nesta_p = ate - de + 1
        
        # VALIDAÇÃO DE REGRA DE NEGÓCIO: Páginas intermediárias DEVEM ter 10 registros.
        # Exceção: Primeira e Última página.
        if (pagina_atual > 1 and not is_last_page) and esperado_nesta_p < 10:
            print(f"[ALERTA] Pagina intermediária {pagina_atual} indica apenas {esperado_nesta_p} registros.")
            print("[ACAO] Forcando refresh total para corrigir carregamento incompleto...")
            driver.refresh()
            time.sleep(PAUSA_APOS_REFRESH)
            navegar_para_listagem(driver)
            reabrir_busca_avancada_e_modal(driver)
            selecionar_unidade_e_buscar(driver, id_unidade)
            definir_quantidade_por_pagina(driver)
            for _ in range(pagina_atual - 1): avancar_pagina(driver)
            continue

        print(f"[LOG] Processando Pagina {pagina_atual} (Meta: {esperado_nesta_p} registros).")

        tentativa_pag = 1
        while True:
            total_antes = contar_pdfs_nos_zips(pasta_destino)
            
                # Workflow: Selecionar Um a Um
            try:
                print(f"[LOG] Selecionando checkboxes um a um na pagina {pagina_atual}...")
                driver.execute_script("window.scrollTo(0, 0);")
                time.sleep(1)
                
                # Localiza os checkboxes dentro da ibox-content conforme solicitado
                xpath_checkboxes = "//div[@class='ibox-content']//table[contains(@id, 'DataTables_Table')]//tbody//input[@type='checkbox']"
                checkboxes = driver.find_elements(By.XPATH, xpath_checkboxes)
                
                for cb in checkboxes:
                    try:
                        if not cb.is_selected():
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cb)
                            driver.execute_script("arguments[0].click();", cb)
                            time.sleep(0.5) # Delay para seleção lenta
                    except: continue
                
                time.sleep(2)
                # Verificacao se selecionou o esperado (checando os checkboxes do corpo da tabela)
                checkboxes = driver.find_elements(By.XPATH, xpath_checkboxes)
                selecionados = len([cb for cb in checkboxes if cb.is_selected()])
                
                if selecionados < esperado_nesta_p:
                    print(f"[DEBUG] Selecionados {selecionados} de {esperado_nesta_p}. Tentando scroll para carregar DOM...")
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2)
                    checkboxes = driver.find_elements(By.XPATH, "//table[contains(@id, 'DataTables_Table')]//tbody//input[@type='checkbox']")
                    selecionados = len([cb for cb in checkboxes if cb.is_selected()])

                if selecionados < esperado_nesta_p:
                    print(f"[AVISO] Divergencia na selecao ({selecionados}/{esperado_nesta_p}). Reiniciando workflow da pagina...")
                    driver.refresh()
                    time.sleep(PAUSA_APOS_REFRESH)
                    navegar_para_listagem(driver)
                    reabrir_busca_avancada_e_modal(driver)
                    selecionar_unidade_e_buscar(driver, id_unidade)
                    definir_quantidade_por_pagina(driver)
                    for _ in range(pagina_atual - 1): avancar_pagina(driver)
                    tentativa_pag += 1
                    continue

                print(f"[LOG] {selecionados} registros selecionados. Iniciando download...")
                status = baixar_zip_unidade(driver, pagina_atual, pasta_destino, is_last_page, selecionados)
                
                if status == "CORRIGIR_SELECAO":
                    print("[ACAO] Iniciando protocolo avançado de correção (Dança das Páginas)...")
                    # 1. Volta uma página (se possível)
                    if voltar_pagina(driver):
                        # 2. Seleciona um a um na página anterior
                        try:
                            xpath_cb_prev = "//div[@class='ibox-content']//table[contains(@id, 'DataTables_Table')]//tbody//input[@type='checkbox']"
                            cbs_prev = driver.find_elements(By.XPATH, xpath_cb_prev)
                            for cb in cbs_prev:
                                try:
                                    if not cb.is_selected(): driver.execute_script("arguments[0].click();", cb)
                                except: pass
                            time.sleep(1)
                        except: pass
                        
                        # 3. Abre e apenas fecha o modal de processamento na página anterior
                        try:
                            btn_gerar_v = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Baixar Certificados')]")))
                            driver.execute_script("arguments[0].click();", btn_gerar_v)
                            time.sleep(3) # Tempo para o portal "sentir" a seleção
                            btn_cancelar_v = driver.find_element(By.XPATH, "//button[contains(text(),'Cancelar')]")
                            driver.execute_script("arguments[0].click();", btn_cancelar_v)
                            WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".modal-backdrop, .modal")))
                            time.sleep(1)
                        except: pass
                        
                        # 4. Limpa a seleção da página anterior individualmente
                        try:
                            cbs_prev = driver.find_elements(By.XPATH, "//div[@class='ibox-content']//table[contains(@id, 'DataTables_Table')]//tbody//input[@type='checkbox']")
                            for cb in cbs_prev:
                                try:
                                    if cb.is_selected(): driver.execute_script("arguments[0].click();", cb)
                                except: pass
                            time.sleep(1)
                        except: pass
                        
                        # 5. Volta para a página atual
                        avancar_pagina(driver)
                    
                    # 6. Limpa a seleção da página atual individualmente (se houver resquício)
                    try:
                        cbs_curr = driver.find_elements(By.XPATH, "//div[@class='ibox-content']//table[contains(@id, 'DataTables_Table')]//tbody//input[@type='checkbox']")
                        for cb in cbs_curr:
                            try:
                                if cb.is_selected(): driver.execute_script("arguments[0].click();", cb)
                            except: pass
                        time.sleep(1)
                    except: pass
                    
                    # O loop 'while True' da tentativa_pag subirá e fará a seleção limpa da página atual
                    tentativa_pag += 1
                    continue

                if status == "REFRESH_RETRY":
                    print("[ACAO] Refresh solicitado pelo fluxo de download.")
                    driver.refresh()
                    time.sleep(PAUSA_APOS_REFRESH)
                    navegar_para_listagem(driver)
                    reabrir_busca_avancada_e_modal(driver)
                    selecionar_unidade_e_buscar(driver, id_unidade)
                    definir_quantidade_por_pagina(driver)
                    for _ in range(pagina_atual - 1): avancar_pagina(driver)
                    tentativa_pag += 1
                    continue

                if status == "RETRY":
                    print("[LOG] Falha detectada no download ou auditoria. Forçando REFRESH para nova tentativa...")
                    driver.refresh()
                    time.sleep(PAUSA_APOS_REFRESH)
                    navegar_para_listagem(driver)
                    reabrir_busca_avancada_e_modal(driver)
                    selecionar_unidade_e_buscar(driver, id_unidade)
                    definir_quantidade_por_pagina(driver)
                    for _ in range(pagina_atual - 1): avancar_pagina(driver)
                    tentativa_pag += 1
                    continue

                total_depois = contar_pdfs_nos_zips(pasta_destino)
                obtidos = total_depois - total_antes
                
                if obtidos >= selecionados:
                    print(f"[SUCESSO] Pagina {pagina_atual} confirmada: {obtidos} arquivos salvos.")
                    total_confirmado += obtidos
                    # Desmarca tudo para a proxima pagina ou unidade individualmente
                    for cb in checkboxes:
                        try:
                            if cb.is_selected(): driver.execute_script("arguments[0].click();", cb)
                        except: pass
                    break
                else:
                    print(f"[ERRO] Integridade falhou na Pagina {pagina_atual}: Esperava {selecionados}, obteve {obtidos}.")
                    tentativa_pag += 1
                    # Refresh e tenta de novo
                    driver.refresh()
                    time.sleep(PAUSA_APOS_REFRESH)
                    navegar_para_listagem(driver)
                    reabrir_busca_avancada_e_modal(driver)
                    selecionar_unidade_e_buscar(driver, id_unidade)
                    definir_quantidade_por_pagina(driver)
                    for _ in range(pagina_atual - 1): avancar_pagina(driver)
                    continue
            except Exception as e:
                print(f"[ERRO] Falha no workflow da pagina {pagina_atual}: {e}")
                driver.refresh()
                time.sleep(PAUSA_APOS_REFRESH)
                navegar_para_listagem(driver)
                reabrir_busca_avancada_e_modal(driver)
                selecionar_unidade_e_buscar(driver, id_unidade)
                definir_quantidade_por_pagina(driver)
                for _ in range(pagina_atual - 1): avancar_pagina(driver)
                tentativa_pag += 1
                continue

        processadas_global[str(id_unidade)] = {"pagina": pagina_atual, "acumulado": total_confirmado}
        salvar_checkpoint(processadas_global)
        
        if not avancar_pagina(driver): break
        pagina_atual += 1

    total_fisico = contar_pdfs_nos_zips(pasta_destino)
    print(f"[LOG] Unidade {id_unidade} concluida. Total auditado: {total_fisico} arquivos.")
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

def voltar_pagina(driver):
    info_ant = ""
    try: info_ant = driver.find_element(By.ID, "DataTables_Table_0_info").text
    except: pass
    seletores = ["//li[contains(@class,'prev') and not(contains(@class,'disabled'))]/a", "//li[contains(@class,'previous') and not(contains(@class,'disabled'))]/a"]
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

def baixar_zip_unidade(driver, pagina_atual, pasta_destino, is_last_page, esperado_selecao):
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    
    # 1. Clicar em 'Baixar Certificados' para abrir o modal
    print("Abrindo modal de processamento...")
    btn_gerar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Baixar Certificados')]")))
    driver.execute_script("arguments[0].click();", btn_gerar)
    
    # VERIFICACAO RIGOROSA: Total no modal deve bater com o esperado
    try:
        print("Aguardando estabilização do total no modal...")
        wait_curto = WebDriverWait(driver, 35)
        
        # 1. Aguarda o spinner sumir (ng-hide) para garantir que o processamento do modal terminou
        try:
            wait_curto.until(EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'fa-spinner') and contains(@class, 'ng-hide')]")))
        except:
            print("Aviso: Spinner não indicou finalização, mas prosseguindo.")

        # 2. Localiza o elemento de "Total Selecionado" (A nova Meta Real conforme HTML fornecido)
        xpath_selecionado = "//div[contains(@class, 'total-certificados')]"
        el_selecionado = wait_curto.until(EC.visibility_of_element_located((By.XPATH, xpath_selecionado)))
        
        # 3. Localiza o texto de "TOTAL PROCESSADO" para capturar o "X de Y"
        xpath_total = "//*[contains(text(), 'TOTAL PROCESSADO:')]"
        total_element = wait_curto.until(EC.visibility_of_element_located((By.XPATH, xpath_total)))
        
        time.sleep(4) # Estabilização Angular
        
        texto_sel = el_selecionado.text
        texto_proc = total_element.text
        print(f"Auditoria Modal -> Selecionado: '{texto_sel}' | Processado: '{texto_proc}'")
        
        # Extrai números de ambos
        match_sel = re.search(r'Total Selecionado:\s*([\d.,]+)', texto_sel, re.IGNORECASE)
        match_proc = re.search(r'TOTAL PROCESSADO:\s*(\d+)\s*de\s*([\d.,]+)', texto_proc, re.IGNORECASE)
        
        if match_sel and match_proc:
            meta_real = int(match_sel.group(1).replace('.', '').replace(',', ''))
            total_no_modal = int(match_proc.group(2).replace('.', '').replace(',', ''))
            
            print(f"Meta detectada: {meta_real} | Total no Modal: {total_no_modal} (Esperado: {esperado_selecao})")
            
            # VALIDAÇÃO CRÍTICA: Se a meta real ou o total do modal fugirem do esperado (ex: 1050)
            if meta_real != esperado_selecao or total_no_modal != esperado_selecao:
                print(f"[ALERTA] Divergência! Meta Real {meta_real} ou Modal {total_no_modal} não batem com {esperado_selecao}.")
                try:
                    btn_cancelar = driver.find_element(By.XPATH, "//button[contains(text(),'Cancelar')]")
                    driver.execute_script("arguments[0].click();", btn_cancelar)
                    WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".modal-backdrop, .modal")))
                except: pass
                return "CORRIGIR_SELECAO"
        else:
            # Fallback caso um dos elementos não tenha o número
            print("Aviso: Falha ao extrair números do modal. Verificando apenas o processado...")
            if match_proc:
                total_no_modal = int(match_proc.group(2).replace('.', '').replace(',', ''))
                if total_no_modal != esperado_selecao:
                    # Cancelar...
                    try:
                        btn_cancelar = driver.find_element(By.XPATH, "//button[contains(text(),'Cancelar')]")
                        driver.execute_script("arguments[0].click();", btn_cancelar)
                    except: pass
                    return "CORRIGIR_SELECAO"

    except Exception as e:
        print(f"Aviso: Não conseguiu validar meta no modal ({e}). Prosseguindo...")

    # 2. Clicar em 'Iniciar' no modal
    print("Clicando no botao Iniciar...")
    btn_confirmar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Iniciar')]")))
    driver.execute_script("arguments[0].click();", btn_confirmar)
    
    # 3. Aguardar o processamento e o surgimento do botao 'Download do ZIP'
    print("Processando certificados... Aguardando conclusao.")
    btn_download = None
    inicio_proc = time.time()
    ultimo_log_proc = time.time()
    
    # Seletor ultra-específico: Texto exato e comando exato
    xpath_zip = "//button[contains(text(), 'Download do ZIP') and @ng-click='downloadCertificados();']"
    
    while time.time() - inicio_proc < TIMEOUT_DOWNLOAD:
        try:
            # Busca apenas botões que estejam visíveis
            botoes = driver.find_elements(By.XPATH, xpath_zip)
            for b in botoes:
                if b.is_displayed() and b.is_enabled():
                    btn_download = b
                    break
            if btn_download: break
        except: pass
        
        if time.time() - ultimo_log_proc > 30:
            print(f"... Ainda processando no portal (aguardando ha {int(time.time() - inicio_proc)}s) ...")
            ultimo_log_proc = time.time()
            
        time.sleep(2)
    
    if not btn_download:
        raise TimeoutException("O botao 'Download do ZIP' nao apareceu ou nao ficou pronto para clique.")

    # 4. Iniciar o download
    print(f"Botao de download confirmado! Disparando transferencia via Script...")
    # Pega todos os arquivos atuais para comparar depois
    antes = set([os.path.join(PASTA_DOWNLOAD, f) for f in os.listdir(PASTA_DOWNLOAD)])
    
    # Usamos execute_script como primeira opção para garantir que o clique seja registrado
    # mesmo se houver um overlay invisível do modal na frente.
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_download)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", btn_download)
    
    # 5. Aguardar o arquivo no disco
    arquivo = aguardar_download(PASTA_DOWNLOAD, antes)
    
    # --- AUDITORIA FÍSICA DO ARQUIVO ---
    ext = os.path.splitext(arquivo)[1].lower()
    print(f"Auditando integridade do arquivo {os.path.basename(arquivo)} (Tipo: {ext})...")
    
    total_no_arquivo = 0
    if ext == ".zip":
        try:
            with zipfile.ZipFile(arquivo, 'r') as z:
                # Conta apenas arquivos que terminam com .pdf (ignora pastas ou arquivos de sistema)
                lista_arquivos = [f for f in z.namelist() if f.lower().endswith('.pdf')]
                total_no_arquivo = len(lista_arquivos)
            print(f"Auditoria ZIP: {total_no_arquivo} PDFs encontrados.")
        except Exception as e:
            print(f"[ERRO] Falha ao ler ZIP: {e}")
            return "RETRY"
    elif ext == ".pdf":
        total_no_arquivo = 1
        print(f"Auditoria PDF: Identificado arquivo PDF único.")
    else:
        print(f"[AVISO] Extensão inesperada ({ext}). Tratando como 1 arquivo para continuidade.")
        total_no_arquivo = 1

    if total_no_arquivo < esperado_selecao and ext == ".zip":
        print(f"FALHA NA AUDITORIA! Esperados {esperado_selecao} PDFs, mas o ZIP contém apenas {total_no_arquivo}.")
        print("Removendo arquivo defeituoso e solicitando nova tentativa...")
        try: os.remove(arquivo)
        except: pass
        return "RETRY"
    
    print("Auditoria concluída com sucesso!")

    # 6. Mover/Renomear para a pasta da unidade
    os.makedirs(pasta_destino, exist_ok=True)
    if ext == ".zip":
        nome_final = f"pagina {pagina_atual}.zip"
    else:
        nome_final = f"pagina {pagina_atual}{ext}"
        
    destino_final = os.path.join(pasta_destino, nome_final)
    
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
            except Exception as e:
                print(f"[ERRO] Falha ao identificar descendentes da raiz {id_raiz}: {e}")
                continue

            for id_u in ids_desc:
                if processadas_global.get(str(id_u), {}).get("pagina") == 9999: continue
                print(f"\n[LOG] Iniciando processamento da Unidade ID: {id_u}")
                try:
                    def rodar():
                        nonlocal driver, id_u, processadas_global, pasta_raiz
                        cb = WebDriverWait(driver, TIMEOUT_PADRAO).until(EC.presence_of_element_located((By.ID, f"input_selected_unit_{id_u}")))
                        li = cb.find_element(By.XPATH, "./ancestor::li[1]")
                        if li.find_elements(By.XPATH, ".//i[contains(@class,'fa-plus-square-o')]") or li.find_elements(By.XPATH, ".//a[starts-with(@id, 'link_child_')]"):
                            print(f"[INFO] Unidade {id_u} possui descendentes. Pulando processamento direto...")
                            processadas_global[str(id_u)] = {"pagina": 9999, "acumulado": 0}
                            salvar_checkpoint(processadas_global); return

                        caminho = obter_caminho_hierarquico(driver, cb)
                        pasta_u = os.path.join(pasta_raiz, *caminho)
                        print(f"[CAMINHO] {' > '.join(caminho)}")
                        
                        selecionar_unidade_e_buscar(driver, id_u)
                        if not verificar_tabela_vazia(driver):
                            definir_quantidade_por_pagina(driver)
                            processar_paginas_da_unidade(driver, caminho[-1], pasta_u, id_u, processadas_global)
                        else:
                            print(f"[INFO] Unidade {id_u} nao possui registros para baixar.")
                            processadas_global[str(id_u)] = {"pagina": 9999, "acumulado": 0}
                            salvar_checkpoint(processadas_global)

                    executar_com_retry(driver, rodar)
                    print(f"[LOG] Finalizando Unidade {id_u}. Reiniciando estado para proxima...")
                    driver.refresh(); navegar_para_listagem(driver); reabrir_busca_avancada_e_modal(driver)
                except Exception as e:
                    print(f"[ERRO CRITICO] Falha na Unidade {id_u}: {e}")
                    driver.refresh(); navegar_para_listagem(driver); reabrir_busca_avancada_e_modal(driver)

        if os.path.exists(pasta_raiz) and os.listdir(pasta_raiz):
            print("[LOG] Criando arquivo consolidado de resultados...")
            shutil.make_archive(os.path.join(PASTA_DOWNLOAD, "Resultado_Final_Bot"), 'zip', pasta_raiz)
            print("\n[SUCESSO] PROCESSO COMPLETO E FINALIZADO.")

    except Exception as e:
        print(f"\n[ERRO FATAL] O bot parou devido a um erro nao tratado: {e}")
        traceback.print_exc()
        time.sleep(60)
    finally:
        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__": main()