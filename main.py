from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
import time, os, re, unicodedata, shutil, json, traceback, zipfile, logging

# ==========================
# CONFIGURACAO DE LOGGING
# ==========================
LOG_FILE = "log_processamento.txt"
formatter = logging.Formatter('%(asctime)s - [%(levelname)s] - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger()
logger.setLevel(logging.INFO)
if not logger.handlers:
    fh = logging.FileHandler(LOG_FILE, encoding='utf-8'); fh.setFormatter(formatter); logger.addHandler(fh)
    ch = logging.StreamHandler(); ch.setFormatter(formatter); logger.addHandler(ch)

logging.info("==================================================")
logging.info("INICIALIZANDO BOT DE MIGRACAO RPA")
logging.info("==================================================")

def carregar_config():
    if not os.path.exists("configs.json"): logging.error("configs.json ausente!"); exit(1)
    with open("configs.json", "r", encoding="utf-8") as f: return json.load(f)

CONFIG = carregar_config()
DRIVE = CONFIG.get("drive_letter", "C").upper().strip().replace(":", "")
PASTA_DOWNLOAD = os.path.abspath(os.path.join(f"{DRIVE}:\\", "RPA_MIGRACAO_TEMP"))
TIMEOUT_PADRAO, TIMEOUT_DOWNLOAD = 30, 600
PAUSA_APOS_REFRESH, ARQUIVO_CHECKPOINT = 3, "checkpoint.json"

os.makedirs(PASTA_DOWNLOAD, exist_ok=True)
ambiente_escolhido = next((d for d in CONFIG.get("ambientes", {}).values() if d.get("ativo")), None)
if not ambiente_escolhido: logging.error("Ambiente ativo ausente."); exit(1)

URL_LOGIN, URL_LISTAGEM = ambiente_escolhido["url_login"], ambiente_escolhido["url_listagem"]
LOGIN, SENHA = ambiente_escolhido["credenciais"]["login"], ambiente_escolhido["credenciais"]["senha"]
UNIDADES_ALVO = [str(u) for u in ambiente_escolhido.get("unidades", [])]

def carregar_checkpoint():
    if os.path.exists(ARQUIVO_CHECKPOINT):
        with open(ARQUIVO_CHECKPOINT, "r", encoding="utf-8") as f:
            try: return json.load(f)
            except: return {}
    return {}

def salvar_checkpoint(processadas):
    try:
        with open(ARQUIVO_CHECKPOINT, "w", encoding="utf-8") as f: json.dump(processadas, f, indent=2)
    except Exception as e: logging.error(f"Erro checkpoint: {e}")

def normalizar_nome_unidade(nome):
    if not nome or nome.strip() == ";;": return None
    return re.sub(r'[\\/*?:"<>|]', "", unicodedata.normalize('NFKD', nome).encode('ASCII', 'ignore').decode('ASCII')).strip()

def contar_pdfs_nos_zips(pasta):
    if not os.path.exists(pasta): return 0
    total = 0
    for f in os.listdir(pasta):
        caminho_f = os.path.join(pasta, f)
        if f.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(caminho_f, 'r') as z: total += len([n for n in z.namelist() if n.lower().endswith('.pdf')])
            except: pass
        elif f.lower().endswith('.pdf'): total += 1
    return total

def aguardar_download(pasta, arquivos_antes, timeout=600):
    inicio = time.time()
    while time.time() - inicio < timeout:
        arquivos_agora = set([os.path.join(pasta, f) for f in os.listdir(pasta) if not f.endswith(".crdownload") and not f.endswith(".tmp")])
        novos = arquivos_agora - arquivos_antes
        if novos:
            candidato = list(novos)[0]
            for _ in range(10):
                try:
                    t1 = os.path.getsize(candidato); time.sleep(2)
                    if t1 == os.path.getsize(candidato) and t1 > 0: return candidato
                except: pass
        time.sleep(1)
    raise TimeoutException("Download falhou.")

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
        el = wait.until(lambda d: d.find_element(By.ID, "DataTables_Table_0_info") if d.find_element(By.ID, "DataTables_Table_0_info").text.strip() else False)
        m = re.search(r'([\d.,]+)\s+até\s+([\d.,]+)\s+de\s+([\d.,]+)', el.text)
        if m: return int(re.sub(r'\D', '', m.group(1))), int(re.sub(r'\D', '', m.group(2))), int(re.sub(r'\D', '', m.group(3)))
    except: pass
    return None, None, None

def verificar_tabela_vazia(driver):
    aguardar_carregamento(driver)
    try:
        for sel in ["//td[@class='dataTables_empty']", "//*[contains(text(), 'Não foram encontrados resultados')]"]:
            if driver.find_elements(By.XPATH, sel): return True
        return False
    except: return False

def esta_na_tela_de_login(driver):
    try: return len(driver.find_elements(By.ID, "login")) > 0 or len(driver.find_elements(By.CLASS_NAME, "form-login")) > 0
    except: return False

def executar_com_retry(driver, funcao, *args, **kwargs):
    tentativa = 1
    while tentativa <= 5:
        try: return funcao(*args, **kwargs)
        except Exception as e:
            logging.warning(f"Retry {tentativa}/5: {e}"); tentativa += 1
            try:
                driver.refresh(); time.sleep(PAUSA_APOS_REFRESH)
                if esta_na_tela_de_login(driver): fazer_login(driver)
                navegar_para_listagem(driver); reabrir_busca_avancada_e_modal(driver)
            except: pass
    raise Exception("Falha total.")

def fazer_login(driver):
    logging.info("Autenticando...")
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    if not esta_na_tela_de_login(driver): driver.get(URL_LOGIN); time.sleep(2)
    try: wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Entendi')]"))).click()
    except: pass
    try:
        u = wait.until(EC.visibility_of_element_located((By.ID, "login"))); u.clear(); u.send_keys(LOGIN)
        p = wait.until(EC.visibility_of_element_located((By.ID, "inputPassword1"))); p.clear(); p.send_keys(SENHA)
        driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.XPATH, "//button[@type='submit']"))))
        wait.until(lambda d: not esta_na_tela_de_login(d) and ("/portal" in d.current_url or "/admin" in d.current_url))
        logging.info("Autenticado.")
    except Exception as e: logging.error(f"Erro login: {e}")
    time.sleep(3)

def navegar_para_listagem(driver):
    while True:
        if "/login" in driver.current_url or esta_na_tela_de_login(driver): fazer_login(driver)
        driver.get(URL_LISTAGEM); time.sleep(5)
        if "dashboard" in driver.current_url:
            for s in ["//a[contains(text(),'Admin')]", "//a[contains(text(),'Usuários Certificados')]"]:
                try: WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, s))).click(); time.sleep(3)
                except: continue
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//table | //button[@idaut='botao_busca_avancada']")))
            return
        except: driver.refresh(); time.sleep(5)

def expandir_arvore(driver):
    logging.info("Abrindo arvore...")
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//i[contains(@class,'fa-plus-square-o')]")))
        while True:
            icones = driver.find_elements(By.XPATH, "//i[contains(@class,'fa-plus-square-o')]")
            if not icones: break
            for i in icones:
                try: driver.execute_script("arguments[0].click();", i.find_element(By.XPATH, "./ancestor::*[self::a or self::span or self::li][1]"))
                except: continue
            time.sleep(1)
    except: pass

def obter_caminho_hierarquico(driver, checkbox):
    caminho = []
    try:
        li = checkbox.find_element(By.XPATH, "./ancestor::li[1]")
        while li:
            nome = None
            for sel in [".//a[starts-with(@id, 'link_child_')]", ".//span[contains(@class, 'ng-binding')]"]:
                try:
                    el = li.find_element(By.XPATH, sel)
                    if el.text.strip(): nome = el.text.strip(); break
                except: continue
            if nome:
                norm = normalizar_nome_unidade(nome)
                if norm: caminho.insert(0, norm)
            try: li = li.find_element(By.XPATH, "ancestor::li[1]")
            except: break
    except: pass
    return caminho if caminho else ["Unidade_Desconhecida"]

def obter_todos_ids_descendentes(driver, checkbox_raiz):
    li = checkbox_raiz.find_element(By.XPATH, "./ancestor::li[1]")
    return [re.search(r'_(\d+)', cb.get_attribute("id")).group(1) for cb in li.find_elements(By.XPATH, ".//input[starts-with(@id, 'input_selected_unit_')]") if re.search(r'_(\d+)', cb.get_attribute("id"))]

def selecionar_unidade_e_buscar(driver, id_u):
    logging.info(f"Unidade ID {id_u}")
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)
    try:
        driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.ID, f"input_selected_unit_{id_u}"))))
        driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.ID, "btn_select_unit"))))
        driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Buscar']"))))
        time.sleep(3)
    except: pass

def definir_quantidade_por_pagina(driver):
    try:
        sel = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//select[option[@value='100']]")))
        Select(sel).select_by_value("100"); time.sleep(2); aguardar_carregamento(driver)
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[@title='Buscar']"))
        time.sleep(4); aguardar_carregamento(driver)
    except: pass

def baixar_individualmente(driver, pasta, esperado, pagina):
    logging.info(f"MODO MANUAL: Pagina {pagina} Meta {esperado}")
    os.makedirs(pasta, exist_ok=True)
    baixados = 0
    linhas = driver.find_elements(By.XPATH, "//table[contains(@id, 'DataTables_Table')]//tbody//tr[not(contains(@class, 'dataTables_empty'))]")
    for i, linha in enumerate(linhas):
        if i >= esperado: break
        try:
            btn = linha.find_element(By.XPATH, ".//button[i[contains(@class, 'fa-download')]]")
            antes = set([os.path.join(PASTA_DOWNLOAD, f) for f in os.listdir(PASTA_DOWNLOAD)])
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            time.sleep(1); driver.execute_script("arguments[0].click();", btn)
            try:
                arq = aguardar_download(PASTA_DOWNLOAD, antes, timeout=60)
                dest = os.path.join(pasta, f"pagina_{pagina}_item_{i+1}" + os.path.splitext(arq)[1].lower())
                if os.path.exists(dest): os.remove(dest)
                shutil.move(arq, dest); baixados += 1
            except: pass
        except: pass
    return baixados >= esperado

def processar_paginas_da_unidade(driver, nome, pasta, id_u, global_cp):
    pagina = 1; cp = global_cp.get(str(id_u), {"pagina": 0, "acumulado": 0})
    ultima_ok, total_acc = cp.get("pagina", 0), cp.get("acumulado", 0)
    while True:
        aguardar_carregamento(driver)
        de, ate, total = obter_info_paginacao(driver)
        if de is None or verificar_tabela_vazia(driver): break
        is_last, meta = (ate == total), (ate - de + 1)
        
        if pagina <= ultima_ok:
            total_acc = ate; 
            if not avancar_pagina(driver): break
            pagina += 1; continue

        logging.info(f"--- Pagina {pagina} (Meta: {meta}) - DOWNLOAD MANUAL ---")
        try:
            if baixar_individualmente(driver, pasta, meta, pagina):
                total_acc += meta
                global_cp[str(id_u)] = {"pagina": pagina, "acumulado": total_acc}
                salvar_checkpoint(global_cp)
            else:
                logging.error(f"Falha no download manual da P{pagina}. Tentando novamente apos refresh.")
                driver.refresh(); time.sleep(PAUSA_APOS_REFRESH)
                navegar_para_listagem(driver); reabrir_busca_avancada_e_modal(driver); selecionar_unidade_e_buscar(driver, id_u); definir_quantidade_por_pagina(driver)
                for _ in range(pagina - 1): avancar_pagina(driver)
                if baixar_individualmente(driver, pasta, meta, pagina):
                    total_acc += meta; global_cp[str(id_u)] = {"pagina": pagina, "acumulado": total_acc}; salvar_checkpoint(global_cp)
                else:
                    logging.error(f"Falha persistente na P{pagina}. Pulando para proxima unidade.")
                    break
        except Exception as e: 
            logging.error(f"Erro P{pagina}: {e}"); driver.refresh(); time.sleep(3); continue
        
        if not avancar_pagina(driver): break
        pagina += 1
    global_cp[str(id_u)] = {"pagina": 9999, "acumulado": contar_pdfs_nos_zips(pasta)}; salvar_checkpoint(global_cp)

def avancar_pagina(driver):
    try:
        info_ant = driver.find_element(By.ID, "DataTables_Table_0_info").text
        btn = driver.find_element(By.XPATH, "//li[contains(@class,'next') and not(contains(@class,'disabled'))]/a")
        driver.execute_script("arguments[0].click();", btn)
        WebDriverWait(driver, 20).until(lambda d: d.find_element(By.ID, "DataTables_Table_0_info").text != info_ant)
        aguardar_carregamento(driver); return True
    except: return False

def reabrir_busca_avancada_e_modal(driver):
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        driver.execute_script("arguments[0].click();", WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[idaut='botao_busca_avancada']"))))
        time.sleep(2); driver.find_element(By.ID, "search_unidades").click(); time.sleep(2); expandir_arvore(driver)
    except: pass

def main():
    options = Options(); options.add_argument("--start-maximized"); options.add_argument("--incognito"); options.add_argument("--disable-features=DownloadBubble,DownloadBubbleV2")
    prefs = {"profile.default_content_setting_values.automatic_downloads": 1, "download.default_directory": PASTA_DOWNLOAD, "download.prompt_for_download": False, "safebrowsing.enabled": True}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": PASTA_DOWNLOAD})
    processadas = carregar_checkpoint()
    try:
        fazer_login(driver); navegar_para_listagem(driver); reabrir_busca_avancada_e_modal(driver)
        if not UNIDADES_ALVO or "XXXX" in UNIDADES_ALVO:
            ids_trab = [re.search(r'_(\d+)', el.get_attribute("id")).group(1) for el in driver.find_elements(By.XPATH, "//input[starts-with(@id, 'input_selected_unit_')]") if re.search(r'_(\d+)', el.get_attribute("id"))]
        else:
            ids_trab = []
            for r in UNIDADES_ALVO:
                try: ids_trab.extend(obter_todos_ids_descendentes(driver, WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, f"input_selected_unit_{r}")))))
                except: ids_trab.append(r)
        ids_trab = list(dict.fromkeys(ids_trab)); pasta_raiz = os.path.join(PASTA_DOWNLOAD, "resultado do bot"); os.makedirs(pasta_raiz, exist_ok=True)
        for id_u in ids_trab:
            if processadas.get(str(id_u), {}).get("pagina") == 9999: continue
            logging.info(f"--- UNIDADE {id_u} ---")
            try:
                def rodar():
                    cb = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, f"input_selected_unit_{id_u}")))
                    li = cb.find_element(By.XPATH, "./ancestor::li[1]")
                    if li.find_elements(By.XPATH, ".//i[contains(@class,'fa-plus-square-o')]") or li.find_elements(By.XPATH, ".//a[starts-with(@id, 'link_child_')]"):
                        processadas[str(id_u)] = {"pagina": 9999, "acumulado": 0}; salvar_checkpoint(processadas); return
                    caminho = obter_caminho_hierarquico(driver, cb); selecionar_unidade_e_buscar(driver, id_u)
                    if not verificar_tabela_vazia(driver):
                        definir_quantidade_por_pagina(driver); processar_paginas_da_unidade(driver, caminho[-1], os.path.join(pasta_raiz, *caminho), id_u, processadas)
                    else: processadas[str(id_u)] = {"pagina": 9999, "acumulado": 0}; salvar_checkpoint(processadas)
                executar_com_retry(driver, rodar)
                driver.refresh(); navegar_para_listagem(driver); reabrir_busca_avancada_e_modal(driver)
            except Exception as e: logging.error(f"Falha unidade {id_u}: {e}"); driver.refresh(); navegar_para_listagem(driver); reabrir_busca_avancada_e_modal(driver)
        if os.path.exists(pasta_raiz) and os.listdir(pasta_raiz):
            shutil.make_archive(os.path.join(PASTA_DOWNLOAD, "Resultado_Final_Bot"), 'zip', pasta_raiz)
            logging.info("PROCESSO COMPLETO.")
    except Exception as e: logging.critical(f"Fatal: {e}\n{traceback.format_exc()}")
    finally:
        try: driver.quit()
        except: pass

if __name__ == "__main__": main()
