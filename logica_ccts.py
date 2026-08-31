import time
import os
import re
import json
import shutil
import win32com.client
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException

LOJAS_POR_REINICIO = 10

class ReinicioProgramadoNecessario(Exception):
    """Sinaliza que o driver deve ser reiniciado (reinício programado a cada N lojas)."""

URL_LOGIN = "https://queroquero.mybeehome.com/login"
URL_WIKI = "https://queroquero.mybeehome.com/wiki/492/detail"
SECAO_ACORDOS = "Acordos e Convenções Coletivas"
DIRETORIO_ORIGEM = ""
PASTA_TEMP = ""
LOG_FILE_PATH = ""

def log(message, log_file, espacamento=False, registrar_em_arquivo=True):
    try:
        if espacamento:
            log_message = "\n" + "-" * 80 + "\n" + message.strip() + "\n"
        else:
            log_message = message.strip()
        print(log_message)
        if registrar_em_arquivo:
            log_file.write(log_message + "\n")
            log_file.flush()
    except Exception as e:
        print(f"[ERRO] Falha ao escrever no log: {e}")

def converter_doc_para_pdf(caminho_arquivo, diretorio_saida, log_file):
    word = None; doc = None
    try:
        nome_base = os.path.splitext(os.path.basename(caminho_arquivo))[0]
        caminho_pdf = os.path.join(diretorio_saida, f"{nome_base}.pdf")
        caminho_arquivo_abs = os.path.abspath(caminho_arquivo)
        caminho_pdf_abs = os.path.abspath(caminho_pdf)
        log(f"[INFO] Convertendo '{os.path.basename(caminho_arquivo)}' para PDF...", log_file)
        if not os.path.exists(diretorio_saida): os.makedirs(diretorio_saida)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(caminho_arquivo_abs)
        doc.SaveAs(caminho_pdf_abs, FileFormat=17)
        log(f"[SUCESSO] Arquivo convertido para PDF: {caminho_pdf_abs}", log_file)
        return caminho_pdf
    except Exception as e:
        log(f"[ERRO] Falha ao converter o arquivo '{os.path.basename(caminho_arquivo)}' para PDF. Erro: {e}", log_file)
        return None
    finally:
        if doc: doc.Close(0)
        if word: word.Quit()

def extrair_numeros_loja(nome_arquivo):
    primeira_parte = nome_arquivo.split('_')[0]
    return re.findall(r'\b\d{1,3}\b', primeira_parte)

def carregar_flag_criar_loja():
    """Lê do config.json se a criação automática de loja está habilitada."""
    try:
        caminho_config = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        with open(caminho_config, 'r', encoding='utf-8') as f:
            return bool(json.load(f).get("criar_loja_automaticamente", False))
    except Exception:
        return False

def extrair_cidade_loja(nome_arquivo_limpo):
    """Extrai a cidade do nome do arquivo (2º token separado por '_'). Ex.: '690_LAJEADO_SIND0088_2026' -> 'LAJEADO'."""
    partes = nome_arquivo_limpo.split('_')
    return partes[1] if len(partes) > 1 and partes[1] else None

def montar_nome_loja(numero_formatado, cidade):
    """Monta o nome usado na criação da página. Ex.: '690 - LAJEADO' (ou apenas '690' se não houver cidade)."""
    return f"{numero_formatado} - {cidade}" if cidade else numero_formatado

def _fechar_popup_onesignal(driver, timeout=3):
    try:
        popup_wait = WebDriverWait(driver, timeout)
        popup_wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='onesignal-slidedown-cancel-button']"))).click()
    except TimeoutException:
        pass

# navegação por cliques até a página de categorias de Acordos 
def navegar_para_categoria_acordos(driver, wait, log_file):
    """Navega até a página de categorias de 'Acordos e Convenções Coletivas' usando apenas cliques."""
    log("[CRIACAO] Navegando até a página de categorias de Acordos...", log_file)

    # Sidebar → "Biblioteca do Conhecimento"
    link_biblioteca = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//a[@href='/wiki']//span[normalize-space(.)='Biblioteca do Conhecimento']/ancestor::a")
    ))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link_biblioteca)
    time.sleep(1)
    link_biblioteca.click()
    log("[CRIACAO] Passo 1/5: Clicou em 'Biblioteca do Conhecimento'.", log_file)

    # Clica em "Ver mais" (pelo texto exato do link)
    time.sleep(3)
    botao_ver_mais = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//a[normalize-space(.)='Ver mais']")
    ))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_ver_mais)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", botao_ver_mais)
    log("[CRIACAO] Passo 2/5: Clicou em 'Ver mais'.", log_file)

    # Clica no h6 "Remuneração (C&S)" que direciona para a página /wiki/130/category
    time.sleep(3)
    texto_remuneracao = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//h6[normalize-space(.)='Remuneração (C&S)']")
    ))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", texto_remuneracao)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", texto_remuneracao)
    log("[CRIACAO] Passo 3/5: Clicou em 'Remuneração (C&S)'.", log_file)

    # Toggle "Remuneração CS" na página de categoria
    toggle_remuneracao = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[@aria-label='Toggle Remuneração CS']")
    ))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", toggle_remuneracao)
    toggle_remuneracao.click()
    log("[CRIACAO] Passo 4/5: Expandiu toggle 'Remuneração CS'.", log_file)

    # Toggle "Acordos e Convenções Coletivas"
    toggle_acordos = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[@aria-label='Toggle Acordos e Convenções Coletivas']")
    ))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", toggle_acordos)
    toggle_acordos.click()
    log("[CRIACAO] Passo 5/5: Expandiu toggle 'Acordos e Convenções Coletivas'.", log_file)

    return True


def criar_loja_no_comuniq(driver, wait, nome_loja, log_file):
    """Cria a página da loja no ComuniQ (modelo 'Em branco') preenchendo título, descrição e texto com o mesmo valor."""
    try:
        log(f"[CRIACAO] Loja inexistente — criando página '{nome_loja}' no ComuniQ...", log_file)
       
        if not navegar_para_categoria_acordos(driver, wait, log_file):
            return 'FALHA_CRIACAO_LOJA'

        seletor_criar_acordos = (
            "//span[contains(@class,'mdc-button__label') and normalize-space(.)='" + SECAO_ACORDOS + "']"
            "/ancestor::div[contains(@class,'full-width')][1]"
            "/following-sibling::a[contains(@class,'btn-link')][1]"
        )
        wait.until(EC.visibility_of_element_located((By.XPATH, f"//span[contains(@class,'mdc-button__label') and normalize-space(.)='{SECAO_ACORDOS}']")))
        botao_criar = wait.until(EC.element_to_be_clickable((By.XPATH, seletor_criar_acordos)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_criar)
        time.sleep(3)
        botao_criar.click()
        log("[CRIACAO] Clicou no botão '+' (criar página) da seção de Acordos.", log_file)

        wait.until(EC.visibility_of_element_located((By.XPATH, "//h5[normalize-space()='Escolha o modelo da página']")))
        wait.until(EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='Em branco']/ancestor::div[contains(@class,'card')][1]//button"))).click()

        campo_titulo = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@formcontrolname='title']")))
        campo_titulo.clear(); campo_titulo.send_keys(nome_loja)

        campo_descricao = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@formcontrolname='description']")))
        campo_descricao.clear(); campo_descricao.send_keys(nome_loja)

        campo_texto = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'fr-element') and @contenteditable='true']")))
        campo_texto.click(); campo_texto.send_keys(nome_loja)

        botao_salvar = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'footer-bar')]//button[contains(@class,'btn-success')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", botao_salvar)
        botao_salvar.click()
        log(f"[DEBUG] Página '{nome_loja}' salva; aguardando aparecimento na lista de acordos...", log_file)

        if not navegar_para_acordos(driver, wait):
            return 'FALHA_CRIACAO_LOJA'

        confirmacao_xpath = f"//span[contains(@class, 'mdc-button__label') and contains(normalize-space(.), '{nome_loja}')]"
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, confirmacao_xpath)))
        except TimeoutException:
            log(f"[ERRO] Página '{nome_loja}' salva, mas o botão da loja não apareceu na lista de acordos.", log_file)
            return 'FALHA_CRIACAO_LOJA'

        log(f"[SUCESSO] Página da loja '{nome_loja}' criada no ComuniQ.", log_file)
        return 'OK'
    except Exception as e:
        log(f"[ERRO] Falha ao criar a loja '{nome_loja}'. Erro: {e}", log_file)
        return 'FALHA_CRIACAO_LOJA'

def normalizar(nome):
    n = nome.strip().lower()
    n = re.sub(r'\s*\(\d+\)\s*$', '', n)
    if n.endswith('.pdf'):
        n = n[:-4]
    n = re.sub(r'[\s_-]+', '_', n)
    return n

def renomear_e_preparar_arquivo(nome_real_com_prefixo, nome_limpo_para_processar, log_file):
    caminho_origem_completo = os.path.join(DIRETORIO_ORIGEM, nome_real_com_prefixo)
    
    nome_sem_extensao, extensao = os.path.splitext(nome_limpo_para_processar)
    caminho_para_upload = caminho_origem_completo
    extensao_final = extensao

    if extensao.lower() in ['.doc', '.docx']:
        caminho_pdf_convertido = converter_doc_para_pdf(caminho_origem_completo, PASTA_TEMP, log_file)
        if not caminho_pdf_convertido:
            return None
        caminho_para_upload = caminho_pdf_convertido
        extensao_final = '.pdf'
        nome_sem_extensao = os.path.splitext(os.path.basename(caminho_pdf_convertido))[0]

    match_vigencia = re.search(r'_(\d{4}(?:-\d{4})?)', nome_sem_extensao)
    if not match_vigencia:
        log(f"[ERRO] Padrão de VIGÊNCIA não encontrado em: {nome_limpo_para_processar}", log_file)
        if 'caminho_pdf_convertido' in locals() and os.path.exists(caminho_pdf_convertido):
            os.remove(caminho_pdf_convertido)
        return None

    vigencia = match_vigencia.group(1)
    fim_da_vigencia = match_vigencia.end(0)
    complemento_bruto = nome_sem_extensao[fim_da_vigencia:]
    complemento = "" 

    if complemento_bruto and complemento_bruto.startswith('_'):
        complemento = complemento_bruto.lstrip('_')

    novo_nome_base = f"CCT_{vigencia}"
    if complemento:
        novo_nome_base += f"_{complemento}"
    
    novo_nome_final = novo_nome_base + extensao_final

    if not os.path.exists(PASTA_TEMP):
        os.makedirs(PASTA_TEMP)
    
    caminho_destino_temporario = os.path.join(PASTA_TEMP, novo_nome_final)
    try:
        if 'caminho_pdf_convertido' in locals() and os.path.exists(caminho_para_upload):
            if os.path.exists(caminho_destino_temporario):
                os.remove(caminho_destino_temporario)
            shutil.move(caminho_para_upload, caminho_destino_temporario)
        else:
            shutil.copy2(caminho_para_upload, caminho_destino_temporario)
        
        log(f"[INFO] Arquivo '{nome_real_com_prefixo}' preparado para upload como: {novo_nome_final}", log_file)
        return caminho_destino_temporario
    except Exception as e:
        log(f"[ERRO] Falha ao preparar arquivo temporário para {nome_real_com_prefixo}: {e}", log_file)
        return None

def inicializar_driver():
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        print("Chrome driver inicializado com sucesso.")
        return driver
    except Exception as e:
        print(f"Erro fatal ao inicializar o Chrome driver: {e}")
        return None

def fazer_login(driver, wait, usuario, senha):
    try:
        driver.get(URL_LOGIN)
        driver.maximize_window()
        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@formcontrolname='username']"))).send_keys(usuario)
        driver.find_element(By.XPATH, "//input[@formcontrolname='password']").send_keys(senha)
        driver.find_element(By.XPATH, "//button[contains(text(), 'Entrar')]").click()
        wait.until(lambda d: '/login' not in d.current_url)
        print("Login realizado com sucesso.")
        return True
    except Exception as e:
        print(f"Erro durante o login: {e}")
        return False

def navegar_para_acordos(driver, wait):
    try:
        driver.get(URL_WIKI)
        remuneracao_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Toggle Remuneração CS']")))
        remuneracao_toggle.click()
        print("Clicou no Toggle Remuneração CS")
        acordos_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Toggle Acordos e Convenções Coletivas']")))
        acordos_toggle.click()
        print("Clicou no Toggle Acordos e Convenções Coletivas")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", acordos_toggle)
        print("Scroll realizado para a seção de Acordos.")
        return True
    except Exception as e:
        print(f"Erro ao navegar para a seção de acordos: {e}")
        return False

def mover_e_renomear_arquivo(caminho_arquivo_original, pasta_destino, sufixo_tag, log_file):
    try:
        nome_base_arquivo = os.path.basename(caminho_arquivo_original)

        nome_limpo = nome_base_arquivo
        if nome_limpo.startswith("[FEITO]_"):
            nome_limpo = nome_limpo[len("[FEITO]_"):]
        elif nome_limpo.startswith("[FALHA_MANUAL]_"):
            nome_limpo = nome_limpo[len("[FALHA_MANUAL]_"):]
        elif nome_limpo.startswith("[JA_EXISTE]_"):
            nome_limpo = nome_limpo[len("[JA_EXISTE]_"):]
        
        nome_base, extensao = os.path.splitext(nome_limpo)
        
        novo_nome = f"{nome_base}_{sufixo_tag}{extensao}"
        
        caminho_destino_final = os.path.join(pasta_destino, novo_nome)
        
        shutil.move(caminho_arquivo_original, caminho_destino_final)
        
        log(f"[INFO] Arquivo movido e renomeado para: {caminho_destino_final}", log_file)
    except Exception as e:
        log(f"[ERRO] Falha ao mover/renomear o arquivo {caminho_arquivo_original}. Erro: {e}", log_file)

def abrir_edicao_da_loja(driver, wait, numero_loja_formatado, nome_loja=None, log_file=None):
    try:
        loja_xpath = f"//span[contains(@class, 'mdc-button__label') and starts-with(normalize-space(.), '{numero_loja_formatado}')]"
        elemento_loja = wait.until(EC.element_to_be_clickable((By.XPATH, loja_xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_loja)
        elemento_loja.click()
    except TimeoutException:
        if not (carregar_flag_criar_loja() and nome_loja and log_file):
            return 'LOJA_NAO_ENCONTRADA'

        loja_xpath_ampla = f"//span[contains(@class, 'mdc-button__label') and contains(normalize-space(.), '{numero_loja_formatado}')]"
        try:
            elemento_loja = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, loja_xpath_ampla)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_loja)
            elemento_loja.click()
        except TimeoutException:
            if criar_loja_no_comuniq(driver, wait, nome_loja, log_file) != 'OK':
                return 'FALHA_CRIACAO_LOJA'
            try:
                loja_xpath_nova = f"//span[contains(@class, 'mdc-button__label') and contains(normalize-space(.), '{nome_loja}')]"
                elemento_loja = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, loja_xpath_nova)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_loja)
                elemento_loja.click()
            except TimeoutException:
                log(f"[ERRO] Loja '{nome_loja}' criada, mas o botão não ficou clicável.", log_file)
                return 'FALHA_CRIACAO_LOJA'

    try:
        popup_wait = WebDriverWait(driver, 3)
        botao_cancelar = popup_wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='onesignal-slidedown-cancel-button']")))
        botao_cancelar.click()
    except TimeoutException:
        pass

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Ações']"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space(text())='Editar']"))).click()

    anexos = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.dz-clickable")))
    driver.execute_script("arguments[0].scrollIntoView(true);", anexos)

    wait.until(lambda d: len(d.find_elements(By.XPATH, "//input[@type='file']")) >= 3)
    return 'OK'

def obter_anexos_salvos(driver, wait, log_file):
    timeout = 30; check_interval = 0.5; stable_checks_required = 4
    start_time = time.time(); last_count = -1; stable_counter = 0
    nomes_existentes = []

    while time.time() - start_time < timeout:
        try:
            spans = driver.find_elements(By.CSS_SELECTOR, "span[data-dz-name]")
            nomes_existentes = [
                (s.get_attribute('data-dz-name') or s.get_attribute('textContent') or '').strip()
                for s in spans
            ]
            nomes_existentes = [n for n in nomes_existentes if n]
            current_count = len(nomes_existentes)
        except Exception:
            current_count = -1
            nomes_existentes = []

        if current_count == last_count:
            stable_counter += 1
        else:
            stable_counter = 0

        if stable_counter >= stable_checks_required and current_count > 0:
            break

        last_count = current_count
        time.sleep(check_interval)
    else:
        log("[AVISO] Timeout! A lista de arquivos não carregou a tempo (continua com o que conseguiu ler).", log_file)

    if nomes_existentes:
        log(f"[DEBUG] Lista de anexos estabilizada com {len(nomes_existentes)} arquivos.", log_file)
    return nomes_existentes

def upload_arquivo_para_loja(driver, wait, numero_loja_formatado, caminho_arquivo_completo, log_file, nome_arquivo_original, nome_loja=None):
    try:
        log(f"--- Iniciando upload de '{nome_arquivo_original}' para a loja {numero_loja_formatado} ---", log_file, registrar_em_arquivo=False)
        
        status_loja = abrir_edicao_da_loja(driver, wait, numero_loja_formatado, nome_loja, log_file)
        if status_loja == 'LOJA_NAO_ENCONTRADA':
            log(f"[ERRO MANUAL] Loja {numero_loja_formatado} não encontrada para o arquivo '{nome_arquivo_original}'.", log_file)
            return 'FALHA_LOJA_NAO_ENCONTRADA'
        if status_loja == 'FALHA_CRIACAO_LOJA':
            log(f"[ERRO MANUAL] Falha ao criar a loja {numero_loja_formatado} ('{nome_loja}') no ComuniQ para o arquivo '{nome_arquivo_original}'.", log_file)
            return 'FALHA_CRIACAO_LOJA'
        
        input_alvo = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@type='file']")))[2]

        nome_final_alvo = os.path.basename(caminho_arquivo_completo)

        nomes_existentes = obter_anexos_salvos(driver, wait, log_file)
        log(f"[DEBUG] Anexos na loja {numero_loja_formatado}: {nomes_existentes}", log_file)

        contagem_antes = len(nomes_existentes)
        log(f"[INFO] Contagem antes do upload para '{nome_arquivo_original}': {contagem_antes} arquivos.", log_file)

        if any(normalizar(ex) == normalizar(nome_final_alvo) for ex in nomes_existentes):
            log(f"[SKIP] '{nome_final_alvo}' ja existe na loja {numero_loja_formatado}. Pulando.", log_file)
            navegar_para_acordos(driver, wait)
            return 'JA_EXISTE'
        
        input_alvo.send_keys(caminho_arquivo_completo)

        log(f"[OK] Anexado '{nome_arquivo_original}' para a loja {numero_loja_formatado}.", log_file)

        try:
            total_inputs = len(driver.find_elements(By.XPATH, "//input[@type='file']"))
            valor_input = input_alvo.get_attribute('value') or '(vazio)'
            total_previews = len(driver.find_elements(By.CSS_SELECTOR, "div.dz-preview"))
            log(f"[DEBUG] Inputs de arquivo na página: {total_inputs} | Value do input alvo: '{valor_input}' | Pré-visualizações no dropzone: {total_previews}", log_file)
        except Exception as e:
            log(f"[DEBUG] Falha ao inspecionar o dropzone após o envio ({type(e).__name__}).", log_file)
        
        try:
            WebDriverWait(driver, 30).until(
                lambda d: any(
                    'dz-complete' in (classes := (p.get_attribute('class') or ''))
                    and 'dz-error' not in classes
                    for p in d.find_elements(By.CSS_SELECTOR, "div.dz-preview")
                    if any(
                        normalizar((s.get_attribute('data-dz-name') or s.get_attribute('textContent') or '').strip()) == normalizar(nome_final_alvo)
                        for s in p.find_elements(By.CSS_SELECTOR, "span[data-dz-name]")
                    )
                )
            )
        except TimeoutException:
            classes_previews = [
                (p.get_attribute('class') or '')
                for p in driver.find_elements(By.CSS_SELECTOR, "div.dz-preview")
            ]
            log(f"[ERRO] Upload de '{nome_final_alvo}' não concluiu no Dropzone (timeout 30s). Classes dos previews: {classes_previews}", log_file)
            raise Exception("Upload do arquivo não concluído no Dropzone (timeout 30s).")
        
        log(f"[DEBUG] Upload de '{nome_final_alvo}' concluído (dz-success); procurando o botão Salvar.", log_file)

        seletor_salvar = "//button[contains(normalize-space(.), 'Salvar') and contains(@class, 'btn-success')]"

        def _botao_salvar_visivel():
            botoes = [b for b in driver.find_elements(By.XPATH, seletor_salvar)
                      if b.is_displayed() and b.is_enabled()]
            return botoes[0] if botoes else None

        botao_salvar = None
        inicio_espera = time.time()
        while time.time() - inicio_espera < 8:
            botao_salvar = _botao_salvar_visivel()
            if botao_salvar:
                break
            time.sleep(0.5)

        if not botao_salvar:
            todos_botoes = driver.find_elements(By.XPATH, seletor_salvar)
            log(f"[AVISO] Nenhum botão 'Salvar' visível/habilitado em 8s. Total de botões 'Salvar' na página: {len(todos_botoes)}.", log_file)
            for b in todos_botoes:
                try:
                    log(f"[DEBUG]   - 'Salvar' visível={b.is_displayed()} habilitado={b.is_enabled()}", log_file)
                except Exception:
                    pass
            raise Exception("Nenhum botão 'Salvar' visível/habilitado encontrado (timeout 8s).")

        driver.execute_script("arguments[0].scrollIntoView(true);", botao_salvar)
        try:
            botao_salvar.click()
        except Exception:
            log(f"[AVISO] Clique normal no botão Salvar bloqueado; tentando clicar via JavaScript.", log_file)
            driver.execute_script("arguments[0].click();", botao_salvar)
        log(f"[DEBUG] Botão Salvar clicado; aguardando retorno à tela de Ações.", log_file)

        try:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Ações']")))
        except TimeoutException:
            log(f"[AVISO] Botão 'Ações' não reapareceu em 15s na loja {numero_loja_formatado}; reabrindo a loja para conferir.", log_file)

        navegar_para_acordos(driver, wait)
        if abrir_edicao_da_loja(driver, wait, numero_loja_formatado, nome_loja, log_file) == 'OK':
            anexos_salvos = obter_anexos_salvos(driver, wait, log_file)
            if any(normalizar(a) == normalizar(nome_final_alvo) for a in anexos_salvos):
                navegar_para_acordos(driver, wait)
                log(f"[SUCESSO CONFIRMADO] '{nome_final_alvo}' confirmado como salvo na loja {numero_loja_formatado}.", log_file, registrar_em_arquivo=False)
                return 'SUCESSO'
            navegar_para_acordos(driver, wait)
            return 'FALHA_VERIFICACAO'
        else:
            return 'FALHA_VERIFICACAO'

    except Exception as e:
        log(f"[ERRO GERAL] Falha no upload de '{nome_arquivo_original}' para a loja {numero_loja_formatado}. Erro: {e}", log_file)
        driver.get(URL_WIKI)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Toggle Remuneração CS']")))
        navegar_para_acordos(driver, wait)
        return 'FALHA_RECUPERAVEL'

def upload_com_verificacao_e_retry(driver, wait, numero_loja_formatado, caminho_arquivo_completo, log_file, nome_arquivo_original, nome_loja=None):
    max_tentativas = 3
    for tentativa in range(1, max_tentativas + 1):
        resultado = upload_arquivo_para_loja(driver, wait, numero_loja_formatado, caminho_arquivo_completo, log_file, nome_arquivo_original, nome_loja)

        if resultado == 'JA_EXISTE' and tentativa > 1:
            log(f"[SUCESSO CONFIRMADO] '{nome_arquivo_original}' confirmado na loja {numero_loja_formatado} via retry.", log_file)
            return 'SUCESSO'
        if resultado in ('SUCESSO', 'JA_EXISTE', 'FALHA_LOJA_NAO_ENCONTRADA', 'FALHA_CRIACAO_LOJA'):
            return resultado

        log(f"[AVISO] Tentativa {tentativa}/{max_tentativas} falhou ({resultado}) para loja {numero_loja_formatado}.", log_file)
        if tentativa < max_tentativas:
            time.sleep(3)

    log(f"[ERRO] Falha confirmada após {max_tentativas} tentativas na loja {numero_loja_formatado}.", log_file)
    return 'UPLOAD_NAO_CONFIRMADO'

def _reiniciar_sessao(driver, wait, usuario, senha, log_file):

    try:
        driver.quit()
    except Exception:
        pass

    driver = inicializar_driver()
    if not driver:
        log("[ERRO] Falha ao reinicializar o driver.", log_file)
        return None, None

    wait = WebDriverWait(driver, 60)

    if not fazer_login(driver, wait, usuario, senha):
        log("[ERRO] Falha no login após reinicialização.", log_file)
        return None, None

    if not navegar_para_acordos(driver, wait):
        log("[ERRO] Falha na navegação após reinicialização.", log_file)
        return None, None

    return driver, wait

def processar_um_unico_arquivo(driver, wait, nome_arquivo, log_file, lojas_confirmadas=None):
    if lojas_confirmadas is None:
        lojas_confirmadas = {}

    nome_arquivo_completo = nome_arquivo
    
    nome_arquivo_limpo = nome_arquivo
    nome_base, extensao = os.path.splitext(nome_arquivo)

    sufixo_feito = "_[FEITO]"
    sufixo_falha = "_[FALHA_MANUAL]"
    sufixo_ja_existe = "_[JA_EXISTE]"

    if nome_base.endswith(sufixo_feito):
        nome_base_limpo = nome_base[:-len(sufixo_feito)]
        nome_arquivo_limpo = nome_base_limpo + extensao
    elif nome_base.endswith(sufixo_falha):
        nome_base_limpo = nome_base[:-len(sufixo_falha)]
        nome_arquivo_limpo = nome_base_limpo + extensao
    elif nome_base.endswith(sufixo_ja_existe):
        nome_base_limpo = nome_base[:-len(sufixo_ja_existe)]
        nome_arquivo_limpo = nome_base_limpo + extensao

    log(f"Processando arquivo: {nome_arquivo_completo}", log_file, espacamento=True)
    if nome_arquivo_completo != nome_arquivo_limpo:
        log(f"(Nome limpo para processamento: {nome_arquivo_limpo})", log_file)

    resultados_lojas = {'SUCESSO': [], 'JA_EXISTE': [], 'FALHA_MANUAL': [], 'UPLOAD_NAO_CONFIRMADO': [], 'FALHA_RECUPERAVEL': [], 'FALHA_CRIACAO': []}

    numeros_loja = extrair_numeros_loja(nome_arquivo_limpo)
    if not numeros_loja:
        return 'FALHA_MANUAL', "Nenhum número de loja encontrado no nome"

    caminho_arquivo_preparado = renomear_e_preparar_arquivo(nome_arquivo_completo, nome_arquivo_limpo, log_file)
    if not caminho_arquivo_preparado:
        return 'FALHA_MANUAL', "Erro ao preparar arquivo (verifique padrão da vigência)"

    cidade_loja = extrair_cidade_loja(nome_arquivo_limpo)

    contador_lojas = 0
    try:
        for numero in numeros_loja:
            numero_formatado = numero.zfill(3)

            nome_loja = montar_nome_loja(numero_formatado, cidade_loja)

            if numero_formatado in lojas_confirmadas:
                resultado_anterior = lojas_confirmadas[numero_formatado]
                log(f"[PULO] Loja {numero_formatado} já confirmada ('{resultado_anterior}') nesta execução para '{nome_arquivo_completo}'; pulando reenvio.", log_file)
                resultados_lojas[resultado_anterior].append(numero_formatado)
                continue

            resultado_upload = upload_com_verificacao_e_retry(driver, wait, numero_formatado, caminho_arquivo_preparado, log_file, nome_arquivo_completo, nome_loja)

            if resultado_upload == 'SUCESSO':
                lojas_confirmadas[numero_formatado] = 'SUCESSO'
                resultados_lojas['SUCESSO'].append(numero_formatado)
            elif resultado_upload == 'FALHA_LOJA_NAO_ENCONTRADA':
                lojas_confirmadas[numero_formatado] = 'FALHA_MANUAL'
                resultados_lojas['FALHA_MANUAL'].append(numero_formatado)
            elif resultado_upload == 'FALHA_CRIACAO_LOJA':
                lojas_confirmadas[numero_formatado] = 'FALHA_CRIACAO_LOJA'
                resultados_lojas['FALHA_CRIACAO'].append(numero_formatado)
            elif resultado_upload == 'JA_EXISTE':
                lojas_confirmadas[numero_formatado] = 'JA_EXISTE'
                resultados_lojas['JA_EXISTE'].append(numero_formatado)
            elif resultado_upload == 'UPLOAD_NAO_CONFIRMADO':
                resultados_lojas['UPLOAD_NAO_CONFIRMADO'].append(numero_formatado)
            else:
                resultados_lojas['FALHA_RECUPERAVEL'].append(numero_formatado)

            contador_lojas += 1
            if contador_lojas % LOJAS_POR_REINICIO == 0:
                log(f"[INFO] {contador_lojas} lojas processadas neste arquivo. Reinício programado do driver...", log_file, espacamento=True)
                raise ReinicioProgramadoNecessario
    except ReinicioProgramadoNecessario:
        if os.path.exists(caminho_arquivo_preparado):
            os.remove(caminho_arquivo_preparado)
        raise
    
    if os.path.exists(caminho_arquivo_preparado):
        os.remove(caminho_arquivo_preparado)

    if resultados_lojas['FALHA_MANUAL'] or resultados_lojas['UPLOAD_NAO_CONFIRMADO'] or resultados_lojas['FALHA_CRIACAO']:
        motivos = []
        if resultados_lojas['FALHA_MANUAL']:
            motivos.append(f"Lojas não encontradas: {', '.join(resultados_lojas['FALHA_MANUAL'])}")
        if resultados_lojas['FALHA_CRIACAO']:
            motivos.append(f"Falha ao criar lojas no ComuniQ: {', '.join(resultados_lojas['FALHA_CRIACAO'])}")
        if resultados_lojas['UPLOAD_NAO_CONFIRMADO']:
            motivos.append(f"Upload não confirmado após 3 tentativas: {', '.join(resultados_lojas['UPLOAD_NAO_CONFIRMADO'])}")
        return 'FALHA_MANUAL', '; '.join(motivos)
    elif resultados_lojas['FALHA_RECUPERAVEL']:
        motivo = f"Falha de sistema nas lojas: {', '.join(resultados_lojas['FALHA_RECUPERAVEL'])}"
        return 'FALHA_RECUPERAVEL', motivo
    elif resultados_lojas['SUCESSO']:
        return 'SUCESSO', "Todas as lojas processadas com sucesso"
    elif resultados_lojas['JA_EXISTE']:
        motivo = f"Arquivo já existia em todas as lojas: {', '.join(resultados_lojas['JA_EXISTE'])}"
        return 'JA_EXISTE', motivo

    else:
        return 'SUCESSO', "Todas as lojas processadas com sucesso"

def _sessao_morta(erro):
    texto = str(erro).lower()
    termos = [
        'invalid session id',
        'connection aborted',
        'max retries exceeded',
        'failed to establish a new connection',
        'connectionreseterror',
        '10061',
        'no such window',
        'target frame detached',
    ]
    return any(termo in texto for termo in termos)

def processar_arquivos(diretorio_origem, pasta_temp, log_file_path, usuario, senha, parar_event=None):
    if os.path.exists(log_file_path): os.remove(log_file_path)

    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log(f"=== INÍCIO DA AUTOMAÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===", log_file, espacamento=True)

        global DIRETORIO_ORIGEM, PASTA_TEMP, LOG_FILE_PATH
        DIRETORIO_ORIGEM = diretorio_origem; PASTA_TEMP = pasta_temp; LOG_FILE_PATH = log_file_path
        
        nome_pasta_feitos = "Feitos"
        nome_pasta_falhas = "Falhas_Requer_Acao_Manual"
        nome_pasta_ja_existe = "Ja_Existe"
        
        if os.path.basename(diretorio_origem) in [nome_pasta_feitos, nome_pasta_falhas, nome_pasta_ja_existe]:
            diretorio_raiz = os.path.dirname(diretorio_origem)
        else:
            diretorio_raiz = diretorio_origem

        pasta_feitos = os.path.join(diretorio_raiz, nome_pasta_feitos)
        pasta_falhas_manual = os.path.join(diretorio_raiz, nome_pasta_falhas)
        pasta_ja_existe = os.path.join(diretorio_raiz, nome_pasta_ja_existe)
        
        os.makedirs(pasta_feitos, exist_ok=True)
        os.makedirs(pasta_falhas_manual, exist_ok=True)
        os.makedirs(pasta_ja_existe, exist_ok=True)
        
        resultados_finais_falhas = []
        driver = None
        try:
            driver = inicializar_driver()
            if not driver: return []

            wait = WebDriverWait(driver, 60)

            if not fazer_login(driver, wait, usuario, senha):
                raise Exception("Falha no login.")

            if not navegar_para_acordos(driver, wait):
                raise Exception("Falha na navegação inicial.")

            extensoes_permitidas = ['.pdf', '.doc', '.docx']
            arquivos_para_processar = [f for f in os.listdir(diretorio_origem) if os.path.isfile(os.path.join(diretorio_origem, f)) and os.path.splitext(f)[1].lower() in extensoes_permitidas]

            log(f"Encontrados {len(arquivos_para_processar)} arquivos para processar em '{diretorio_origem}'.", log_file)

            lojas_confirmadas_por_arquivo = {}

            for nome_arquivo in arquivos_para_processar:
                if parar_event and parar_event.is_set():
                    log("[INTERRUPÇÃO] Processo interrompido.", log_file); break
                
                caminho_original = os.path.join(diretorio_origem, nome_arquivo)
                resultado, motivo = None, None
               
                tentativas_sessao = 0
                while True:
                    try:
                        resultado, motivo = processar_um_unico_arquivo(driver, wait, nome_arquivo, log_file, lojas_confirmadas_por_arquivo.setdefault(nome_arquivo, {}))
                        break
                    except ReinicioProgramadoNecessario:
                        log(f"[INFO] Executando reinício programado do driver para '{nome_arquivo}'...", log_file, espacamento=True)
                        driver, wait = _reiniciar_sessao(driver, wait, usuario, senha, log_file)
                        if not driver:
                            log("[ERRO] Falha no reinício programado do driver.", log_file)
                            resultado, motivo = 'FALHA_RECUPERAVEL', 'Falha ao reinicializar o driver (reinício programado)'
                            break
                        continue
                    except Exception as e:
                        if _sessao_morta(e):
                            tentativas_sessao += 1
                            log(f"[AVISO] Sessão do navegador encerrada ao processar '{nome_arquivo}' (tentativa {tentativas_sessao}/2). Reinicializando...", log_file)
                            driver, wait = _reiniciar_sessao(driver, wait, usuario, senha, log_file)
                            if not driver:
                                resultado, motivo = 'FALHA_RECUPERAVEL', 'Erro de sessão durante o processamento'
                                break
                            continue
                        else:
                            log(f"[ERRO FATAL] Erro inesperado ao processar '{nome_arquivo}': {e}", log_file, espacamento=True)
                            resultado, motivo = 'FALHA_RECUPERAVEL', f'Erro inesperado: {e}'
                            break 
                if resultado is None:
                    resultado, motivo = 'FALHA_RECUPERAVEL', 'Erro de sessão/inesperado durante o processamento'

                if resultado == 'SUCESSO':
                    mover_e_renomear_arquivo(caminho_original, pasta_feitos, "[FEITO]", log_file)
                elif resultado == 'FALHA_MANUAL':
                    mover_e_renomear_arquivo(caminho_original, pasta_falhas_manual, "[FALHA_MANUAL]", log_file)
                    resultados_finais_falhas.append({'arquivo': nome_arquivo, 'motivo': motivo, 'tipo': 'MANUAL'})
                elif resultado == 'FALHA_RECUPERAVEL':
                    log(f"[AVISO] Arquivo {nome_arquivo} falhou, mas será mantido para nova tentativa.", log_file)
                    resultados_finais_falhas.append({'arquivo': nome_arquivo, 'motivo': motivo, 'tipo': 'RECUPERAVEL'})
                elif resultado == 'JA_EXISTE':
                    mover_e_renomear_arquivo(caminho_original, pasta_ja_existe, "[JA_EXISTE]", log_file)
        
        except Exception as e:
            log(f"[ERRO FATAL] Ocorreu um erro geral no processo: {e}", log_file, espacamento=True)
        
        finally:
            falhas_que_exigem_acao = [f for f in resultados_finais_falhas if f.get('tipo') == 'MANUAL']
            if falhas_que_exigem_acao:
                caminho_txt_erros = os.path.join(pasta_falhas_manual, "motivo_dos_erros.txt")
                with open(caminho_txt_erros, 'w', encoding='utf-8') as f_erros:
                    f_erros.write(f"Relatório de Falhas Manuais - {datetime.now().strftime('%Y-%m-%d %H:%M')}\n" + "="*40 + "\n\n")
                    for falha in falhas_que_exigem_acao:
                        f_erros.write(f"ARQUIVO: {falha['arquivo']}\n  MOTIVO: {falha['motivo']}\n\n")
            
            log(f"=== FIM DA AUTOMAÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===", log_file, espacamento=True)
            if driver: driver.quit()
        return resultados_finais_falhas