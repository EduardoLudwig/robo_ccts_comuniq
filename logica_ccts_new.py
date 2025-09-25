import time
import os
import re
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

# --- CONFIGURAÇÕES GERAIS ---
URL_LOGIN = "https://queroquero.mybeehome.com/login"
URL_WIKI = "https://queroquero.mybeehome.com/wiki/492/detail"
DIRETORIO_ORIGEM = ""
PASTA_TEMP = ""
LOG_FILE_PATH = ""

# --- SUAS FUNÇÕES ORIGINAIS (MANTIDAS INTOCADAS) ---

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

def renomear_e_preparar_arquivo(nome_original, log_file):
    caminho_origem_completo = os.path.join(DIRETORIO_ORIGEM, nome_original)
    nome_sem_extensao, extensao = os.path.splitext(nome_original)
    caminho_para_upload = caminho_origem_completo
    extensao_final = extensao

    if extensao.lower() in ['.doc', '.docx']:
        caminho_pdf_convertido = converter_doc_para_pdf(caminho_origem_completo, PASTA_TEMP, log_file)
        if not caminho_pdf_convertido: return None
        caminho_para_upload = caminho_pdf_convertido
        extensao_final = '.pdf'
        nome_sem_extensao = os.path.splitext(os.path.basename(caminho_pdf_convertido))[0]
    match_vigencia = re.search(r'_(\d{4}(?:-\d{4})?)', nome_sem_extensao)
    if not match_vigencia:
        log(f"[ERRO] Padrão de VIGÊNCIA (ex: _2025 ou _2025-2027) não encontrado em: {nome_original}", log_file)
        if 'caminho_pdf_convertido' in locals() and os.path.exists(caminho_pdf_convertido): os.remove(caminho_pdf_convertido)
        return None
    vigencia = match_vigencia.group(1)
    
    fim_da_vigencia = match_vigencia.end(0)
    complemento_bruto = nome_sem_extensao[fim_da_vigencia:]
    complemento = complemento_bruto.lstrip('_')
    novo_nome_base = f"CCT_{vigencia}"
    if complemento: novo_nome_base += f"_{complemento}"
    novo_nome_final = novo_nome_base + extensao_final
    if not os.path.exists(PASTA_TEMP): os.makedirs(PASTA_TEMP)
    caminho_destino_temporario = os.path.join(PASTA_TEMP, novo_nome_final)
    try:
        if 'caminho_pdf_convertido' in locals() and os.path.exists(caminho_para_upload):
            if os.path.exists(caminho_destino_temporario): os.remove(caminho_destino_temporario)
            shutil.move(caminho_para_upload, caminho_destino_temporario)
        else:
            shutil.copy2(caminho_para_upload, caminho_destino_temporario)
        log(f"[INFO] Arquivo preparado para upload como: {novo_nome_final}", log_file)
        return caminho_destino_temporario
    except Exception as e:
        log(f"[ERRO] Falha ao preparar arquivo temporário para {nome_original}: {e}", log_file)
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
        print("Login realizado com sucesso.")
        time.sleep(4)
        return True
    except Exception as e:
        print(f"Erro durante o login: {e}")
        return False

def navegar_para_acordos(driver, wait):
    try:
        driver.get(URL_WIKI)
        remuneracao_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Toggle Remuneração CS']")))
        time.sleep(2)
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

# --- NOVAS FUNÇÕES E FUNÇÕES ATUALIZADAS PARA A LÓGICA DE PASTAS ---

def mover_e_renomear_arquivo(caminho_arquivo_original, pasta_destino, prefixo, log_file):
    """
    NOVA FUNÇÃO: Move um arquivo para uma pasta de destino, renomeando-o com um prefixo.
    """
    try:
        nome_base_arquivo = os.path.basename(caminho_arquivo_original)
        novo_nome = f"{prefixo}_{nome_base_arquivo}"
        caminho_destino_final = os.path.join(pasta_destino, novo_nome)
        shutil.move(caminho_arquivo_original, caminho_destino_final)
        log(f"[INFO] Arquivo movido e renomeado para: {caminho_destino_final}", log_file)
    except Exception as e:
        log(f"[ERRO] Falha ao mover/renomear o arquivo {caminho_arquivo_original}. Erro: {e}", log_file)

def upload_arquivo_para_loja(driver, wait, numero_loja_formatado, caminho_arquivo_completo, log_file):
    """
    SUA FUNÇÃO DE UPLOAD, COM A MÍNIMA ALTERAÇÃO NECESSÁRIA (NOS RETORNOS).
    """
    try:
        log(f"--- Iniciando upload para loja {numero_loja_formatado}' para a loja {numero_loja_formatado} ---", log_file, registrar_em_arquivo=False)
        
        try:
            loja_xpath = f"//span[contains(@class, 'mdc-button__label') and starts-with(normalize-space(.), '{numero_loja_formatado}')]"
            elemento_loja = wait.until(EC.element_to_be_clickable((By.XPATH, loja_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_loja)
            elemento_loja.click()
            log(f"[OK] Clicou na loja que começa com: {numero_loja_formatado}", log_file, registrar_em_arquivo=False)
        except TimeoutException:
            log(f"[ERRO MANUAL] Loja {numero_loja_formatado} não encontrada na plataforma.", log_file)
            return 'FALHA_LOJA_NAO_ENCONTRADA' # <-- MUDANÇA 1

        try:
            popup_wait = WebDriverWait(driver, 3)
            popup_cancel_button_xpath = "//button[@id='onesignal-slidedown-cancel-button']"
            botao_cancelar = popup_wait.until(EC.element_to_be_clickable((By.XPATH, popup_cancel_button_xpath)))
            botao_cancelar.click()
        except TimeoutException:
            pass # Continua se não houver pop-up
        
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Ações']"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space(text())='Editar']"))).click()
        
        anexos = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.dz-clickable")))
        driver.execute_script("arguments[0].scrollIntoView(true);", anexos)
        
        input_alvo = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@type='file']")))[2]

        timeout = 30; check_interval = 1; stable_checks_required = 3
        start_time = time.time(); last_count = -1; stable_counter = 0
        while time.time() - start_time < timeout:
            current_count = len(driver.find_elements(By.XPATH, "//a[normalize-space()='Remover Arquivo']"))
            if current_count == last_count: stable_counter += 1
            else: stable_counter = 0
            if stable_counter >= stable_checks_required: break
            last_count = current_count; time.sleep(check_interval)
        else:
            log("[AVISO] Timeout! A lista de arquivos não estabilizou a tempo.", log_file)

        botoes_remover_antes = driver.find_elements(By.XPATH, "//a[normalize-space()='Remover Arquivo']")
        contagem_antes = len(botoes_remover_antes)
        log(f"[INFO] Contagem antes do upload: {contagem_antes} arquivos.")
        
        input_alvo.send_keys(caminho_arquivo_completo)
        log(f"[OK] Arquivo '{os.path.basename(caminho_arquivo_completo)}' enviado para a loja {numero_loja_formatado}.", log_file)

        try:
            wait.until(lambda d: len(d.find_elements(By.XPATH, "//a[normalize-space()='Remover Arquivo']")) == contagem_antes + 1)
        except TimeoutException:
            raise Exception("Falha ao verificar o upload do anexo na interface.")

        # input("Veficar se o upload esta ok:") # Mantendo sua linha de debug, comente se não precisar
        
        seletor_salvar = "//button[contains(normalize-space(.), 'Salvar') and contains(@class, 'btn-success')]"
        todos_botoes_salvar = driver.find_elements(By.XPATH, seletor_salvar)
        botoes_visiveis = [btn for btn in todos_botoes_salvar if btn.is_displayed() and btn.is_enabled()]
        if not botoes_visiveis: raise Exception("Nenhum botão 'Salvar' visível encontrado.")
        botao_salvar = botoes_visiveis[0]
        driver.execute_script("arguments[0].scrollIntoView(true);", botao_salvar)
        time.sleep(1)
        botao_salvar.click()
        
        time.sleep(3)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Ações']")))
        
        if not navegar_para_acordos(driver, wait):
            raise Exception("Falha ao retornar para a seção de acordos após salvar.")
        
        log(f"[SUCESSO FINAL] Processo para loja {numero_loja_formatado} concluído.", log_file, registrar_em_arquivo=False)
        return 'SUCESSO' # <-- MUDANÇA 2

    except Exception as e:
        log(f"[ERRO GERAL] Falha no processo para a loja {numero_loja_formatado}. Erro: {e}", log_file)
        driver.get(URL_WIKI)
        time.sleep(5)
        navegar_para_acordos(driver, wait)
        return 'FALHA_RECUPERAVEL' # <-- MUDANÇA 3

def processar_um_unico_arquivo(driver, wait, nome_arquivo, log_file):
    """
    NOVA FUNÇÃO: Concentra a lógica para processar um único arquivo e retorna seu status final.
    """
    log(f"Processando arquivo: {nome_arquivo}", log_file, espacamento=True)
    resultados_lojas = {'SUCESSO': [], 'FALHA_MANUAL': [], 'FALHA_RECUPERAVEL': []}

    numeros_loja = extrair_numeros_loja(nome_arquivo)
    if not numeros_loja:
        return 'FALHA_MANUAL', "Nenhum número de loja encontrado no nome"

    caminho_arquivo_preparado = renomear_e_preparar_arquivo(nome_arquivo, log_file)
    if not caminho_arquivo_preparado:
        return 'FALHA_MANUAL', "Erro ao preparar arquivo (verifique padrão da vigência)"

    for numero in numeros_loja:
        numero_formatado = numero.zfill(3)
        resultado_upload = upload_arquivo_para_loja(driver, wait, numero_formatado, caminho_arquivo_preparado, log_file)
        
        if resultado_upload == 'SUCESSO':
            resultados_lojas['SUCESSO'].append(numero_formatado)
        elif resultado_upload == 'FALHA_LOJA_NAO_ENCONTRADA':
            resultados_lojas['FALHA_MANUAL'].append(numero_formatado)
        else:
            resultados_lojas['FALHA_RECUPERAVEL'].append(numero_formatado)
    
    if os.path.exists(caminho_arquivo_preparado):
        os.remove(caminho_arquivo_preparado)

    if resultados_lojas['FALHA_MANUAL']:
        motivo = f"Lojas não encontradas: {', '.join(resultados_lojas['FALHA_MANUAL'])}"
        return 'FALHA_MANUAL', motivo
    elif resultados_lojas['FALHA_RECUPERAVEL']:
        motivo = f"Falha de sistema nas lojas: {', '.join(resultados_lojas['FALHA_RECUPERAVEL'])}"
        return 'FALHA_RECUPERAVEL', motivo
    else:
        return 'SUCESSO', "Todas as lojas processadas com sucesso"

def processar_arquivos(diretorio_origem, pasta_temp, log_file_path, usuario, senha, parar_event=None):
    """
    FUNÇÃO PRINCIPAL RECONSTRUÍDA: Agora ela gerencia o processo geral.
    """
    if os.path.exists(log_file_path): os.remove(log_file_path)

    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log(f"=== INÍCIO DA AUTOMAÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===", log_file, espacamento=True)

        global DIRETORIO_ORIGEM, PASTA_TEMP, LOG_FILE_PATH
        DIRETORIO_ORIGEM = diretorio_origem; PASTA_TEMP = pasta_temp; LOG_FILE_PATH = log_file_path

        pasta_feitos = os.path.join(diretorio_origem, "Feitos")
        pasta_falhas_manual = os.path.join(diretorio_origem, "Falhas_Requer_Acao_Manual")
        os.makedirs(pasta_feitos, exist_ok=True)
        os.makedirs(pasta_falhas_manual, exist_ok=True)
        
        resultados_finais_falhas = []
        driver = None
        try:
            driver = inicializar_driver()
            if not driver: return []

            wait = WebDriverWait(driver, 120)

            if not fazer_login(driver, wait, usuario, senha):
                raise Exception("Falha no login.")

            if not navegar_para_acordos(driver, wait):
                raise Exception("Falha na navegação inicial.")

            extensoes_permitidas = ['.pdf', '.doc', '.docx']
            arquivos_para_processar = [f for f in os.listdir(DIRETORIO_ORIGEM) if os.path.isfile(os.path.join(DIRETORIO_ORIGEM, f)) and os.path.splitext(f)[1].lower() in extensoes_permitidas]

            log(f"Encontrados {len(arquivos_para_processar)} arquivos para processar.", log_file)
            
            for nome_arquivo in arquivos_para_processar:
                if parar_event and parar_event.is_set():
                    log("[INTERRUPÇÃO] Processo interrompido.", log_file); break
                
                caminho_original = os.path.join(diretorio_origem, nome_arquivo)
                resultado, motivo = processar_um_unico_arquivo(driver, wait, nome_arquivo, log_file)

                if resultado == 'SUCESSO':
                    mover_e_renomear_arquivo(caminho_original, pasta_feitos, "[FEITO]", log_file)
                elif resultado == 'FALHA_MANUAL':
                    mover_e_renomear_arquivo(caminho_original, pasta_falhas_manual, "[FALHA_MANUAL]", log_file)
                    resultados_finais_falhas.append({'arquivo': nome_arquivo, 'motivo': motivo, 'tipo': 'MANUAL'})
                elif resultado == 'FALHA_RECUPERAVEL':
                    log(f"[AVISO] Arquivo {nome_arquivo} falhou, mas será mantido para nova tentativa.", log_file)
                    resultados_finais_falhas.append({'arquivo': nome_arquivo, 'motivo': motivo, 'tipo': 'RECUPERAVEL'})
        
        except Exception as e:
            log(f"[ERRO FATAL] Ocorreu um erro geral no processo: {e}", log_file, espacamento=True)
        
        finally:
            falhas_que_exigem_acao = [f for f in resultados_finais_falhas if f.get('tipo') == 'MANUAL']
            
            if falhas_que_exigem_acao:
                caminho_txt_erros = os.path.join(pasta_falhas_manual, "motivo_dos_erros.txt")
                log(f"[INFO] Gerando relatório de falhas manuais em: {caminho_txt_erros}", log_file)
                with open(caminho_txt_erros, 'w', encoding='utf-8') as f_erros:
                    f_erros.write(f"Relatório de Falhas Manuais - {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
                    f_erros.write("="*40 + "\n\n")
                    for falha in falhas_que_exigem_acao:
                        f_erros.write(f"ARQUIVO: {falha['arquivo']}\n")
                        f_erros.write(f"  MOTIVO: {falha['motivo']}\n\n")
            
            log(f"=== FIM DA AUTOMAÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===", log_file, espacamento=True)
            if driver: driver.quit()
            return resultados_finais_falhas