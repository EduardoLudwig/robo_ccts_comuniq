import os
import re
import shutil
import time
import win32com.client
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException

# --- CONFIGURAÇÕES GERAIS ---
#USUARIO = "pamela.lanferdini@Quero-Quero.com.br"
#SENHA = "Quero@2025"

URL_LOGIN = "https://queroquero.mybeehome.com/login"
URL_WIKI = "https://queroquero.mybeehome.com/wiki/492/detail"

DIRETORIO_ORIGEM = r""
PASTA_TEMP = r"M:\Remuneração e Benefícios\SINDICAIS\Robô CCTS ComuniQ\temp_upload"
LOG_FILE_PATH = r"M:\Remuneração e Benefícios\SINDICAIS\Robô CCTS ComuniQ\log_historico.txt"

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
    word = None
    doc = None
    try:
        nome_base = os.path.splitext(os.path.basename(caminho_arquivo))[0]
        caminho_pdf = os.path.join(diretorio_saida, f"{nome_base}.pdf")
        caminho_arquivo_abs = os.path.abspath(caminho_arquivo)
        caminho_pdf_abs = os.path.abspath(caminho_pdf)
        log(f"[INFO] Convertendo '{os.path.basename(caminho_arquivo)}' para PDF...", log_file)
        if not os.path.exists(diretorio_saida):
            os.makedirs(diretorio_saida)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(caminho_arquivo_abs)
        wdFormatPDF = 17
        doc.SaveAs(caminho_pdf_abs, FileFormat=wdFormatPDF)
        log(f"[SUCESSO] Arquivo convertido para PDF: {caminho_pdf_abs}", log_file)
        return caminho_pdf
    except Exception as e:
        log(f"[ERRO] Falha ao converter o arquivo '{os.path.basename(caminho_arquivo)}' para PDF. Erro: {e}", log_file)
        log("[AVISO] Verifique se o Microsoft Word está instalado e se o arquivo não está corrompido.", log_file)
        return None
    finally:
        if doc:
            doc.Close(0)
        if word:
            word.Quit()

def extrair_numeros_loja(nome_arquivo):
    primeira_parte = nome_arquivo.split('_')[0]
    numeros = re.findall(r'\b\d{1,3}\b', primeira_parte)
    return numeros

def renomear_e_preparar_arquivo(nome_original, log_file):
    caminho_origem_completo = os.path.join(DIRETORIO_ORIGEM, nome_original)
    nome_sem_extensao, extensao = os.path.splitext(nome_original)
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
        log(f"[ERRO] Padrão de VIGÊNCIA (ex: _2025 ou _2025-2027) não encontrado em: {nome_original}", log_file)
        if 'caminho_pdf_convertido' in locals() and os.path.exists(caminho_pdf_convertido):
            os.remove(caminho_pdf_convertido)
        return None
    vigencia = match_vigencia.group(1)
    fim_da_vigencia = match_vigencia.end(0)
    complemento_bruto = nome_sem_extensao[fim_da_vigencia:]
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
        time.sleep(8)
        return True
    except Exception as e:
        print(f"Erro durante o login: {e}")
        return False

def navegar_para_acordos(driver, wait):
    try:
        driver.get(URL_WIKI)
        time.sleep(3)
        driver.refresh()
        time.sleep(7)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Toggle Remuneração CS']"))).click()
        print("Clicou no Toggle Remuneração CS")
        time.sleep(2)
        acordos_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Toggle Acordos e Convenções Coletivas']")))
        acordos_btn.click()
        print("Clicou no Toggle Acordos e Convenções Coletivas")
        time.sleep(5)
        driver.execute_script("arguments[0].scrollIntoView(true);", acordos_btn)
        print("Scroll realizado para a seção de Acordos.")
        return True
    except Exception as e:
        print(f"Erro ao navegar para a seção de acordos: {e}")
        return False

def upload_arquivo_para_loja(driver, wait, numero_loja_formatado, caminho_arquivo_completo, log_file):
    try:
        log(f"                                                                            ", log_file, registrar_em_arquivo=False)
        log(f"--- Iniciando upload para loja {numero_loja_formatado} ---", log_file, registrar_em_arquivo=False)
        log(f"                                                                            ", log_file, registrar_em_arquivo=False)
        loja_xpath = f"//span[contains(@class, 'mdc-button__label') and starts-with(normalize-space(.), '{numero_loja_formatado}')]"
        elemento_loja = wait.until(EC.presence_of_element_located((By.XPATH, loja_xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_loja)
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, loja_xpath))).click()
        log(f"[OK] Clicou na loja que começa com: {numero_loja_formatado}", log_file, registrar_em_arquivo=False)
        time.sleep(3)
        try:
            popup_wait = WebDriverWait(driver, 3)
            popup_cancel_button_xpath = "//button[@id='onesignal-slidedown-cancel-button']"
            botao_cancelar = popup_wait.until(EC.element_to_be_clickable((By.XPATH, popup_cancel_button_xpath)))
            log("[AVISO] Pop-up 'OneSignal' detectado! Tentando fechar...", log_file, registrar_em_arquivo=False)
            botao_cancelar.click()
            log("[OK] Pop-up fechado com sucesso.", log_file, registrar_em_arquivo=False)
        except TimeoutException:
            log("[INFO] Nenhum pop-up 'OneSignal' detectado.", log_file, registrar_em_arquivo=False)
        except Exception as e_popup:
            log(f"[AVISO] Ocorreu um erro não crítico ao tentar fechar o pop-up: {e_popup}", log_file, registrar_em_arquivo=False)
        log("[AÇÃO] Procurando o botão 'Ações'...", log_file, registrar_em_arquivo=False)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Ações']"))).click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space(text())='Editar']"))).click()
        log(f"[OK] Clicou em Ações -> Editar", log_file, registrar_em_arquivo=False)
        time.sleep(3)
        anexos = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.dz-clickable")))
        driver.execute_script("arguments[0].scrollIntoView(true);", anexos)
        print("Scroll realizado para anexos.")
        time.sleep(2)
        log("[AÇÃO] Procurando TODOS os elementos de input de arquivo na página...", log_file, registrar_em_arquivo=False)
        seletor_todos_inputs = "//input[@type='file']"
        todos_os_inputs = wait.until(EC.presence_of_all_elements_located((By.XPATH, seletor_todos_inputs)))
        if len(todos_os_inputs) < 3:
            raise FileNotFoundError(f"Esperava encontrar pelo menos 3 inputs de arquivo, mas encontrou apenas {len(todos_os_inputs)}.")
        log(f"[INFO] Encontrados {len(todos_os_inputs)} inputs. Selecionando o terceiro...", log_file, registrar_em_arquivo=False)
        input_alvo = todos_os_inputs[2]
        input_alvo.send_keys(caminho_arquivo_completo)
        log(f"[OK] Arquivo '{os.path.basename(caminho_arquivo_completo)}' enviado para o input.", log_file, registrar_em_arquivo=False)
        time.sleep(5)
        seletor_salvar = "//button[contains(normalize-space(.), 'Salvar') and contains(@class, 'btn-success')]"
        todos_botoes_salvar = driver.find_elements(By.XPATH, seletor_salvar)
        botoes_visiveis = [btn for btn in todos_botoes_salvar if btn.is_displayed() and btn.is_enabled()]
        if not botoes_visiveis:
            log("[ERRO] Nenhum botão 'Salvar' visível e habilitado encontrado.", log_file)
            return False
        botao_salvar = botoes_visiveis[0]
        driver.execute_script("arguments[0].scrollIntoView(true);", botao_salvar)
        time.sleep(1)
        botao_salvar.click()
        log(f"[OK] Botão 'Salvar' clicado com sucesso.", log_file, registrar_em_arquivo=False)
        time.sleep(10)
        log("[SUCESSO] Upload realizado com sucesso!", log_file, registrar_em_arquivo=False)
        if not navegar_para_acordos(driver, wait):
            log("[ERRO] Falha ao retornar para a seção de acordos após salvar.", log_file)
            return False
        log("[INFO] Tela de acordos recarregada com sucesso para o próximo upload.", log_file, registrar_em_arquivo=False)
        log(f"[SUCESSO FINAL] Processo para loja {numero_loja_formatado} concluído.", log_file, registrar_em_arquivo=False)
        return True
    except Exception as e:
        log(f"[ERRO GERAL] Falha no processo de upload para a loja {numero_loja_formatado}. Erro: {e}", log_file)
        driver.get(URL_WIKI)
        time.sleep(5)
        navegar_para_acordos(driver, wait)
        return False


def processar_arquivos(diretorio_origem, pasta_temp, log_file_path, usuario, senha, parar_event=None):
    if os.path.exists(log_file_path):
        os.remove(log_file_path)

    with open(log_file_path, 'w', encoding='utf-8') as log_file:
        log(f"=== INÍCIO DA AUTOMAÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===", log_file, espacamento=True)

        global DIRETORIO_ORIGEM, PASTA_TEMP, LOG_FILE_PATH
        DIRETORIO_ORIGEM = diretorio_origem
        PASTA_TEMP = pasta_temp
        LOG_FILE_PATH = log_file_path

        driver = inicializar_driver()
        if not driver: return

        wait = WebDriverWait(driver, 20)

        if not fazer_login(driver, wait, usuario, senha):
            log("[ERRO FATAL] Falha no login. Verifique se o usuário e senha estão corretos.", log_file)
            driver.quit()
            return

        if not navegar_para_acordos(driver, wait):
            log("[ERRO FATAL] Falha na navegação.", log_file)
            driver.quit()
            return

        extensoes_permitidas = ['.pdf', '.doc', '.docx']
        arquivos_para_processar = [f for f in os.listdir(DIRETORIO_ORIGEM) if os.path.splitext(f)[1].lower() in extensoes_permitidas]

        if not arquivos_para_processar:
            log("[AVISO] Nenhum arquivo com as extensões permitidas (.pdf, .doc, .docx) foi encontrado na pasta de origem.", log_file, espacamento=True)

        for nome_arquivo in arquivos_para_processar:
            if parar_event and parar_event.is_set():
                log("[INTERRUPÇÃO] Processo interrompido pelo usuário.", log_file, espacamento=True)
                break

            log(f"Processando arquivo: {nome_arquivo}", log_file, espacamento=True)

            numeros_loja = extrair_numeros_loja(nome_arquivo)
            if not numeros_loja:
                log(f"[ERRO] Nenhum número de loja encontrado no nome do arquivo. Pulando: {nome_arquivo}", log_file)
                continue

            log(f"[INFO] Lojas encontradas para este arquivo: {', '.join(numeros_loja)}", log_file)

            caminho_arquivo_preparado = renomear_e_preparar_arquivo(nome_arquivo, log_file)
            if not caminho_arquivo_preparado:
                log(f"[ERRO] Não foi possível preparar o arquivo {nome_arquivo} para upload. Pulando.", log_file)
                continue

            lojas_sucesso = []
            lojas_falha = []

            for numero in numeros_loja:
                if parar_event and parar_event.is_set():
                    log(f"[INTERRUPÇÃO] Upload interrompido durante o envio para a loja {numero.zfill(3)}.", log_file)
                    break

                numero_formatado = numero.zfill(3)
                sucesso = upload_arquivo_para_loja(driver, wait, numero_formatado, caminho_arquivo_preparado, log_file)

                if sucesso:
                    lojas_sucesso.append(numero_formatado)
                else:
                    lojas_falha.append(numero_formatado)
                    log(f"[FALHA] Upload falhou para loja {numero_formatado}.", log_file)

            if lojas_sucesso:
                log(f"[RESUMO SUCESSO] Upload concluído para as lojas: {', '.join(lojas_sucesso)} | Documento de origem: {nome_arquivo}", log_file)

            if lojas_falha:
                log(f"[RESUMO FALHA] Upload falhou para as lojas: {', '.join(lojas_falha)} | Documento de origem: {nome_arquivo}", log_file)

            if os.path.exists(caminho_arquivo_preparado):
                os.remove(caminho_arquivo_preparado)
                log(f"[INFO] Arquivo temporário '{os.path.basename(caminho_arquivo_preparado)}' removido.", log_file)

        if os.path.exists(PASTA_TEMP) and not os.listdir(PASTA_TEMP):
            try:
                os.rmdir(PASTA_TEMP)
            except OSError as e:
                log(f"[AVISO] Não foi possível remover a pasta temporária (pode estar em uso): {e}", log_file)

        log(f"=== FIM DA AUTOMAÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===", log_file, espacamento=True)
        driver.quit()