import time
import os
import re
import shutil
import configparser  # leitura/gravação do config.ini
import win32com.client
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from enviar_email import enviar_relatorio, montar_relatorio  # relatório simples e envio por email ao fim da execução

APP_CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".RoboComuniq")  # pasta de config fora do projeto: C:\Users\<user>\.RoboComuniq (funciona igual no .exe e no script)
os.makedirs(APP_CONFIG_DIR, exist_ok=True)  # cria a pasta se não existir
config_file_path = os.path.join(APP_CONFIG_DIR, "config.ini")  # caminho do config.ini (mesmo usado pelo enviar_email.py)

################################################################

def _ler_config():  # lê o config.ini e devolve o ConfigParser (vazio se o arquivo não existir)
    config = configparser.ConfigParser(interpolation=None)  # interpolation=None para senhas com '%' não quebrarem a leitura
    if os.path.exists(config_file_path):  # só lê se o arquivo já existir
        config.read(config_file_path, encoding='utf-8')  # carrega em utf-8 (acentos)
    return config  # devolve o objeto de configuração

def _gravar_config(config):  # grava o ConfigParser no config.ini
    with open(config_file_path, 'w', encoding='utf-8') as configfile:  # sobrescreve o arquivo com o conteúdo atual
        config.write(configfile)  # escreve as seções/chaves

def garantir_config_inicial():
    # Cria o config.ini na primeira execução com as seções [Login], [Settings] e [Email].
    # Se o arquivo já existir, NÃO sobrescreve nada: só adiciona o que estiver faltando.
    config = _ler_config()  # conteúdo atual (vazio se o arquivo não existir)
    alterado = False  # só regrava o arquivo se algo for adicionado

    if 'Login' not in config:  # seção de credenciais do ComuniQ ausente
        config['Login'] = {'Usuario': '', 'Senha': ''}  # cria vazia (preenchida pelo botão "Editar Login")
        alterado = True  # precisa gravar

    if 'Settings' not in config:  # seção de configurações ausente
        config['Settings'] = {'CriarLojaAutomaticamente': 'False'}  # por padrão não cria loja automaticamente
        alterado = True  # precisa gravar

    if 'PastaBase' not in config['Settings']:  # usada no --auto-start para achar a pasta do mês
        config['Settings']['PastaBase'] = r'M:\Remuneração e Benefícios\SINDICAIS'
        alterado = True

    if 'Email' not in config:  # seção de email ausente
        config['Email'] = {}  # cria vazia; as chaves são adicionadas abaixo
        alterado = True  # precisa gravar

    email_cfg = config['Email']  # atalho para a seção [Email]
    for chave in ('EmailRemetente', 'SenhaApp', 'EmailDestinatario'):  # chaves obrigatórias do envio do relatório
        if chave not in email_cfg:  # chave ausente (não mexe nas que já existem)
            email_cfg[chave] = ''  # adiciona vazia para o usuário preencher
            alterado = True  # precisa gravar

    if alterado:  # algo foi adicionado
        _gravar_config(config)  # salva o config.ini

def carregar_credenciais():  # devolve (usuario, senha) do config.ini
    config = _ler_config()  # lê o arquivo
    return config.get('Login', 'Usuario', fallback=''), config.get('Login', 'Senha', fallback='')  # vazio se não configurado

def salvar_credenciais(usuario, senha):  # grava usuário/senha no config.ini preservando as outras seções
    config = _ler_config()  # lê o conteúdo atual para não perder [Settings]/[Email]
    if 'Login' not in config:  # seção ausente
        config['Login'] = {}  # cria a seção
    config['Login']['Usuario'] = usuario  # atualiza o usuário
    config['Login']['Senha'] = senha  # atualiza a senha
    _gravar_config(config)  # salva o arquivo

garantir_config_inicial()  # garante o config.ini assim que o módulo é importado (igual ao robô do mediador)

################################################################

LOJAS_POR_REINICIO = 20  # depois de tantas lojas, reinicia o navegador ENTRE um arquivo e outro (libera memória)
MAX_REINICIOS_SESSAO = 3  # reinícios por arquivo quando a sessão do navegador cai; depois disso o arquivo vai para Reprocessar/
MAX_TENTATIVAS_REINICIO = 3  # tentativas para abrir o navegador já logado; se todas falharem, a execução para
CLIQUES_BLOQUEADOS_PARA_REINICIAR = 3  # cliques bloqueados seguidos indicam página travada; reinicia o navegador na hora
lojas_desde_ultimo_reinicio = 0  # lojas processadas desde que o navegador atual foi aberto (conta a execução inteira)
lojas_criadas_na_execucao = []  # lojas criadas pelo robô nesta execução, ex.: {'numero': '999', 'nome': '999 - TESTE'}
cliques_bloqueados_seguidos = 0  # cliques via JavaScript por causa de overlay; zera a cada clique normal bem-sucedido

class SessaoTravada(Exception): # Para diferenciar erro comum de um tipo próprio. fecha e reabre o chrome imediatamente
    # Página travada (overlay bloqueando cliques) ou loja que não abre: o navegador precisa ser reiniciado.
    # É separada de erro comum para o processamento do arquivo reiniciar a sessão e retomar só o que faltou
    pass

URL_LOGIN = "https://queroquero.mybeehome.com/login"
URL_WIKI = "https://queroquero.mybeehome.com/wiki/492/detail"
# Página "Remuneração (C&S)", onde se cria a loja. Só serve para CONFERIR que o robô chegou nela:
# abrir esse endereço direto dá "Acesso negado" ao criar a loja (o site exige chegar navegando por dentro dele)
CAMINHO_CATEGORIA_REMUNERACAO = "/wiki/130/category"
SECAO_ACORDOS = "Acordos e Convenções Coletivas"
DIRETORIO_ORIGEM = ""
PASTA_TEMP = ""
LOG_FILE_PATH = ""

def log(message, log_file, espacamento=False, registrar_em_arquivo=True):
    try:
        horario = datetime.now().strftime("%H:%M:%S")  # horário em cada linha: mostra onde estão as demoras
        if espacamento:
            log_message = "\n" + "-" * 80 + "\n" + f"{horario} {message.strip()}" + "\n"
        else:
            log_message = f"{horario} {message.strip()}"
        print(log_message)
        if registrar_em_arquivo:
            log_file.write(log_message + "\n")
            log_file.flush()
    except Exception as e:
        print(f"[ERRO] Falha ao escrever no log: {e}")

def resumir_erro(e):
    # Tipo do erro + 1ª linha da mensagem, sem o rastro técnico do Selenium ("Stacktrace", "No symbol"...).
    # Ex.: "TimeoutException" (tempo esgotado) ou "ElementClickInterceptedException: element click intercepted..."
    primeira_linha = str(e).split("\n")[0].replace("Message:", "").strip()
    if primeira_linha:
        return f"{type(e).__name__}: {primeira_linha}"
    return type(e).__name__

MAX_LOGS_GUARDADOS = 30

def arquivar_log_anterior(log_file_path):
    # Em vez de apagar o log da execução anterior, guarda uma cópia na pasta "Logs" ao lado dele,
    # mantendo só os 30 mais recentes. Assim dá para investigar problemas de execuções passadas
    if not os.path.exists(log_file_path):
        return
    try:
        pasta_logs = os.path.join(os.path.dirname(log_file_path), "Logs")
        os.makedirs(pasta_logs, exist_ok=True)
        data_do_log = datetime.fromtimestamp(os.path.getmtime(log_file_path)).strftime("%Y-%m-%d_%H-%M-%S")
        shutil.move(log_file_path, os.path.join(pasta_logs, f"log_{data_do_log}.txt"))

        # O nome começa pela data (ano-mês-dia), então a ordem alfabética é a ordem cronológica
        logs = sorted(os.listdir(pasta_logs))
        while len(logs) > MAX_LOGS_GUARDADOS:
            os.remove(os.path.join(pasta_logs, logs.pop(0)))  # apaga o mais antigo
    except Exception as e:
        print(f"[AVISO] Não foi possível guardar o log anterior: {e}")
        if os.path.exists(log_file_path):
            os.remove(log_file_path)

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
    try:  # valor inválido no ini (ex.: "talvez") levanta ValueError
        return _ler_config().getboolean('Settings', 'CriarLojaAutomaticamente', fallback=False)  # lê do config.ini (aceita True/False, sim: 1/0, yes/no, on/off)
    except ValueError:  # valor não reconhecido como booleano
        return False  # por segurança, não cria loja

def extrair_cidade_loja(nome_arquivo_limpo):
    partes = nome_arquivo_limpo.split('_')
    if len(partes) < 2:
        return None
    cidade = partes[1].strip()  # strip: "_ TESTE" vira "TESTE" (evita "999 -  TESTE" com espaço duplo)
    if not cidade:
        return None
    return cidade

def montar_nome_loja(numero_formatado, cidade):
    return f"{numero_formatado} - {cidade}" if cidade else numero_formatado

def encontrar_pasta_mes(ano, mes):
    # Acha <PastaBase>\Análises sindicatos lojas {ano}\{mes:02d}_* e devolve (caminho, None) ou (None, motivo).
    # Ignora acento/maiúsculas (até 2025 a pasta era "Analises") e usa só o número do mês (evita MARÇO x MARCO)
    pasta_base = _ler_config().get('Settings', 'PastaBase', fallback='')
    if not os.path.isdir(pasta_base):
        return None, f"PastaBase não encontrada: '{pasta_base}'"

    nome_pasta_ano = f"analises sindicatos lojas {ano}"
    pastas_ano = [d for d in os.listdir(pasta_base) if d.lower().replace('á', 'a') == nome_pasta_ano and os.path.isdir(os.path.join(pasta_base, d))]
    if not pastas_ano:
        return None, f"Pasta 'Análises sindicatos lojas {ano}' não encontrada em '{pasta_base}'"
    pasta_ano = os.path.join(pasta_base, pastas_ano[0])

    pastas_mes = [d for d in os.listdir(pasta_ano) if d.startswith(f"{mes:02d}_") and os.path.isdir(os.path.join(pasta_ano, d))]
    if not pastas_mes:
        return None, f"Pasta do mês '{mes:02d}_...' não encontrada em '{pasta_ano}'"
    return os.path.join(pasta_ano, pastas_mes[0]), None

def nome_no_padrao(nome_arquivo):
    # Mesmo critério do processamento: números de loja antes do 1º '_' e vigência '_AAAA' ou '_AAAA-AAAA'
    if not extrair_numeros_loja(nome_arquivo):
        return False
    nome_sem_extensao = os.path.splitext(nome_arquivo)[0]
    if not re.search(r'_(\d{4}(?:-\d{4})?)', nome_sem_extensao):
        return False
    return True

def listar_arquivos_para_upload(pasta, log_file, fora_do_padrao):
    # Só arquivos soltos (subpastas são ignoradas) .pdf/.doc/.docx ainda não marcados como [FEITO]/[JA_EXISTE].
    # Os fora do padrão de nome não são enviados nem movidos: ficam onde estão e entram na lista
    # fora_do_padrao, que aparece no relatório para a pessoa saber que precisa renomear
    arquivos = []
    for f in os.listdir(pasta):
        if not os.path.isfile(os.path.join(pasta, f)) or os.path.splitext(f)[1].lower() not in ('.pdf', '.doc', '.docx'):
            continue
        if re.search(r'_(?:\[FEITO\]|\[JA_EXISTE\])\.', f):
            continue
        if not nome_no_padrao(f):
            log(f"[IGNORADO] '{f}' está fora do padrão de nome; não será enviado.", log_file)
            fora_do_padrao.append(f)
            continue
        arquivos.append(f)
    return arquivos

def endereco_atual(driver):
    # Endereço da página atual sem o "?..." do final e sem a "/" final. Ex.: "https://.../wiki"
    return driver.current_url.split("?")[0].rstrip("/")

def procurar_rolando(driver, xpath, tentativas=10):
    # Procura um elemento rolando a página aos poucos até o fim (para listas que só aparecem conforme desce).
    # Devolve o elemento, ou None se não apareceu depois de todas as tentativas (1 por segundo)
    for _ in range(tentativas):
        elementos = driver.find_elements(By.XPATH, xpath)
        if elementos:
            return elementos[0]

        # Rola até o último título de categoria (funciona mesmo se quem rola é uma parte da página)
        # e também a janela inteira
        titulos = driver.find_elements(By.TAG_NAME, "h6")
        if titulos:
            driver.execute_script("arguments[0].scrollIntoView({block: 'end'});", titulos[-1])
        driver.execute_script("window.scrollBy(0, 600);")
        time.sleep(1)
    return None

# Navega por cliques até a página da categoria "Remuneração (C&S)", onde fica o botão de criar a página da loja.
# Precisa ser por cliques: entrando direto pelo endereço, o site nega o acesso ao formulário de criação.
# Em cada passo, espera a página certa carregar antes de seguir; sem isso, o robô podia achar
# o "Ver mais" da página anterior e clicar no lugar errado
def navegar_para_categoria_acordos(driver, wait, log_file):
    log("[CRIACAO] Navegando até a página de categorias de Acordos...", log_file)

    # Passo 1: menu lateral → "Biblioteca do Conhecimento" e espera o endereço virar ".../wiki"
    link_biblioteca = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//a[@href='/wiki']//span[normalize-space(.)='Biblioteca do Conhecimento']/ancestor::a")
    ))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link_biblioteca)
    clicar_fechando_popup(driver, link_biblioteca, log_file)
    WebDriverWait(driver, 20).until(lambda d: endereco_atual(d).endswith("/wiki"))
    log("[CRIACAO] Passo 1/5: Abriu a 'Biblioteca do Conhecimento'.", log_file)

    # Passo 2: "Ver mais" (é o único da página). Espera ficar clicável, rola até ele e espera 2s
    # para a Biblioteca terminar de carregar: clicando cedo demais, a lista de categorias não abre
    botao_ver_mais = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
        (By.XPATH, "//a[normalize-space(.)='Ver mais']")
    ))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_ver_mais)
    time.sleep(2)
    clicar_fechando_popup(driver, botao_ver_mais, log_file)  # clique normal, como uma pessoa
    log("[CRIACAO] Passo 2/5: Clicou em 'Ver mais'.", log_file)

    # Passo 3: procura "Remuneração (C&S)" rolando a página até o fim
    texto_remuneracao = procurar_rolando(driver, "//h6[normalize-space(.)='Remuneração (C&S)']")
    if not texto_remuneracao:
        # Registra o que apareceu, para saber se o "Ver mais" não abriu nada ou se o nome está diferente
        titulos = []
        for titulo in driver.find_elements(By.TAG_NAME, "h6"):
            if titulo.text.strip():
                titulos.append(titulo.text.strip())
        log(f"[DEBUG] {len(titulos)} título(s) de categoria na página: {titulos[:20]}", log_file)
        raise Exception("'Remuneração (C&S)' não apareceu depois de clicar em 'Ver mais'.")
    log("[CRIACAO] Encontrou 'Remuneração (C&S)'.", log_file)
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", texto_remuneracao)
    driver.execute_script("arguments[0].click();", texto_remuneracao)
    WebDriverWait(driver, 20).until(lambda d: endereco_atual(d).endswith(CAMINHO_CATEGORIA_REMUNERACAO))
    log("[CRIACAO] Passo 3/5: Abriu a categoria 'Remuneração (C&S)'.", log_file)

    # Passo 4: toggle "Remuneração CS"
    toggle_remuneracao = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[@aria-label='Toggle Remuneração CS']")
    ))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", toggle_remuneracao)
    clicar_fechando_popup(driver, toggle_remuneracao, log_file)
    log("[CRIACAO] Passo 4/5: Expandiu toggle 'Remuneração CS'.", log_file)

    # Passo 5: toggle "Acordos e Convenções Coletivas"
    toggle_acordos = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[@aria-label='Toggle Acordos e Convenções Coletivas']")
    ))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", toggle_acordos)
    clicar_fechando_popup(driver, toggle_acordos, log_file)
    log("[CRIACAO] Passo 5/5: Expandiu toggle 'Acordos e Convenções Coletivas'.", log_file)

    return True


def criar_loja_no_comuniq(driver, wait, nome_loja, log_file):
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
        clicar_fechando_popup(driver, botao_criar, log_file)
        log("[CRIACAO] Clicou no botão '+' (criar página) da seção de Acordos.", log_file)

        wait.until(EC.visibility_of_element_located((By.XPATH, "//h5[normalize-space()='Escolha o modelo da página']")))
        botao_em_branco = wait.until(EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='Em branco']/ancestor::div[contains(@class,'card')][1]//button")))
        clicar_fechando_popup(driver, botao_em_branco, log_file)

        campo_titulo = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@formcontrolname='title']")))
        campo_titulo.clear(); campo_titulo.send_keys(nome_loja)

        campo_descricao = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@formcontrolname='description']")))
        campo_descricao.clear(); campo_descricao.send_keys(nome_loja)

        campo_texto = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'fr-element') and @contenteditable='true']")))
        clicar_fechando_popup(driver, campo_texto, log_file)
        campo_texto.send_keys(nome_loja)

        botao_salvar = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'footer-bar')]//button[contains(@class,'btn-success')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", botao_salvar)
        clicar_fechando_popup(driver, botao_salvar, log_file)
        log(f"[DEBUG] Página '{nome_loja}' salva; aguardando aparecimento na lista de acordos...", log_file)

        if not navegar_para_acordos(driver, wait, log_file):
            return 'FALHA_CRIACAO_LOJA'

        confirmacao_xpath = f"//span[contains(@class, 'mdc-button__label') and contains(normalize-space(.), '{nome_loja}')]"
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, confirmacao_xpath)))
        except TimeoutException:
            log(f"[ERRO] Página '{nome_loja}' salva, mas o botão da loja não apareceu na lista de acordos.", log_file)
            return 'FALHA_CRIACAO_LOJA'

        log(f"[SUCESSO] Página da loja '{nome_loja}' criada no ComuniQ.", log_file)
        # Registra para o relatório (mesmo que o upload do arquivo falhe depois, a página já existe)
        numero_loja = nome_loja.split(" - ")[0]
        lojas_criadas_na_execucao.append({'numero': numero_loja, 'nome': nome_loja})
        return 'OK'
    except Exception as e:
        log(f"[ERRO] Falha ao criar a loja '{nome_loja}'. Erro: {resumir_erro(e)}", log_file)
        return 'FALHA_CRIACAO_LOJA'

def normalizar(nome):
    n = nome.strip().lower()
    n = re.sub(r'\s*\(\d+\)\s*$', '', n)
    if n.endswith('.pdf'):
        n = n[:-4]
    n = re.sub(r'[\s_-]+', '_', n)
    return n

def renomear_e_preparar_arquivo(nome_real_com_prefixo, nome_limpo_para_processar, log_file, diretorio_override=None):
    caminho_origem_completo = os.path.join(diretorio_override or DIRETORIO_ORIGEM, nome_real_com_prefixo)  # [MUDANÇA - Reorganização de pastas]
    
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

def navegar_para_acordos(driver, wait, log_file=None):
    try:
        driver.get(URL_WIKI)
        remuneracao_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Toggle Remuneração CS']")))
        clicar_fechando_popup(driver, remuneracao_toggle, log_file)
        print("Clicou no Toggle Remuneração CS")
        acordos_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Toggle Acordos e Convenções Coletivas']")))
        clicar_fechando_popup(driver, acordos_toggle, log_file)
        print("Clicou no Toggle Acordos e Convenções Coletivas")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", acordos_toggle)
        print("Scroll realizado para a seção de Acordos.")
        return True
    except Exception as e:
        if log_file:
            log(f"[DEBUG] Erro ao abrir a página de Acordos: {resumir_erro(e)}", log_file)
        else:
            print(f"Erro ao navegar para a seção de acordos: {e}")
        return False

def mover_e_renomear_arquivo(caminho_arquivo_original, pasta_destino, sufixo_tag, log_file):
    try:
        nome_base_arquivo = os.path.basename(caminho_arquivo_original)
        nome_sem_ext, extensao = os.path.splitext(nome_base_arquivo)

        for suf in ("_[FEITO]", "_[JA_EXISTE]"):
            if nome_sem_ext.endswith(suf):
                nome_sem_ext = nome_sem_ext[:-len(suf)]
                break

        novo_nome = f"{nome_sem_ext}_{sufixo_tag}{extensao}"
        caminho_destino_final = os.path.join(pasta_destino, novo_nome)
        shutil.move(caminho_arquivo_original, caminho_destino_final)
        log(f"[INFO] Arquivo movido e renomeado para: {caminho_destino_final}", log_file)
    except Exception as e:
        log(f"[ERRO] Falha ao mover/renomear o arquivo {caminho_arquivo_original}. Erro: {e}", log_file)

def fechar_popup_do_site(driver, log_file):
    # Pop-up (modal) do ComuniQ que abre por cima da página e bloqueia os cliques.
    # O conhecido é o "Habilitar Notificações", fechado pelo botão "Agora não"
    popups = driver.find_elements(By.CSS_SELECTOR, "div.modal.show")
    if not popups:
        return

    popup = popups[0]
    texto_popup = popup.text.replace("\n", " | ")
    log(f"[AVISO] Pop-up do site aberto. Texto: {texto_popup}", log_file)

    botoes_agora_nao = popup.find_elements(By.XPATH, ".//*[normalize-space(text())='Agora não']")
    if botoes_agora_nao:
        botoes_agora_nao[0].click()
    else:
        # Outro pop-up, que ainda não conhecemos: tenta fechar com ESC
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
    time.sleep(1)

    # Se o pop-up continuar aberto depois do "Agora não", tenta o ESC
    if driver.find_elements(By.CSS_SELECTOR, "div.modal.show"):
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        time.sleep(1)

def clicar_fechando_popup(driver, elemento, log_file):
    # Clica no elemento; se um pop-up do site estiver na frente, fecha e clica de novo.
    # Clique normal que funciona zera a contagem; fallback via JavaScript conta como bloqueio seguido.
    global cliques_bloqueados_seguidos
    try:
        elemento.click()
        cliques_bloqueados_seguidos = 0
    except ElementClickInterceptedException:
        fechar_popup_do_site(driver, log_file)
        try:
            elemento.click()
            cliques_bloqueados_seguidos = 0
        except ElementClickInterceptedException:
            # Continua bloqueado por outro elemento da página (não é o pop-up de notificações):
            # clica via JavaScript, que ignora o que estiver por cima
            cliques_bloqueados_seguidos += 1
            log(f"[AVISO] Clique bloqueado por outro elemento da página ({cliques_bloqueados_seguidos}/{CLIQUES_BLOQUEADOS_PARA_REINICIAR}); clicando via JavaScript.", log_file)
            driver.execute_script("arguments[0].click();", elemento)

def esperar_lista_de_lojas(wait):
    # Espera a lista de Acordos carregar: pelo menos um botão cujo texto começa com 3 dígitos (ex.: "013 - SAO BORJA").
    # translate(...) apaga os dígitos dos 3 primeiros caracteres; se não sobrar nada, os 3 eram dígitos.
    # Se a lista não carregar (até 60s), levanta TimeoutException: o upload falha e é tentado de novo,
    # em vez de o robô achar que a loja não existe e criar uma duplicada
    xpath_qualquer_loja = (
        "//span[contains(@class, 'mdc-button__label')"
        " and string-length(normalize-space(.)) > 3"
        " and translate(substring(normalize-space(.), 1, 3), '0123456789', '') = '']"
    )
    wait.until(EC.presence_of_element_located((By.XPATH, xpath_qualquer_loja)))

def verificar_sessao_travada(log_file, contexto):
    # Se os cliques seguidos foram bloqueados por um overlay, aborta já: reiniciar o navegador é mais rápido
    # do que esperar o WebDriverWait (até 60s) e falhar de novo na mesma página travada
    if cliques_bloqueados_seguidos >= CLIQUES_BLOQUEADOS_PARA_REINICIAR:
        log(f"[AVISO] {contexto}: {cliques_bloqueados_seguidos} cliques bloqueados seguidos. Reiniciando o navegador...", log_file)
        raise SessaoTravada(f"Página travada ({contexto}).")

def abrir_edicao_da_loja(driver, wait, numero_loja_formatado, nome_loja=None, log_file=None):
    # Devolve 'OK' (loja existente aberta para edição), 'CRIADA' (loja criada agora e aberta para edição),
    # 'LOJA_NAO_ENCONTRADA' ou 'FALHA_CRIACAO_LOJA'
    loja_criada = False

    # Com a lista já carregada, a loja aparece na hora se existir: não precisa esperar 60s por ela
    esperar_lista_de_lojas(wait)

    try:
        loja_xpath = f"//span[contains(@class, 'mdc-button__label') and starts-with(normalize-space(.), '{numero_loja_formatado}')]"
        elemento_loja = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, loja_xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_loja)
        clicar_fechando_popup(driver, elemento_loja, log_file)
        verificar_sessao_travada(log_file, f"ao abrir a loja {numero_loja_formatado}")
    except TimeoutException:
        if not (carregar_flag_criar_loja() and nome_loja and log_file):
            return 'LOJA_NAO_ENCONTRADA'

        loja_xpath_ampla = f"//span[contains(@class, 'mdc-button__label') and contains(normalize-space(.), '{numero_loja_formatado}')]"
        try:
            elemento_loja = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, loja_xpath_ampla)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_loja)
            clicar_fechando_popup(driver, elemento_loja, log_file)
            verificar_sessao_travada(log_file, f"ao abrir a loja {numero_loja_formatado}")
        except TimeoutException:
            if criar_loja_no_comuniq(driver, wait, nome_loja, log_file) != 'OK':
                return 'FALHA_CRIACAO_LOJA'
            try:
                loja_xpath_nova = f"//span[contains(@class, 'mdc-button__label') and contains(normalize-space(.), '{nome_loja}')]"
                elemento_loja = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, loja_xpath_nova)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_loja)
                clicar_fechando_popup(driver, elemento_loja, log_file)
                verificar_sessao_travada(log_file, f"ao abrir a loja {numero_loja_formatado}")
                loja_criada = True
            except TimeoutException:
                log(f"[ERRO] Loja '{nome_loja}' criada, mas o botão não ficou clicável.", log_file)
                return 'FALHA_CRIACAO_LOJA'

    try:
        popup_wait = WebDriverWait(driver, 3)
        botao_cancelar = popup_wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='onesignal-slidedown-cancel-button']")))
        clicar_fechando_popup(driver, botao_cancelar, log_file)
    except TimeoutException:
        pass

    botao_acoes = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Ações']")))
    clicar_fechando_popup(driver, botao_acoes, log_file)
    verificar_sessao_travada(log_file, f"ao abrir as Ações da loja {numero_loja_formatado}")
    link_editar = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space(text())='Editar']")))
    clicar_fechando_popup(driver, link_editar, log_file)
    verificar_sessao_travada(log_file, f"ao editar a loja {numero_loja_formatado}")

    anexos = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.dz-clickable")))
    driver.execute_script("arguments[0].scrollIntoView(true);", anexos)

    wait.until(lambda d: len(d.find_elements(By.XPATH, "//input[@type='file']")) >= 3)

    if loja_criada:
        return 'CRIADA'
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
    global cliques_bloqueados_seguidos

    def _sessao_suspeita():
        return cliques_bloqueados_seguidos >= CLIQUES_BLOQUEADOS_PARA_REINICIAR

    try:
        log(f"--- Iniciando upload de '{nome_arquivo_original}' para a loja {numero_loja_formatado} ---", log_file, registrar_em_arquivo=False)

        if _sessao_suspeita():
            log(f"[AVISO] Loja {numero_loja_formatado}: página já vinha com {cliques_bloqueados_seguidos} cliques bloqueados seguidos. Reiniciando o navegador...", log_file)
            raise SessaoTravada(f"Página travada antes de abrir a loja {numero_loja_formatado}.")
        
        status_loja = abrir_edicao_da_loja(driver, wait, numero_loja_formatado, nome_loja, log_file)
        if _sessao_suspeita():
            log(f"[AVISO] Loja {numero_loja_formatado}: {cliques_bloqueados_seguidos} cliques bloqueados seguidos ao abrir a loja. Reiniciando o navegador...", log_file)
            raise SessaoTravada(f"Página travada (cliques bloqueados) ao abrir a loja {numero_loja_formatado}.")
        if status_loja == 'LOJA_NAO_ENCONTRADA':
            log(f"[ERRO MANUAL] Loja {numero_loja_formatado} não encontrada para o arquivo '{nome_arquivo_original}'.", log_file)
            return 'FALHA_LOJA_NAO_ENCONTRADA'
        if status_loja == 'FALHA_CRIACAO_LOJA':
            log(f"[ERRO MANUAL] Falha ao criar a loja {numero_loja_formatado} ('{nome_loja}') no ComuniQ para o arquivo '{nome_arquivo_original}'.", log_file)
            # A criação pode ter parado no meio de outra página: volta para a lista de Acordos,
            # senão as próximas lojas são procuradas na página errada e todas falham em sequência
            navegar_para_acordos(driver, wait, log_file)
            return 'FALHA_CRIACAO_LOJA'

        input_alvo = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@type='file']")))[2]

        nome_final_alvo = os.path.basename(caminho_arquivo_completo)

        if status_loja == 'CRIADA':
            # Loja acabou de ser criada: não tem anexos, não precisa esperar a lista carregar (economiza 30s)
            nomes_existentes = []
            log(f"[INFO] Loja {numero_loja_formatado} criada agora: sem anexos ainda.", log_file)
        else:
            nomes_existentes = obter_anexos_salvos(driver, wait, log_file)
            log(f"[DEBUG] Anexos na loja {numero_loja_formatado}: {nomes_existentes}", log_file)

        contagem_antes = len(nomes_existentes)
        log(f"[INFO] Contagem antes do upload para '{nome_arquivo_original}': {contagem_antes} arquivos.", log_file)

        if contagem_antes == 0 and status_loja != 'CRIADA':
            log(f"[AVISO] Lista de anexos vazia para loja {numero_loja_formatado} — a checagem de duplicata pode não estar efetiva.", log_file)

        nome_normalizado_alvo = normalizar(nome_final_alvo)
        if any(normalizar(ex) == nome_normalizado_alvo for ex in nomes_existentes):
            log(f"[SKIP] '{nome_final_alvo}' (normalizado: '{nome_normalizado_alvo}') ja existe na loja {numero_loja_formatado}. Pulando upload.", log_file)
            navegar_para_acordos(driver, wait, log_file)
            return 'JA_EXISTE'
        
        input_alvo.send_keys(caminho_arquivo_completo)

        log(f"[OK] Anexado '{nome_arquivo_original}' para a loja {numero_loja_formatado}.", log_file)

        try:
            inputs_atuais = driver.find_elements(By.XPATH, "//input[@type='file']")
            total_inputs = len(inputs_atuais)
            valor_input = inputs_atuais[2].get_attribute('value') or '(vazio)'
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
            cliques_bloqueados_seguidos = 0
        except Exception:
            cliques_bloqueados_seguidos += 1
            log(f"[AVISO] Clique normal no botão Salvar bloqueado ({cliques_bloqueados_seguidos}/{CLIQUES_BLOQUEADOS_PARA_REINICIAR}); tentando clicar via JavaScript.", log_file)
            driver.execute_script("arguments[0].click();", botao_salvar)
        log(f"[DEBUG] Botão Salvar clicado; aguardando retorno à tela de Ações.", log_file)
        verificar_sessao_travada(log_file, f"ao salvar a loja {numero_loja_formatado}")

        try:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Ações']")))
        except TimeoutException:
            log(f"[AVISO] Botão 'Ações' não reapareceu em 15s na loja {numero_loja_formatado}; reabrindo a loja para conferir.", log_file)

        navegar_para_acordos(driver, wait, log_file)
        if abrir_edicao_da_loja(driver, wait, numero_loja_formatado, nome_loja, log_file) == 'OK':
            anexos_salvos = obter_anexos_salvos(driver, wait, log_file)
            contagem_depois = len(anexos_salvos)
            if contagem_depois > contagem_antes:
                navegar_para_acordos(driver, wait, log_file)
                cliques_bloqueados_seguidos = 0
                log(f"[SUCESSO CONFIRMADO] '{nome_final_alvo}' salvo na loja {numero_loja_formatado} (anexos: {contagem_antes} → {contagem_depois}).", log_file, registrar_em_arquivo=False)
                return 'SUCESSO'
            navegar_para_acordos(driver, wait, log_file)
            return 'FALHA_VERIFICACAO'
        else:
            return 'FALHA_VERIFICACAO'

    except SessaoTravada:
        raise
    except Exception as e:
        if isinstance(e, TimeoutException) or _sessao_suspeita():
            log(f"[AVISO] Loja {numero_loja_formatado}: página travada (timeout ao abrir/confirmar a loja ou cliques bloqueados). Reiniciando o navegador...", log_file)
            raise SessaoTravada(f"Página travada ao abrir/confirmar a loja {numero_loja_formatado}.") from e
        log(f"[ERRO GERAL] Falha no upload de '{nome_arquivo_original}' para a loja {numero_loja_formatado}. Erro: {resumir_erro(e)}", log_file)
        driver.get(URL_WIKI)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Toggle Remuneração CS']")))
        navegar_para_acordos(driver, wait, log_file)
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
    # Fecha o navegador atual (se houver) e abre outro já logado na página de Acordos.
    # Tenta até MAX_TENTATIVAS_REINICIO vezes; se todas falharem, devolve (None, None)
    global lojas_desde_ultimo_reinicio, cliques_bloqueados_seguidos

    try:
        driver.quit()
    except Exception:
        pass

    for tentativa in range(1, MAX_TENTATIVAS_REINICIO + 1):
        log(f"[INFO] Abrindo o navegador (tentativa {tentativa}/{MAX_TENTATIVAS_REINICIO})...", log_file)
        novo_driver = inicializar_driver()

        if not novo_driver:
            log("[AVISO] O Chrome não abriu.", log_file)
        else:
            novo_wait = WebDriverWait(novo_driver, 60)
            if not fazer_login(novo_driver, novo_wait, usuario, senha):
                log("[AVISO] O login no ComuniQ falhou.", log_file)
            elif not navegar_para_acordos(novo_driver, novo_wait, log_file):
                log("[AVISO] Não foi possível abrir a página de Acordos.", log_file)
            else:
                lojas_desde_ultimo_reinicio = 0  # navegador novo: zera a contagem do reinício programado
                cliques_bloqueados_seguidos = 0  # navegador novo: sem bloqueios pendentes da sessão antiga
                return novo_driver, novo_wait

            # Abriu, mas não logou ou não navegou: fecha antes de tentar de novo
            try:
                novo_driver.quit()
            except Exception:
                pass

        if tentativa < MAX_TENTATIVAS_REINICIO:
            time.sleep(10)

    log(f"[ERRO] Não foi possível abrir o navegador após {MAX_TENTATIVAS_REINICIO} tentativas.", log_file)
    return None, None

def reinicio_programado(driver, wait, usuario, senha, log_file):
    # Chamado ANTES de cada arquivo: se já passou de LOJAS_POR_REINICIO lojas, reinicia o navegador.
    # Assim o reinício nunca acontece no meio de um arquivo
    if lojas_desde_ultimo_reinicio < LOJAS_POR_REINICIO:
        return driver, wait

    log(f"[INFO] {lojas_desde_ultimo_reinicio} lojas desde o último reinício. Reiniciando o navegador...", log_file, espacamento=True)
    return _reiniciar_sessao(driver, wait, usuario, senha, log_file)

def processar_um_unico_arquivo(driver, wait, nome_arquivo, log_file, lojas_confirmadas=None, diretorio_override=None):
    global lojas_desde_ultimo_reinicio

    if lojas_confirmadas is None:
        lojas_confirmadas = {}

    nome_arquivo_completo = nome_arquivo
    
    nome_arquivo_limpo = nome_arquivo
    nome_base, extensao = os.path.splitext(nome_arquivo)

    sufixo_feito = "_[FEITO]"
    sufixo_ja_existe = "_[JA_EXISTE]"

    if nome_base.endswith(sufixo_feito):
        nome_base_limpo = nome_base[:-len(sufixo_feito)]
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

    # O robô só aceita lojas de até 3 dígitos; números maiores na parte das lojas são ignorados (só avisa no log)
    numeros_ignorados = re.findall(r'\d{4,}', nome_arquivo_limpo.split('_')[0])
    if numeros_ignorados:
        log(f"[AVISO] Número(s) com 4 ou mais dígitos na parte das lojas foram ignorados: {', '.join(numeros_ignorados)} (o robô só aceita lojas de até 3 dígitos).", log_file)

    caminho_arquivo_preparado = renomear_e_preparar_arquivo(nome_arquivo_completo, nome_arquivo_limpo, log_file, diretorio_override)  # [MUDANÇA - Reorganização de pastas]
    if not caminho_arquivo_preparado:
        return 'FALHA_MANUAL', "Erro ao preparar arquivo (verifique padrão da vigência)"

    cidade_loja = extrair_cidade_loja(nome_arquivo_limpo)

    try:
        for numero in numeros_loja:
            numero_formatado = numero.zfill(3)

            nome_loja = montar_nome_loja(numero_formatado, cidade_loja)

            if numero_formatado in lojas_confirmadas:
                resultado_anterior = lojas_confirmadas[numero_formatado]
                if resultado_anterior in ('SUCESSO', 'JA_EXISTE'):
                    log(f"[PULO] Loja {numero_formatado} já processada ('{resultado_anterior}') nesta execução para '{nome_arquivo_completo}'; pulando reenvio.", log_file)
                    resultados_lojas[resultado_anterior].append(numero_formatado)
                    continue
                # Falhou numa tentativa anterior (ex.: antes de reiniciar o navegador): tenta de novo agora
                log(f"[INFO] Loja {numero_formatado} falhou antes ('{resultado_anterior}'); tentando novamente nesta execução.", log_file)

            resultado_upload = upload_com_verificacao_e_retry(driver, wait, numero_formatado, caminho_arquivo_preparado, log_file, nome_arquivo_completo, nome_loja)

            # Converte o resultado do upload no nome da chave usada em resultados_lojas
            if resultado_upload == 'SUCESSO':
                chave_resultado = 'SUCESSO'
            elif resultado_upload == 'JA_EXISTE':
                chave_resultado = 'JA_EXISTE'
            elif resultado_upload == 'FALHA_LOJA_NAO_ENCONTRADA':
                chave_resultado = 'FALHA_MANUAL'
            elif resultado_upload == 'FALHA_CRIACAO_LOJA':
                chave_resultado = 'FALHA_CRIACAO'
            elif resultado_upload == 'UPLOAD_NAO_CONFIRMADO':
                chave_resultado = 'UPLOAD_NAO_CONFIRMADO'
            else:
                chave_resultado = 'FALHA_RECUPERAVEL'

            # Guarda TODO resultado (inclusive falhas): se o navegador cair e o arquivo for retomado,
            # as lojas já feitas são puladas em vez de reprocessadas
            lojas_confirmadas[numero_formatado] = chave_resultado
            resultados_lojas[chave_resultado].append(numero_formatado)

            lojas_desde_ultimo_reinicio += 1  # usado pelo reinício programado (entre um arquivo e outro)
    finally:
        # Apaga o arquivo temporário sempre, mesmo se der erro no meio
        if os.path.exists(caminho_arquivo_preparado):
            os.remove(caminho_arquivo_preparado)

    # Só vai para Sucesso/ se TODAS as lojas deram certo (sucesso ou já existia).
    # Se alguma falhar, o arquivo vai para Reprocessar/. Na próxima execução, as lojas que já receberam
    # o arquivo dão JA_EXISTE (não sobe de novo) e só a que falhou é enviada.
    motivos = []
    if resultados_lojas['FALHA_MANUAL']:
        motivos.append(f"Lojas não encontradas: {', '.join(resultados_lojas['FALHA_MANUAL'])}")
    if resultados_lojas['FALHA_CRIACAO']:
        motivos.append(f"Falha ao criar lojas no ComuniQ: {', '.join(resultados_lojas['FALHA_CRIACAO'])}")
    if resultados_lojas['UPLOAD_NAO_CONFIRMADO']:
        motivos.append(f"Upload não confirmado após 3 tentativas: {', '.join(resultados_lojas['UPLOAD_NAO_CONFIRMADO'])}")
    if resultados_lojas['FALHA_RECUPERAVEL']:
        motivos.append(f"Falha de sistema: {', '.join(resultados_lojas['FALHA_RECUPERAVEL'])}")

    if motivos:
        if resultados_lojas['SUCESSO']:
            motivos.append(f"Já enviado para: {', '.join(resultados_lojas['SUCESSO'])}")
        return 'FALHA', '; '.join(motivos)

    if resultados_lojas['SUCESSO']:
        return 'SUCESSO', "Todas as lojas processadas com sucesso"

    motivo = f"Arquivo já existia em todas as lojas: {', '.join(resultados_lojas['JA_EXISTE'])}"
    return 'JA_EXISTE', motivo

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

def processar_pasta_reprocessar(driver, wait, pasta_reprocessar, pasta_sucesso, log_file, usuario, senha, lojas_confirmadas_por_arquivo, resumo, parar_event=None):  # resumo: dict compartilhado com processar_arquivos para o email
    # Devolve (driver, wait): se o navegador for reiniciado aqui dentro, quem chamou precisa continuar com o novo
    if not os.path.exists(pasta_reprocessar):
        log(f"[INFO] Nada para reprocessar em '{pasta_reprocessar}'.", log_file)
        return driver, wait

    arquivos_reprocessar = listar_arquivos_para_upload(pasta_reprocessar, log_file, resumo['fora_do_padrao'])

    if not arquivos_reprocessar:
        # Pasta vazia, remove
        try:
            os.rmdir(pasta_reprocessar)
            log("[INFO] Pasta Reprocessar/ excluída (estava vazia).", log_file)
        except OSError:
            pass
        return driver, wait

    os.makedirs(pasta_sucesso, exist_ok=True)  # a Sucesso/ do mês anterior pode não existir
    log(f"=== REPROCESSAMENTO: {len(arquivos_reprocessar)} arquivo(s) em '{pasta_reprocessar}' ===", log_file, espacamento=True)

    for nome_arquivo in arquivos_reprocessar:
        if parar_event and parar_event.is_set():
            log("[INTERRUPÇÃO] Processo interrompido.", log_file); break

        caminho_reprocessar = os.path.join(pasta_reprocessar, nome_arquivo)
        resultado, motivo = None, None

        driver, wait = reinicio_programado(driver, wait, usuario, senha, log_file)
        if not driver:  # o arquivo fica onde está e será tentado na próxima execução
            raise Exception(f"Não foi possível reiniciar o navegador após {MAX_TENTATIVAS_REINICIO} tentativas.")

        tentativas_sessao = 0
        while True:
            try:
                resultado, motivo = processar_um_unico_arquivo(
                    driver, wait, nome_arquivo, log_file,
                    lojas_confirmadas_por_arquivo.setdefault(nome_arquivo, {}),
                    diretorio_override=pasta_reprocessar
                )
                break
            except Exception as e:
                if _sessao_morta(e) or isinstance(e, SessaoTravada):
                    tentativas_sessao += 1
                    if tentativas_sessao > MAX_REINICIOS_SESSAO:
                        log(f"[ERRO] Sessão do navegador precisou ser reiniciada {tentativas_sessao} vezes ao reprocessar '{nome_arquivo}'. Desistindo deste arquivo.", log_file)
                        resultado, motivo = 'FALHA', f'Sessão do navegador precisou ser reiniciada {tentativas_sessao} vezes'
                        break
                    if isinstance(e, SessaoTravada):
                        log(f"[AVISO] Página travada ao reprocessar '{nome_arquivo}' (tentativa {tentativas_sessao}/{MAX_REINICIOS_SESSAO}). Reiniciando o navegador...", log_file)
                    else:
                        log(f"[AVISO] Sessão do navegador encerrada ao reprocessar '{nome_arquivo}' (tentativa {tentativas_sessao}/{MAX_REINICIOS_SESSAO}). Reinicializando...", log_file)
                    driver, wait = _reiniciar_sessao(driver, wait, usuario, senha, log_file)
                    if not driver:  # o arquivo fica onde está e será tentado na próxima execução
                        raise Exception(f"Não foi possível reiniciar o navegador após {MAX_TENTATIVAS_REINICIO} tentativas.")
                    continue
                else:
                    log(f"[ERRO FATAL] Erro inesperado ao reprocessar '{nome_arquivo}': {resumir_erro(e)}", log_file, espacamento=True)
                    resultado, motivo = 'FALHA', f'Erro inesperado: {resumir_erro(e)}'
                    break

        if resultado is None:
            resultado, motivo = 'FALHA', 'Erro de sessão/inesperado durante o reprocessamento'

        if resultado == 'SUCESSO':
            mover_e_renomear_arquivo(caminho_reprocessar, pasta_sucesso, "[FEITO]", log_file)
            resumo['sucesso'].append(nome_arquivo)  # conta no email como enviado com sucesso
        elif resultado == 'JA_EXISTE':
            mover_e_renomear_arquivo(caminho_reprocessar, pasta_sucesso, "[JA_EXISTE]", log_file)
            resumo['ja_existe'].append(nome_arquivo)  # conta no email como já existente
        else:
            # Falha: arquivo continua em Reprocessar/ (não move)
            log(f"[AVISO] Arquivo '{nome_arquivo}' falhou ao reprocessar, mantido na pasta Reprocessar/.", log_file)
            resumo['falhas'].append({'arquivo': nome_arquivo, 'motivo': motivo})  # lista no email com o motivo da falha

        # Dados do arquivo para o relatório simples: resultado final e o que aconteceu em cada loja
        resumo['arquivos'].append({
            'arquivo': nome_arquivo,
            'resultado': resultado,
            'motivo': motivo,
            'lojas': lojas_confirmadas_por_arquivo.get(nome_arquivo, {}),
        })

    # Se Reprocessar/ ficou vazia, exclui
    try:
        if os.path.exists(pasta_reprocessar) and not os.listdir(pasta_reprocessar):
            os.rmdir(pasta_reprocessar)
            log("[INFO] Pasta Reprocessar/ excluída (todos os arquivos processados com sucesso).", log_file, espacamento=True)
    except OSError:
        pass

    return driver, wait

def processar_arquivos(diretorio_origem, pasta_temp, log_file_path, usuario, senha, parar_event=None, automatico=False):
    # automatico=True (--auto-start / agendador): ignora diretorio_origem e usa a pasta do mês atual,
    # verificando antes a Reprocessar/ do mês anterior
    global cliques_bloqueados_seguidos
    arquivar_log_anterior(log_file_path)  # guarda o log da execução anterior na pasta Logs (os 30 últimos)
    lojas_criadas_na_execucao.clear()      # começa a execução sem nenhuma loja criada
    cliques_bloqueados_seguidos = 0        # começa a execução sem bloqueios pendentes de uma sessão anterior
    data_inicio = datetime.now()  # data/hora de início, usada no assunto/corpo do email
    # Acumula os resultados da execução para o relatório simples e o email
    resumo = {
        'sucesso': [], 'ja_existe': [], 'falhas': [], 'erro_fatal': None,
        'pasta': None,         # pasta processada (aparece no cabeçalho do relatório)
        'pasta_mes_anterior': None,    # Reprocessar/ do mês anterior, quando havia pendências lá
        'pasta_nao_encontrada': None,  # motivo, quando a pasta do mês atual não foi encontrada (modo automático)
        'arquivos': [],        # um item por arquivo, com o resultado de cada loja
        'fora_do_padrao': [],  # arquivos ignorados por estarem com o nome fora do padrão
    }

    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log(f"=== INÍCIO DA AUTOMAÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===", log_file, espacamento=True)

        pasta_mes_anterior = None
        if automatico:
            hoje = datetime.now()
            if hoje.month == 1:  # em janeiro, o mês anterior é dezembro do ano anterior
                mes_anterior = 12
                ano_anterior = hoje.year - 1
            else:
                mes_anterior = hoje.month - 1
                ano_anterior = hoje.year
            diretorio_origem, erro_mes_atual = encontrar_pasta_mes(hoje.year, hoje.month)
            pasta_mes_anterior, erro_mes_anterior = encontrar_pasta_mes(ano_anterior, mes_anterior)
            if erro_mes_atual:
                log(f"[ERRO] {erro_mes_atual}", log_file)
                resumo['pasta_nao_encontrada'] = erro_mes_atual  # o email é enviado avisando que nada foi processado
            if erro_mes_anterior:
                log(f"[INFO] Reprocessar/ do mês anterior não verificada: {erro_mes_anterior}", log_file)

        if diretorio_origem:
            resumo['pasta'] = os.path.normpath(diretorio_origem)  # normpath: mostra o caminho com '\' no email

        global DIRETORIO_ORIGEM, PASTA_TEMP, LOG_FILE_PATH
        DIRETORIO_ORIGEM = diretorio_origem or ""; PASTA_TEMP = pasta_temp; LOG_FILE_PATH = log_file_path

        nome_pasta_sucesso = "Sucesso"
        nome_pasta_reprocessar = "Reprocessar"

        base_normalizado = pasta_sucesso = pasta_reprocessar = None
        if diretorio_origem:  # só é None no modo automático quando a pasta do mês atual não existe
            base_normalizado = os.path.basename(diretorio_origem.rstrip(os.sep))
            if base_normalizado in [nome_pasta_sucesso, nome_pasta_reprocessar]:
                diretorio_raiz = os.path.dirname(diretorio_origem.rstrip(os.sep))
            else:
                diretorio_raiz = diretorio_origem

            pasta_sucesso = os.path.join(diretorio_raiz, nome_pasta_sucesso)
            pasta_reprocessar = os.path.join(diretorio_raiz, nome_pasta_reprocessar)

            os.makedirs(pasta_sucesso, exist_ok=True)

        # Pendências do mês anterior (aparece no email como pasta acessada)
        if pasta_mes_anterior and os.path.exists(os.path.join(pasta_mes_anterior, nome_pasta_reprocessar)):
            resumo['pasta_mes_anterior'] = os.path.normpath(os.path.join(pasta_mes_anterior, nome_pasta_reprocessar))

        resultados_finais_falhas = resumo['falhas']  # mesma lista do resumo: falhas da pasta de origem e da Reprocessar/ vão juntas para o email
        driver = None
        wait = None
        try:
            if not diretorio_origem and not resumo['pasta_mes_anterior']:
                # Sem pasta do mês atual e sem pendências do mês anterior: não há o que fazer no ComuniQ
                log("[INFO] Nada para processar: sem pasta do mês atual e sem pendências do mês anterior. O navegador não será aberto.", log_file)
            else:
                # Abre o navegador já logado na página de Acordos (até MAX_TENTATIVAS_REINICIO tentativas)
                driver, wait = _reiniciar_sessao(None, None, usuario, senha, log_file)
                if not driver:
                    raise Exception(f"Não foi possível abrir o navegador e entrar no ComuniQ após {MAX_TENTATIVAS_REINICIO} tentativas.")

            lojas_confirmadas_por_arquivo = {}
            if pasta_mes_anterior:  # modo automático: antes do mês atual, tenta de novo o que ficou pendente no mês anterior
                log(f"[INFO] Verificando Reprocessar/ do mês anterior: '{pasta_mes_anterior}'", log_file, espacamento=True)
                driver, wait = processar_pasta_reprocessar(
                    driver, wait,
                    os.path.join(pasta_mes_anterior, nome_pasta_reprocessar), os.path.join(pasta_mes_anterior, nome_pasta_sucesso),
                    log_file, usuario, senha, lojas_confirmadas_por_arquivo, resumo, parar_event
                )
            if pasta_reprocessar:
                driver, wait = processar_pasta_reprocessar(driver, wait, pasta_reprocessar, pasta_sucesso, log_file, usuario, senha, lojas_confirmadas_por_arquivo, resumo, parar_event)  # passa o resumo para registrar os resultados da Reprocessar/

            veio_da_reprocessar = base_normalizado == nome_pasta_reprocessar

            if not diretorio_origem:
                log("[INFO] Sem pasta do mês atual; nada mais a processar.", log_file)
            elif veio_da_reprocessar:
                log("[INFO] Execução diretamente sobre Reprocessar/; processamento já concluído pelo reprocessador.", log_file, espacamento=True)
            elif not os.path.exists(diretorio_origem):
                log("[INFO] Diretório de origem não existe mais (nada a processar).", log_file)
            else:
                arquivos_para_processar = listar_arquivos_para_upload(diretorio_origem, log_file, resumo['fora_do_padrao'])

                log(f"Encontrados {len(arquivos_para_processar)} arquivos para processar em '{diretorio_origem}'.", log_file)

                for nome_arquivo in arquivos_para_processar:
                    if parar_event and parar_event.is_set():
                        log("[INTERRUPÇÃO] Processo interrompido.", log_file); break
                    
                    caminho_original = os.path.join(diretorio_origem, nome_arquivo)
                    resultado, motivo = None, None

                    driver, wait = reinicio_programado(driver, wait, usuario, senha, log_file)
                    if not driver:  # o arquivo fica onde está e será tentado na próxima execução
                        raise Exception(f"Não foi possível reiniciar o navegador após {MAX_TENTATIVAS_REINICIO} tentativas.")

                    tentativas_sessao = 0
                    while True:
                        try:
                            resultado, motivo = processar_um_unico_arquivo(driver, wait, nome_arquivo, log_file, lojas_confirmadas_por_arquivo.setdefault(nome_arquivo, {}))
                            break
                        except Exception as e:
                            if _sessao_morta(e) or isinstance(e, SessaoTravada):
                                tentativas_sessao += 1
                                if tentativas_sessao > MAX_REINICIOS_SESSAO:
                                    log(f"[ERRO] Sessão do navegador precisou ser reiniciada {tentativas_sessao} vezes ao processar '{nome_arquivo}'. Desistindo deste arquivo.", log_file)
                                    resultado, motivo = 'FALHA', f'Sessão do navegador precisou ser reiniciada {tentativas_sessao} vezes'
                                    break
                                if isinstance(e, SessaoTravada):
                                    log(f"[AVISO] Página travada ao processar '{nome_arquivo}' (tentativa {tentativas_sessao}/{MAX_REINICIOS_SESSAO}). Reiniciando o navegador...", log_file)
                                else:
                                    log(f"[AVISO] Sessão do navegador encerrada ao processar '{nome_arquivo}' (tentativa {tentativas_sessao}/{MAX_REINICIOS_SESSAO}). Reinicializando...", log_file)
                                driver, wait = _reiniciar_sessao(driver, wait, usuario, senha, log_file)
                                if not driver:  # o arquivo fica onde está e será tentado na próxima execução
                                    raise Exception(f"Não foi possível reiniciar o navegador após {MAX_TENTATIVAS_REINICIO} tentativas.")
                                continue
                            else:
                                log(f"[ERRO FATAL] Erro inesperado ao processar '{nome_arquivo}': {resumir_erro(e)}", log_file, espacamento=True)
                                resultado, motivo = 'FALHA', f'Erro inesperado: {resumir_erro(e)}'
                                break 
                    if resultado is None:
                        resultado, motivo = 'FALHA', 'Erro de sessão/inesperado durante o processamento'

                    if resultado == 'SUCESSO':
                        mover_e_renomear_arquivo(caminho_original, pasta_sucesso, "[FEITO]", log_file)
                        resumo['sucesso'].append(nome_arquivo)  # conta no email como enviado com sucesso
                    elif resultado == 'JA_EXISTE':
                        mover_e_renomear_arquivo(caminho_original, pasta_sucesso, "[JA_EXISTE]", log_file)
                        resumo['ja_existe'].append(nome_arquivo)  # conta no email como já existente
                    else:
                        os.makedirs(pasta_reprocessar, exist_ok=True)
                        shutil.move(caminho_original, os.path.join(pasta_reprocessar, nome_arquivo))
                        log(f"[INFO] Arquivo '{nome_arquivo}' movido para Reprocessar/.", log_file)
                        resultados_finais_falhas.append({'arquivo': nome_arquivo, 'motivo': motivo})

                    # Dados do arquivo para o relatório simples: resultado final e o que aconteceu em cada loja
                    resumo['arquivos'].append({
                        'arquivo': nome_arquivo,
                        'resultado': resultado,
                        'motivo': motivo,
                        'lojas': lojas_confirmadas_por_arquivo.get(nome_arquivo, {}),
                    })
        
        except Exception as e:
            log(f"[ERRO FATAL] Ocorreu um erro geral no processo: {resumir_erro(e)}", log_file, espacamento=True)
            resumo['erro_fatal'] = str(e)  # destaca no email que a execução foi interrompida por erro

        finally:
            try:
                if pasta_reprocessar and os.path.exists(pasta_reprocessar) and not os.listdir(pasta_reprocessar):
                    os.rmdir(pasta_reprocessar)
                    log("[INFO] Pasta Reprocessar/ excluída (todos os arquivos processados com sucesso).", log_file, espacamento=True)
            except OSError:
                pass

            log(f"=== FIM DA AUTOMAÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===", log_file, espacamento=True)
            if driver: driver.quit()

    resumo['lojas_criadas'] = list(lojas_criadas_na_execucao)  # aparece no relatório e no corpo do email

    # Relatório simples (para quem não é da área de TI), gravado ao lado do log técnico.
    # É ele que vai no email; o log_historico.txt fica só no computador
    relatorio_path = os.path.join(os.path.dirname(log_file_path), "relatorio.txt")
    with open(relatorio_path, 'w', encoding='utf-8') as arquivo_relatorio:
        arquivo_relatorio.write(montar_relatorio(resumo, data_inicio))

    if parar_event and parar_event.is_set():
        print("[EMAIL] Não enviado: o processo foi interrompido pelo usuário.")
    elif resumo['erro_fatal'] and not resumo['arquivos']:
        # O robô não conseguiu rodar (Chrome, login, pasta do mês...) e nenhum arquivo foi processado
        print("[EMAIL] Não enviado: o robô não conseguiu rodar. Veja o log_historico.txt.")
    else:
        print("[EMAIL] Enviando relatório da execução...")
        enviar_relatorio(resumo, data_inicio, log_file_path, relatorio_path)  # erros do envio vão para o log técnico
    return resultados_finais_falhas  # retorno mantido (lista de falhas) para não mudar o contrato com a interface