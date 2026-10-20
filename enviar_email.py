import smtplib
import os
import html  # html.escape: evita que caracteres como & e < no nome da pasta quebrem o email
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".RoboComuniq", "config.ini")
TIMEOUT_SMTP = 180

################################################################

def carregar_config_email():
    config = configparser.ConfigParser(interpolation=None)  
    if not os.path.exists(CONFIG_PATH):
        raise FileNotFoundError(f"Arquivo de configuração não encontrado: {CONFIG_PATH}")

    config.read(CONFIG_PATH, encoding="utf-8")

    if "Email" not in config:
        raise KeyError("Seção [Email] não encontrada no config.ini. Adicione EmailRemetente, SenhaApp e EmailDestinatario.")

    email_cfg = config["Email"]
    remetente = email_cfg.get("EmailRemetente", "")
    senha = email_cfg.get("SenhaApp", "")  # senha de app do Gmail, não a senha normal da conta
    destinatarios_raw = email_cfg.get("EmailDestinatario", "")

    if not remetente or not senha or not destinatarios_raw:
        raise ValueError("Configuração de email incompleta. Verifique EmailRemetente, SenhaApp e EmailDestinatario no config.ini")

    destinatarios = [
        email.strip()
        for email in destinatarios_raw.replace(",", ";").split(";")
        if email.strip()
    ]

    if not destinatarios:
        raise ValueError("EmailDestinatario vazio ou inválido no config.ini. Separe múltiplos emails por ponto e vírgula (;) ou vírgula (,).")

    return remetente, senha, destinatarios

################################################################

def _log_erro_email(arquivo_log, mensagem):
    # anota erros no log da execução
    data_log = datetime.now().strftime("%d-%m-%Y_%H-%M")
    layout_mensagem = (
        f"Erro na data: {data_log}\n"
        f"ENVIO DE EMAIL\n\n"
        f"{mensagem}\n"
        f"\n-------------------------------------------------\n\n"
    )
    print(layout_mensagem)
    if arquivo_log:
        with open(arquivo_log, "a", encoding="utf-8") as log:
            log.write(layout_mensagem)

################################################################

def mensagem_da_loja(resultado):
    # Traduz o resultado técnico de cada loja para um texto simples
    if resultado == 'SUCESSO':
        return "enviado com sucesso"
    elif resultado == 'JA_EXISTE':
        return "o arquivo já estava no ComuniQ (nada foi enviado)"
    elif resultado == 'FALHA_MANUAL':
        return "NÃO ENVIADO - a loja não foi encontrada no ComuniQ"
    elif resultado == 'FALHA_CRIACAO':
        return "NÃO ENVIADO - não foi possível criar a loja no ComuniQ"
    elif resultado == 'UPLOAD_NAO_CONFIRMADO':
        return "NÃO ENVIADO - o envio não foi confirmado após 3 tentativas"
    else:
        return "NÃO ENVIADO - ocorreu um erro no site durante o envio"

################################################################

SEPARADOR = "=" * 60  # linha que separa uma seção da outra no relatorio.txt

def linha_com_pontos(texto, numero):
    # Linha do resumo, ex.: "   Arquivos enviados ...................... 2"
    pontos = "." * (40 - len(texto))
    return f"   {texto} {pontos} {numero}"

def linhas_do_arquivo(item, numeros_criados):
    # Nome do arquivo e, logo abaixo, o que aconteceu em cada loja.
    # numeros_criados: lojas criadas pelo robô nesta execução (ganham uma marcação ao lado)
    linhas = [item['arquivo']]
    for loja in item['lojas']:
        texto = f"   Loja {loja}: {mensagem_da_loja(item['lojas'][loja])}"
        if loja in numeros_criados:
            texto += " (loja criada nesta execução)"
        linhas.append(texto)
    if not item['lojas']:  # falhou antes de chegar nas lojas
        motivo = str(item['motivo']).split("\n")[0]  # só a 1ª linha: o resto é detalhe técnico
        linhas.append(f"   NÃO ENVIADO - {motivo}")
    linhas.append("")
    return linhas

def linhas_da_secao(titulo, explicacao, itens, numeros_criados):
    # Uma seção: título com a quantidade, a explicação e os arquivos logo abaixo
    linhas = [SEPARADOR, f"{titulo} ({len(itens)})"]
    linhas.extend(explicacao)
    linhas.append("")
    if not itens:
        linhas.append("Nenhum arquivo.")
        linhas.append("")
    for item in itens:
        linhas.extend(linhas_do_arquivo(item, numeros_criados))
    return linhas

def montar_relatorio(resumo, data_ref):
    # Relatório para quem não é da área de TI, enviado em anexo (relatorio.txt), separado por seções.
    # Os detalhes técnicos ficam só no log_historico.txt

    # Separa os arquivos processados pelo resultado final
    enviados = []
    ja_existiam = []
    pendentes = []
    for item in resumo.get('arquivos', []):
        if item['resultado'] == 'SUCESSO':
            enviados.append(item)
        elif item['resultado'] == 'JA_EXISTE':
            ja_existiam.append(item)
        else:
            pendentes.append(item)
    fora_do_padrao = resumo.get('fora_do_padrao', [])

    # Lojas criadas pelo robô nesta execução (cada uma: {'numero': '999', 'nome': '999 - TESTE'})
    lojas_criadas = resumo.get('lojas_criadas', [])
    numeros_criados = []
    for loja in lojas_criadas:
        numeros_criados.append(loja['numero'])

    # Cabeçalho
    linhas = []
    linhas.append("RELATÓRIO DO ROBÔ CCTs COMUNIQ")
    linhas.append("Atividade: Envio das Convenções Coletivas de Trabalho para o ComuniQ")
    linhas.append(f"Data de execução: {data_ref.strftime('%d/%m/%Y às %H:%M')}")
    if resumo.get('pasta'):
        linhas.append(f"Pasta acessada: {resumo['pasta']}")
    if resumo.get('pasta_mes_anterior'):
        linhas.append(f"Pendências do mês anterior: {resumo['pasta_mes_anterior']}")
    linhas.append("")

    # Avisos, logo abaixo do cabeçalho
    if resumo.get('pasta_nao_encontrada'):
        linhas.append(f"AVISO: {resumo['pasta_nao_encontrada']}")
        linhas.append("")
    if not resumo.get('arquivos'):
        linhas.append("Nenhum arquivo foi processado nesta execução.")
        linhas.append("")
    if resumo.get('erro_fatal'):
        motivo_erro = str(resumo['erro_fatal']).split("\n")[0]  # só a 1ª linha: o resto é detalhe técnico
        linhas.append("ATENÇÃO: o robô não conseguiu concluir o trabalho.")
        linhas.append(f"Motivo: {motivo_erro}")
        linhas.append("")

    # Resumo
    linhas.append(SEPARADOR)
    linhas.append("RESUMO")
    linhas.append(linha_com_pontos("Arquivos enviados", len(enviados)))
    linhas.append(linha_com_pontos("Já estavam no ComuniQ", len(ja_existiam)))
    linhas.append(linha_com_pontos("Pendentes", len(pendentes)))
    linhas.append(linha_com_pontos("Com o nome fora do padrão", len(fora_do_padrao)))
    linhas.append(linha_com_pontos("Lojas criadas no ComuniQ", len(lojas_criadas)))
    linhas.append("")

    # Seções
    linhas.extend(linhas_da_secao(
        "ARQUIVOS ENVIADOS",
        ["Subiram para todas as lojas e foram movidos para a pasta Sucesso."],
        enviados, numeros_criados))
    linhas.extend(linhas_da_secao(
        "JÁ ESTAVAM NO COMUNIQ",
        ["O arquivo já existia em todas as lojas; nada foi enviado."],
        ja_existiam, numeros_criados))
    linhas.extend(linhas_da_secao(
        "PENDENTES",
        ["Alguma loja falhou. O arquivo está na pasta Reprocessar e será",
         "tentado de novo na próxima execução."],
        pendentes, numeros_criados))

    # Lojas criadas no ComuniQ pelo robô
    linhas.append(SEPARADOR)
    linhas.append(f"LOJAS CRIADAS NO COMUNIQ ({len(lojas_criadas)})")
    linhas.append("Não existiam no ComuniQ e foram criadas automaticamente pelo robô.")
    linhas.append("Confira se o nome da página ficou correto.")
    linhas.append("")
    if not lojas_criadas:
        linhas.append("Nenhuma loja criada.")
    for loja in lojas_criadas:
        linhas.append(f"   - {loja['nome']}")
    linhas.append("")

    # Nome fora do padrão (aqui são só nomes de arquivo, sem lojas)
    linhas.append(SEPARADOR)
    linhas.append(f"NOME FORA DO PADRÃO ({len(fora_do_padrao)})")
    linhas.append("Não foram enviados. Renomeie no padrão lojas_cidade_sindicato_vigência")
    linhas.append("(exemplo: 160,310_LAJEADO_SIND0088_2026) para que sejam enviados.")
    linhas.append("")
    if not fora_do_padrao:
        linhas.append("Nenhum arquivo.")
    for nome in fora_do_padrao:
        linhas.append(f"   - {nome}")

    return "\n".join(linhas)

################################################################

# O HTML do email é montado com TABELAS (e não <div>) porque o Outlook para Windows usa o motor do Word,
# que ignora largura, fundo e espaçamento em <div>. Tabelas funcionam no Outlook, no Gmail e no celular.
# A fonte é repetida em cada célula porque o Outlook não herda a fonte da tabela de fora.
FONTE = "font-family: Arial, sans-serif;"

def caixa_aviso(texto, cor_borda, cor_fundo):
    # Caixa colorida usada para os avisos no corpo do email (tabela de uma célula)
    return f"""
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td bgcolor="{cor_fundo}" style="background-color: {cor_fundo}; border-left: 5px solid {cor_borda}; padding: 10px 14px; {FONTE} font-size: 14px; color: #333333;">{texto}</td>
        </tr>
      </table>"""

def linha_resumo(titulo, explicacao, numero, cor):
    # Uma linha da tabela de resumo: título e explicação à esquerda, número grande à direita
    return f"""
        <tr>
          <td style="padding: 10px 0; border-bottom: 1px solid #eeeeee; {FONTE}">
            <div style="font-size: 15px; font-weight: bold; color: #333333;">{titulo}</div>
            <div style="font-size: 13px; color: #666666;">{explicacao}</div>
          </td>
          <td width="60" align="right" style="padding: 10px 0; border-bottom: 1px solid #eeeeee; {FONTE} font-size: 22px; font-weight: bold; color: {cor};">{numero}</td>
        </tr>"""

def montar_corpo_html(resumo, data_ref):
    # Corpo do email em HTML: cabeçalho, pasta acessada e o resumo (os números) com uma explicação em cada um.
    # Os detalhes de cada arquivo ficam no relatorio.txt em anexo.
    # Os estilos vão dentro de cada tag (style="...") porque os programas de email ignoram CSS separado
    enviados = len(resumo.get('sucesso', []))
    ja_existiam = len(resumo.get('ja_existe', []))
    pendentes = len(resumo.get('falhas', []))
    fora_do_padrao = len(resumo.get('fora_do_padrao', []))
    lojas_criadas = len(resumo.get('lojas_criadas', []))

    verde = "#397F5A"  # mesmo verde da interface do robô
    cinza = "#666666"
    laranja = "#D9730D"
    vermelho = "#C0392B"
    azul = "#2E6DA4"

    # Lojas criadas ficam em azul quando houver alguma: não é erro, mas vale conferir o nome no anexo
    cor_lojas_criadas = cinza
    if lojas_criadas > 0:
        cor_lojas_criadas = azul

    # Pendentes e fora do padrão ficam em laranja só quando há algo para a pessoa resolver
    cor_pendentes = cinza
    if pendentes > 0:
        cor_pendentes = laranja

    cor_fora_do_padrao = cinza
    if fora_do_padrao > 0:
        cor_fora_do_padrao = laranja

    # Pasta(s) que o robô acessou
    if resumo.get('pasta'):
        pastas_html = f"<div>{html.escape(resumo['pasta'])}</div>"
    else:
        pastas_html = "<div>Pasta do mês atual não encontrada</div>"
    if resumo.get('pasta_mes_anterior'):
        pastas_html += f"<div>{html.escape(resumo['pasta_mes_anterior'])} (pendências do mês anterior)</div>"

    # Avisos que só aparecem quando necessário
    avisos_html = ""
    if resumo.get('erro_fatal'):
        motivo = str(resumo['erro_fatal']).split("\n")[0]  # só a 1ª linha: o resto é detalhe técnico
        avisos_html += caixa_aviso(f"<b>ATENÇÃO:</b> a execução foi interrompida antes da conclusão. Motivo: {html.escape(motivo)}", vermelho, "#FDECEA")
    if not resumo.get('arquivos'):
        if resumo.get('pasta_nao_encontrada'):
            avisos_html += caixa_aviso(f"<b>AVISO:</b> {html.escape(resumo['pasta_nao_encontrada'])}. Nenhum arquivo foi processado.", laranja, "#FFF4E5")
        else:
            avisos_html += caixa_aviso("Não foram localizados arquivos para processamento nesta execução.", cinza, "#F2F2F2")

    # Tabela de resumo
    tabela = ""
    tabela += linha_resumo("Arquivos enviados", "Documentos enviados com sucesso para todas as lojas correspondentes.", enviados, verde)
    tabela += linha_resumo("Arquivos já existentes no ComuniQ", "Documentos que já constavam em todas as lojas; nenhum envio foi necessário.", ja_existiam, cinza)
    tabela += linha_resumo("Arquivos pendentes", "Houve falha no envio para ao menos uma loja; o arquivo será reprocessado na próxima execução.", pendentes, cor_pendentes)
    tabela += linha_resumo("Lojas criadas no ComuniQ", "Lojas inexistentes no ComuniQ, criadas automaticamente. A relação consta no relatório em anexo.", lojas_criadas, cor_lojas_criadas)
    tabela += linha_resumo("Arquivos fora do padrão de nomenclatura", "Documentos não enviados em razão do nome do arquivo.", fora_do_padrao, cor_fora_do_padrao)
    if fora_do_padrao > 0:
        aviso = (
            "<b>ATENÇÃO:</b> foram identificados arquivos fora do padrão de nomenclatura. Revise a pasta para a renomeação.<br>"
            "Padrão: <b>lojas_cidade_sindicato_vigência</b>"
        )
        tabela += f'<tr><td colspan="2" style="padding: 8px 0;">{caixa_aviso(aviso, laranja, "#FFF4E5")}</td></tr>'

    # width="600" é a largura que o Outlook respeita (não estica até a janela);
    # max-width: 100% deixa os outros programas e o celular encolherem quando a tela for menor
    return f"""
<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width: 600px; max-width: 100%; border: 1px solid #dddddd;">
  <tr>
    <td bgcolor="{verde}" style="background-color: {verde}; padding: 18px 24px; {FONTE}">
      <div style="font-size: 20px; font-weight: bold; color: #ffffff;">Relatório do Robô CCTs ComuniQ</div>
      <div style="font-size: 14px; color: #E3F1E8; padding-top: 6px;">Envio das Convenções Coletivas de Trabalho para o ComuniQ</div>
      <div style="font-size: 14px; color: #E3F1E8;">{data_ref.strftime('%d/%m/%Y às %H:%M')}</div>
    </td>
  </tr>
  <tr>
    <td style="padding: 20px 24px; {FONTE} color: #333333;">
      <div style="font-size: 13px; color: #666666;">Pasta acessada:</div>
      <div style="font-size: 13px; color: #333333;">{pastas_html}</div>
      <p style="margin: 16px 0 8px 0; font-size: 15px; color: #333333;">Olá,</p>
      <p style="margin: 0 0 12px 0; font-size: 15px; color: #333333;">Segue o resumo da execução. Os detalhes de cada arquivo estão no relatório em anexo.</p>
      {avisos_html}
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">{tabela}
      </table>
      <p style="margin: 24px 0 0 0; font-size: 15px; color: #333333;">Att,<br>Robô CCTs ComuniQ</p>
    </td>
  </tr>
</table>
"""

################################################################

def enviar_relatorio(resumo, data_execucao=None, arquivo_log=None, arquivo_relatorio=None):
    # arquivo_log: log técnico, onde os erros do envio são anotados (NÃO vai no email)
    # arquivo_relatorio: relatório simples, que vai em anexo
    # Trata os próprios erros: registra no log e devolve False em vez de interromper o robô
    try:
        remetente, senha, destinatarios = carregar_config_email()
    except (FileNotFoundError, KeyError, ValueError) as e:
        _log_erro_email(arquivo_log, f"[EMAIL] Erro ao carregar configuração: {e}")
        return False

    data_ref = data_execucao if data_execucao else datetime.now()
    assunto = f"Relatório Robô CCTs ComuniQ - {data_ref.strftime('%d/%m/%Y')}"

    corpo = montar_corpo_html(resumo, data_ref)

    msg = MIMEMultipart()  # mensagem com corpo de texto + anexo. sem multipart teria só um dos dois
    msg["From"] = remetente
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "html", "utf-8"))

    if arquivo_relatorio and os.path.exists(arquivo_relatorio):
        try:
            with open(arquivo_relatorio, "rb") as f:
                parte = MIMEBase("text", "plain")
                parte.set_payload(f.read())
                encoders.encode_base64(parte)
                parte.add_header(
                    "Content-Disposition",
                    f"attachment; filename=\"{os.path.basename(arquivo_relatorio)}\""
                )
                msg.attach(parte)
        except Exception as e:
            _log_erro_email(arquivo_log, f"[EMAIL] Erro ao anexar o relatório (email segue sem anexo): {e}")

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=TIMEOUT_SMTP) as servidor:
            servidor.starttls()
            servidor.login(remetente, senha)
            servidor.sendmail(remetente, destinatarios, msg.as_string())
        print(f"[EMAIL] Email enviado com sucesso para {', '.join(destinatarios)}")
        return True
    except smtplib.SMTPAuthenticationError:
        _log_erro_email(arquivo_log, "[EMAIL] Erro de autenticação. Verifique o email e a senha de app no config.ini")
        return False
    except Exception as e:
        _log_erro_email(arquivo_log, f"[EMAIL] Erro ao enviar email: {e}")
        return False
