import smtplib
import os
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

def montar_relatorio(resumo, data_ref):
    # Relatório para quem não é da área de TI: vai no corpo do email e no anexo relatorio.txt.
    # Os detalhes técnicos ficam só no log_historico.txt
    separador = "-" * 60
    linhas = []
    linhas.append("RELATÓRIO DO ROBÔ CCTs COMUNIQ")
    linhas.append(f"Data: {data_ref.strftime('%d/%m/%Y às %H:%M')}")
    if resumo.get('pasta'):
        linhas.append(f"Pasta: {resumo['pasta']}")

    if resumo.get('erro_fatal'):
        motivo_erro = str(resumo['erro_fatal']).split("\n")[0]  # só a 1ª linha: o resto é detalhe técnico
        linhas.append("")
        linhas.append("ATENÇÃO: o robô não conseguiu concluir o trabalho.")
        linhas.append(f"Motivo: {motivo_erro}")

    for item in resumo.get('arquivos', []):
        linhas.append("")
        linhas.append(separador)
        linhas.append(f"Arquivo: {item['arquivo']}")

        for loja in item['lojas']:
            linhas.append(f"   Loja {loja}: {mensagem_da_loja(item['lojas'][loja])}")

        if item['resultado'] == 'SUCESSO':
            linhas.append("   Situação: CONCLUÍDO. O arquivo foi movido para a pasta Sucesso.")
        elif item['resultado'] == 'JA_EXISTE':
            linhas.append("   Situação: JÁ ESTAVA NO COMUNIQ. O arquivo foi movido para a pasta Sucesso.")
        else:
            if not item['lojas']:  # falhou antes de chegar nas lojas
                motivo_arquivo = str(item['motivo']).split("\n")[0]
                linhas.append(f"   NÃO ENVIADO - {motivo_arquivo}")
            linhas.append("   Situação: PENDENTE. O arquivo está na pasta Reprocessar e será")
            linhas.append("   tentado de novo na próxima execução.")

    fora_do_padrao = resumo.get('fora_do_padrao', [])
    if fora_do_padrao:
        linhas.append("")
        linhas.append(separador)
        linhas.append("ARQUIVOS NÃO ENVIADOS POR ESTAREM COM O NOME FORA DO PADRÃO")
        linhas.append("Para serem enviados, precisam ser renomeados no padrão:")
        linhas.append("lojas_cidade_sindicato_vigência (exemplo: 160,310_LAJEADO_SIND0088_2026)")
        for nome in fora_do_padrao:
            linhas.append(f"   - {nome}")

    linhas.append("")
    linhas.append(separador)
    linhas.append("RESUMO")
    linhas.append(f"   Arquivos concluídos: {len(resumo.get('sucesso', []))}")
    linhas.append(f"   Arquivos que já estavam no ComuniQ: {len(resumo.get('ja_existe', []))}")
    linhas.append(f"   Arquivos pendentes: {len(resumo.get('falhas', []))}")
    linhas.append(f"   Arquivos com o nome fora do padrão: {len(fora_do_padrao)}")
    return "\n".join(linhas)

################################################################

def montar_corpo_html(resumo, data_ref):
    # Corpo do email em HTML: só o resumo (os números). Os detalhes ficam no relatorio.txt em anexo.
    # Os estilos vão dentro de cada tag (style="...") porque os programas de email ignoram CSS separado
    enviados = len(resumo.get('sucesso', []))
    ja_existiam = len(resumo.get('ja_existe', []))
    pendentes = len(resumo.get('falhas', []))
    fora_do_padrao = len(resumo.get('fora_do_padrao', []))

    verde = "#397F5A"  # mesmo verde da interface do robô
    cinza = "#666666"
    laranja = "#D9730D"

    # Pendentes e fora do padrão ficam em laranja só quando há algo para a pessoa resolver
    cor_pendentes = cinza
    if pendentes > 0:
        cor_pendentes = laranja

    cor_fora_do_padrao = cinza
    if fora_do_padrao > 0:
        cor_fora_do_padrao = laranja

    estilo_texto = "padding: 10px 0; border-bottom: 1px solid #eeeeee; font-size: 15px;"
    estilo_numero = "padding: 10px 0; border-bottom: 1px solid #eeeeee; font-size: 20px; font-weight: bold; text-align: right;"

    return f"""
<div style="font-family: Arial, sans-serif; max-width: 560px; border: 1px solid #dddddd;">
  <div style="background-color: {verde}; color: #ffffff; padding: 18px 24px;">
    <div style="font-size: 20px; font-weight: bold;">Relatório do Robô CCTs ComuniQ</div>
    <div style="font-size: 14px;">Execução de {data_ref.strftime('%d/%m/%Y às %H:%M')}</div>
  </div>
  <div style="padding: 20px 24px; color: #333333;">
    <p style="font-size: 15px;">Olá,</p>
    <p style="font-size: 15px;">Segue o resumo da execução. Os detalhes de cada arquivo estão no relatório em anexo.</p>
    <table style="width: 100%; border-collapse: collapse;">
      <tr>
        <td style="{estilo_texto}">Arquivos enviados</td>
        <td style="{estilo_numero} color: {verde};">{enviados}</td>
      </tr>
      <tr>
        <td style="{estilo_texto}">Já estavam no ComuniQ</td>
        <td style="{estilo_numero} color: {cinza};">{ja_existiam}</td>
      </tr>
      <tr>
        <td style="{estilo_texto}">Pendentes (serão tentados de novo)</td>
        <td style="{estilo_numero} color: {cor_pendentes};">{pendentes}</td>
      </tr>
      <tr>
        <td style="{estilo_texto}">Com o nome fora do padrão</td>
        <td style="{estilo_numero} color: {cor_fora_do_padrao};">{fora_do_padrao}</td>
      </tr>
    </table>
    <p style="font-size: 15px; margin-top: 24px;">Att,<br>Robô CCTs ComuniQ</p>
  </div>
</div>
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
