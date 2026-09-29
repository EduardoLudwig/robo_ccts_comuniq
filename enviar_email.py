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

def montar_corpo(resumo, data_ref):
    falhas = resumo.get("falhas", [])
    linhas = [
        "Olá,\n",
        "Segue o relatório do Robô CCTs ComuniQ.\n",
        f"Data da execução: {data_ref.strftime('%d/%m/%Y %H:%M')}",
        f"Enviados com sucesso: {len(resumo.get('sucesso', []))}",
        f"Já existiam no ComuniQ: {len(resumo.get('ja_existe', []))}",
        f"Com falha (mantidos em Reprocessar\\): {len(falhas)}",
    ]

    if resumo.get("erro_fatal"):
        linhas.append(f"\nERRO GERAL: {resumo['erro_fatal']}")

    if falhas:
        linhas.append("\nArquivos com falha:")
        linhas.extend(f" - {f['arquivo']}: {f['motivo']}" for f in falhas)

    linhas.append("\nLog completo em anexo.\n\nAtt,\nRobô CCTs ComuniQ")
    return "\n".join(linhas)

################################################################

def enviar_relatorio(resumo, data_execucao=None, arquivo_log=None):
    # Trata os próprios erros: registra no log e devolve False em vez de interromper o robô
    try:
        remetente, senha, destinatarios = carregar_config_email()
    except (FileNotFoundError, KeyError, ValueError) as e:
        _log_erro_email(arquivo_log, f"[EMAIL] Erro ao carregar configuração: {e}")
        return False

    data_ref = data_execucao if data_execucao else datetime.now()
    tem_falha = bool(resumo.get("falhas") or resumo.get("erro_fatal"))

    msg = MIMEMultipart()  # mensagem com corpo de texto + anexo. sem multipart teria só um dos dois
    msg["From"] = remetente
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = f"{'[FALHAS] ' if tem_falha else ''}Relatório Robô CCTs ComuniQ - {data_ref.strftime('%d/%m/%Y')}"
    msg.attach(MIMEText(montar_corpo(resumo, data_ref), "plain", "utf-8"))

    if arquivo_log and os.path.exists(arquivo_log):
        try:
            with open(arquivo_log, "rb") as f:
                parte = MIMEBase("text", "plain")
                parte.set_payload(f.read())
                encoders.encode_base64(parte)
                parte.add_header(
                    "Content-Disposition",
                    f"attachment; filename=\"{os.path.basename(arquivo_log)}\""
                )
                msg.attach(parte)
        except Exception as e:
            _log_erro_email(arquivo_log, f"[EMAIL] Erro ao anexar log (email segue sem anexo): {e}")

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
