import pandas as pd
import win32com.client
import tkinter as tk
from tkinter import filedialog
import os
import sys
from datetime import datetime

def selecionar_arquivo_csv():
    root = tk.Tk()
    root.withdraw()
    arquivo_csv = filedialog.askopenfilename(
        title="Selecione o arquivo CSV",
        filetypes=[("Arquivos CSV", "*.csv")],
    )
    return arquivo_csv


def obter_caminho_assinatura():
    if getattr(sys, 'frozen', False): 
        diretorio_base = sys._MEIPASS  
    else:
        diretorio_base = os.path.dirname(os.path.abspath(__file__))  

    return os.path.join(diretorio_base, "assinatura.jpeg")  


def enviar_email(destinatario, assunto, corpo, assinatura_path):
    outlook = win32com.client.Dispatch("Outlook.Application")
    email = outlook.CreateItem(0)

    email.To = destinatario
    email.Subject = assunto
    email.HTMLBody = f"""{corpo}
    <br><br>
    <img src="cid:minha_assinatura">
    """

    if not os.path.exists(assinatura_path):
        print(f"Erro: imagem de assinatura não encontrada em: {assinatura_path}")
        return

    assinatura = email.Attachments.Add(assinatura_path)
    assinatura.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "minha_assinatura")

    email.Send()

    print(f"E-mail enviado para {destinatario}")


def processador_boletos(arquivo_csv):
    if not arquivo_csv:
        print("Nenhum arquivo CSV selecionado")
        return
    
    try:
        df = pd.read_csv(arquivo_csv)
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return
    
    assinatura_path = obter_caminho_assinatura()
    vencimento = datetime.now().strftime("%d/%m/%Y")

    for _, row in df.iterrows():
        email = row["Email"]

        assunto = f"Fatura Medicina do Trabalho {vencimento}"
        corpo = f"""
        <p>Olá Amigos,</p>

        <p>Esperamos que esteja bem.</p>

        <p>Este é um lembrete amigável de que o pagamento referente à fatura de Medicina do Trabalho vence hoje, <b>{vencimento}</b>.</p>

        <p>Pedimos a gentileza de efetuar o pagamento até o final do dia para evitar qualquer inconveniente. Caso já tenha realizado o pagamento, por favor, desconsidere este e-mail.</p>

        <p>Se precisar de suporte, estamos à disposição.</p>

        <p>Agradecemos pela sua atenção e colaboração.</p>

        <p>Atenciosamente,</p>
        """

        enviar_email(email, assunto, corpo, assinatura_path)

if __name__ == "__main__":
    arquivo_csv = selecionar_arquivo_csv()

    processador_boletos(arquivo_csv)