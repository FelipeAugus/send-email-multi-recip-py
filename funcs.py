import smtplib
import traceback
import tkinter as tk
from tkinter import filedialog
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

def run():
    janela = tk.Tk()
    janela.geometry('500x457') 
    for i in range(4): janela.columnconfigure(i, weight=1, minsize=100)
    janela.columnconfigure(0, weight=15, minsize=15)
    janela.columnconfigure(5, weight=15, minsize=15)
    janela.rowconfigure(7, minsize=50)

    janela.title("Enviar email")

    # Criação dos campos do formulário
    tk.Label(janela, text="Email:", font=("arial",12)).grid(row=0, column=1, sticky=tk.W)
    entrada_email_remetente = tk.Entry(janela)
    entrada_email_remetente.grid(row=0, column=2, sticky=tk.E)

    tk.Label(janela, text="Senha:", font=("arial",12)).grid(row=0, column=3, sticky=tk.W)
    entrada_senha = tk.Entry(janela, show="*")
    entrada_senha.grid(row=0, column=4, sticky=tk.E)

    tk.Label(janela, text="Destinatários (Separados por ',' ou ';'):", font=("arial",12)).grid(row=1, column=1, columnspan=4, sticky=tk.W)
    entrada_destinatarios = tk.Text(janela, height=3)
    entrada_destinatarios.grid(row=2, column=1, columnspan=4, sticky=tk.E)

    tk.Label(janela, text="Assunto:", font=("arial",12)).grid(row=3, column=1, sticky=tk.W)
    entrada_assunto = tk.Entry(janela, width=300)
    entrada_assunto.grid(row=3, column=2, columnspan=4, sticky=tk.E)

    tk.Label(janela, text="Corpo do email:", font=("arial",12)).grid(row=4, column=1, columnspan=4, sticky=tk.W)
    entrada_corpo = tk.Text(janela, height=10)
    entrada_corpo.grid(row=5, column=1, columnspan=4)

    tk.Label(janela, text="Imagem para assinatura:", font=("arial",12)).grid(row=6, column=1, columnspan=4, sticky=tk.W)
    entrada_imagem_assinatura = tk.Entry(janela, width=300)
    entrada_imagem_assinatura.grid(row=7, column=1, columnspan=3, sticky=tk.W)
    botao_imagem = tk.Button(janela, text="Selecionar imagem", command=lambda: selecionar_imagem(entrada_imagem_assinatura))
    botao_imagem.grid(row=7, column=4, sticky=tk.E)

    # Botão para enviar o email
    botao_enviar = tk.Button(janela, text="Enviar email", command= lambda: trata_e_encaminha_email(
        entrada_email_remetente,entrada_senha,
        entrada_destinatarios,
        entrada_assunto,entrada_corpo,
        entrada_imagem_assinatura
        ))
    botao_enviar.grid(row=9, column=1, columnspan=4)

    janela.lift()
    janela.mainloop()


def selecionar_imagem(entrada_imagem_assinatura):
    caminho_imagem = filedialog.askopenfilename(
        title="Selecionar imagem",
        filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")]
    )
    if caminho_imagem:
        entrada_imagem_assinatura.delete(0, tk.END)
        entrada_imagem_assinatura.insert(0, caminho_imagem)

def trata_e_encaminha_email(entrada_email_remetente, entrada_senha, entrada_destinatarios, entrada_assunto, entrada_corpo, entrada_imagem_assinatura):
    try:
        email_remetente = entrada_email_remetente.get()
        senha_remetente = entrada_senha.get()
        emails_destinatarios = entrada_destinatarios.get("1.0", tk.END).replace("\n", "").replace(" ", "").replace(",", ";").replace(";", ",")
        assunto_email = entrada_assunto.get()
        corpo_email = entrada_corpo.get("1.0", tk.END)
        caminho_assinatura = entrada_imagem_assinatura.get()
        print(f"Você tem certeza que deseja enviar o email:\nAssunto: {assunto_email}\nCorpo: {corpo_email}\nDestinatários: {emails_destinatarios}?")
        if tk.messagebox.askyesno("Confirmação", f"Você tem certeza que deseja enviar o email:\nAssunto: {assunto_email}\nCorpo: {corpo_email}\nDestinatários: {emails_destinatarios}?"):
            print("Encaminhando emails")
    
            mensagem = obtem_email(assunto_email, corpo_email, caminho_assinatura)
            envia_email(mensagem, emails_destinatarios.split(','), email_remetente, senha_remetente)
            print("Envios finalizados")
            tk.messagebox.showinfo("Email enviado", "🦍 Os emails foram enviados com sucesso! 🤠")

        else:
            print("Cancelado o envio de emails")
            tk.messagebox.showinfo("Envio cancelado", "🦧 Os envios foram cancelados! 😶")

    except Exception as e:
        print("Erro ao enviar email: ", e)
        traceback.print_exc()
        tk.messagebox.showerror("Erro no envio", f"Ocorreu um erro ao enviar os emails 😢\nMe manda um print😪\nErro: {e}\n\nrepr: {repr(e)}")

def obtem_email(assunto_email, corpo_email, caminho_assinatura):
    mensagem = MIMEMultipart()
    mensagem['Subject'] = assunto_email
 
    mensagem.attach(MIMEText(corpo_email))

    with open(caminho_assinatura, 'rb') as f:
        img_assinatura = MIMEImage(f.read())
        img_assinatura.add_header('Content-ID', '<assinatura>')
        mensagem.attach(img_assinatura)
    assinatura = '<div><p><img src="cid:assinatura"></p></div>'
    mensagem.attach(MIMEText(assinatura, 'html'))

    return mensagem

def envia_email(mensagem, destinatarios, email_remetente, senha_remetente):    
    mensagem['From'] = email_remetente

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(email_remetente, senha_remetente)

        for destinatario in destinatarios:
            mensagem['To'] = destinatario
            print(f"Enviando email para: {destinatario}")
            server.sendmail(email_remetente, [destinatario], mensagem.as_string())



mensagem = obtem_email("teste", "Teste\nã\nteste", "C:\\Users\\augus\\OneDrive\\Documentos\\Pessoal\\SendMultMail\\emailV1.0.0\\assinatura.png")
envia_email(mensagem, ["augusto.felipao22@gmail.com"], "leticiacamposs127@gmail.com", "vaaioufuggfjiwsf")