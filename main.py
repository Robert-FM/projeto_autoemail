from imap_tools import MailBox, AND
import os

# Configurações
username = os.getenv("EMAIL_USERNAME", "robert.fis.fernandes@gmail.com")
password = os.getenv("EMAIL_PASSWORD", "keij raqv cnit azua")  # Melhor usar senha de app do Gmail
output_dir = "anexos"  # Pasta para salvar os anexos e o conteúdo do e-mail

# Criar a pasta de saída, se não existir
os.makedirs(output_dir, exist_ok=True)

# Conectar ao Gmail
with MailBox('imap.gmail.com').login(username, password) as meu_email:
    # Buscar e-mails de um remetente específico
    print("Buscando e-mails com anexos do remetente...")
    emails_busca = meu_email.fetch(AND(from_="mariaritafisica@gmail.com"))
    
    for email in emails_busca:
        # Salvar o conteúdo do e-mail, caso necessário
        email_path = os.path.join(output_dir, f"{email.subject}.txt")
        with open(email_path, 'w', encoding='utf-8') as email_file:
            email_file.write(f"De: {email.from_}\n")
            email_file.write(f"Assunto: {email.subject}\n")
            email_file.write(f"Data: {email.date}\n")
            email_file.write(f"Conteúdo:\n{email.text}\n")
        print(f"E-mail salvo: {email_path}")

        # Salvar os anexos do e-mail
        if email.attachments:
            print(f"\nE-mail: {email.subject}")
            for anexo in email.attachments:
                print(f"Tipo de conteúdo do anexo: {anexo.content_type}")  # Depuração
                
                # Filtrar tipos de arquivo desejados (Excel, PDF, imagens, DOCX, PPT)
                if anexo.content_type in [
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",  # Excel (.xlsx)
                        "application/pdf",  # PDF
                        "image/jpeg",  # Imagem JPEG
                        "image/png",  # Imagem PNG
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",  # DOCX
                        "application/vnd.openxmlformats-officedocument.presentationml.presentation",  # PPT
                ]:
                    file_path = os.path.join(output_dir, anexo.filename)
                    with open(file_path, 'wb') as arquivo:
                        arquivo.write(anexo.payload)
                    print(f"Anexo salvo: {file_path}")