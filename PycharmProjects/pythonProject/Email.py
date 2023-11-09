import csv
import smtplib
import email.message
from datetime import datetime
import time
from tqdm import tqdm
from PyQt5.QtWidgets import QVBoxLayout, QWidget, QHBoxLayout, QLabel, QTextEdit, QPushButton, QApplication, QFileDialog
from PyQt5.QtCore import pyqtSignal, Qt
from PyQt5.QtGui import QFont

class MyWindow(QWidget):
    # Adiciona o sinal send_email_signal
    send_email_signal = pyqtSignal(str, str)

    def __init__(self):
        super().__init__()

        # Cria a interface
        self.init_ui()

        # Conecta o sinal da interface à função send_email_wrapper
        self.send_email_signal.connect(self.send_email_wrapper)

    def init_ui(self):
        # Define o layout principal
        vbox = QVBoxLayout()

        # Adciona uma label com texto normal
        my_label = QLabel("ENVIO DE 1 E-MAIL A CADA 60 SEGUNDOS COM LIMITE DE DE 1440 E-MAILS POR DIA. POR FAVOR, "
                          "CLIQUE APENAS UMA VEZ EM 'ENVIAR' CASO CONTRÁRIO O E-MAIL SERÁ ENVIADO VÁRIAS VEZES")
        vbox.addWidget(my_label)

        # Adciona uma label para digitar o assunto do email
        hbox3 = QHBoxLayout()
        label_subject = QLabel('Assunto do e-mail:')
        hbox3.addWidget(label_subject)
        self.edit_subject = QTextEdit()
        hbox3.addWidget(self.edit_subject)
        vbox.addLayout(hbox3)

        # Adiciona um label para o modelo HTML
        hbox1 = QHBoxLayout()
        label_html = QLabel('Modelo HTML:')
        hbox1.addWidget(label_html)
        self.edit_html = QTextEdit()
        hbox1.addWidget(self.edit_html)
        button_html = QPushButton('Selecionar arquivo')
        button_html.clicked.connect(self.select_html)
        hbox1.addWidget(button_html)
        vbox.addLayout(hbox1)

        # Adiciona um label para a lista de destinatários
        hbox2 = QHBoxLayout()
        label_recipients = QLabel('Lista de destinatários:')
        hbox2.addWidget(label_recipients)
        self.edit_recipients = QTextEdit()
        hbox2.addWidget(self.edit_recipients)
        button_recipients = QPushButton('Selecionar arquivo')
        button_recipients.clicked.connect(self.select_recipients)
        hbox2.addWidget(button_recipients)
        vbox.addLayout(hbox2)

        # Adiciona o botão para enviar o email
        self.button_send = QPushButton('Enviar email')
        self.button_send.clicked.connect(self.send_email)
        vbox.addWidget(self.button_send)

        self.setLayout(vbox)
        self.setWindowTitle('Enviar Email')

        # Adciona uma label com texto normal
        my_label = QLabel("                                                                                                                                      2023 - by Inovação Barueri")
        font = QFont()
        font.setBold(True)
        font.setItalic(True)
        my_label.setFont(font)
        vbox.addWidget(my_label)


    def select_html(self):
        # Abre uma janela de seleção de arquivo para o modelo HTML
        html_file, _ = QFileDialog.getOpenFileName(self, 'Selecionar arquivo', '.', 'Arquivos HTML (*.html)')
        self.edit_html.setText(html_file)

    def select_recipients(self):
        # Abre uma janela de seleção de arquivo para a lista de destinatários
        recipients_file, _ = QFileDialog.getOpenFileName(self, 'Selecionar arquivo', '.', 'Arquivos CSV (*.csv)')
        self.edit_recipients.setText(recipients_file)

    def send_email(self):
        # Obtém os arquivos selecionados na interface
        html_file = self.edit_html.toPlainText()
        recipients_file = self.edit_recipients.toPlainText()

        # Emite o sinal send_email_signal com os arquivos selecionados
        self.send_email_signal.emit(html_file, recipients_file)

    def send_email_wrapper(self, html_file, recipients_file):
        # Lê o modelo HTML do arquivo selecionado na interface
        with open(html_file, 'r', encoding='utf-8') as f:
            corpo_email = f.read()



        # Lendo a lista de destinatários a partir do arquivo selecionado na interface
        recipients = []
        with open(recipients_file, 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                recipients.append(row[0])



        # Loop para enviar o e-mail para cada destinatário
        for recipient in tqdm(recipients):
            try:
                # Criar a mensagem de e-mail
                msg = email.message.Message()
                msg['Subject'] = self.edit_subject.toPlainText()
                msg['From'] = 'seu email'
                msg['To'] = recipient
                password = 'sua senha'
                msg.add_header('Content-Type', 'text/html; charset=utf-8')
                msg.set_payload(corpo_email)
                # Gerar Relatório de envio

                with open('registros.csv', 'a', newline='') as file:
                    csv_writer = csv.writer(file)
                    csv_writer.writerow(
                        [recipient, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'Enviado com sucesso'])

            except Exception as e:
                # Registrar o erro ocorrido durante o envio
                with open('registros.csv', 'a', newline='') as file:
                    csv_writer = csv.writer(file)
                    csv_writer.writerow(
                        [recipient, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'Falha ao enviar: ' + str(e)])

            # Enviar a mensagem
            s = smtplib.SMTP('smtp.seuserver', 587)
            s.starttls()
            # Login Credentials for sending the mail
            s.login(msg['From'], password)
            s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))


            time.sleep(60)



if __name__ == '__main__':
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec_()

print("E-mails enviados")
