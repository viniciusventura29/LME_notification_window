import win32com.client as cli
from PyQt5 import uic,QtWidgets
import os
import subprocess

app = QtWidgets.QApplication([])

tela_notificacao = uic.loadUi('./janelas/telanotificacao.ui')

path = os.path.expanduser(r"C:\Users\ct67ca\Desktop\LME_desktop\planilha")

def getExcel():
    outlook = cli.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace("MAPI")

    acc = namespace.Folders['Campinas.ETS@br.bosch.com']
    inbox = acc.Folders('Inbox')

    all_inbox = inbox.Items

    for msg in all_inbox:
        if msg.Class == 43:
            if msg.SenderEmailType == 'EX':
                pass
            else:
                if msg.SenderEmailAddress == "viniciusventura29@icloud.com" and msg.Subject == "Atualização excel LME":
                    
                    attachments = msg.Attachments
                    attachment = attachments.Item(1)
                    
                    for attachment in msg.Attachments:
                        attachment.SaveAsFile(os.path.join(path,"LME_media_mensal.xlsx" ))
                        
                        
                    tela_notificacao.show()
                    msg.Move(acc.Folders('Itens Excluídos'))
                else:
                    pass
                

def fechar_janela():
    tela_notificacao.close()
    
def open_app():
    subprocess.Popen(r'C:\Users\ct67ca\Desktop\LME_desktop\main.exe')
    fechar_janela()
                
tela_notificacao.abrir_app.clicked.connect(open_app)

tela_notificacao.fechar.clicked.connect(fechar_janela) 



getExcel()    
app.exec()