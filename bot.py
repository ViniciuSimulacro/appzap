import pywhatkit
import keyboard
import time
import win32clipboard
from datetime import datetime
import openpyxl

def resultados():
    book = openpyxl.load_workbook('uploads/contatos.xlsx')
    contacts = book['Contatos']
    send_list = []
    for rows in contacts.iter_rows(min_row=2):
        send_list.append([rows[0].value,rows[1].value])
    return send_list

"""def send_msg():
    contatos = []

    for item in send_list:
        contatos.append(item)

    msg_imagem = '*Haduken*'
    imagem = 'uploads/haduken.jpeg'

    while len(contatos)>= 1:
        telefone = contatos[0][0]
        mensagem = f'Teste feito pelo Simulacro Web, gostou {contatos[0][1]}?'
        pywhatkit.sendwhats_image(contatos[0][0], imagem , msg_imagem,20)
        time.sleep(5)
        keyboard.press_and_release('ctrl + W')
        pywhatkit.sendwhatmsg(contatos[0][0],mensagem,datetime.now().hour, datetime.now().minute+1)
        time.sleep(5)
        keyboard.press_and_release('ctrl + W')
        del contatos[0]
    return 'missão concluida!'"""