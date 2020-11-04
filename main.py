# coding: utf8
__author__ = 'smirnov'

import logging
import os
import subprocess
import sys
import time

import win32com.client

from zip import zip_ext, zip_add

BASE_DIR = os.path.dirname(os.path.dirname(__file__))

crypto_pro = 'cryptcp.exe '  # Приложение командной строки для создания запросов на сертификаты, шифрования и расшифрования файлов, создания и проверки электронной подписи файлов с использованием сертификатов открытых ключей, хэширования файлов

# Директория временных файлов
tmp_path = BASE_DIR + '/tmp/'

# Директория входящих файлов
in_path = BASE_DIR + '/in/'

# Директория исходящих файлов
out_path = BASE_DIR + '/out/'

# Каталог с логами
log_dir = '/logs/'

logging.basicConfig(format=u'%(levelname)-8s [%(asctime)s] %(message)s', level=logging.DEBUG,
                    filename=log_dir + time.strftime("%Y%m%d") + 'log.log'.encode('utf-8'))


def folder_today(patch):
    folder = (patch + time.strftime('%Y%m%d') + '\\')
    try:
        os.mkdir(folder)
        print('Folder created')
    except BaseException:
        print ('Folder2 not created')
    return folder


####################################################################################################################################################
# В функции указаны реквизиты "CN" и "E" сертификатов. Если их в локальном справочнике больше одного с одним названием то лишние необходимо удалить.#
####################################################################################################################################################

# Подпись и архивация файла
def sign(file):
    sign_msg = '-signf -dn "Смирнов Алексей Сергеевич,smirnov@in-bank.ru" -der -cert ' + tmp_path + file + ' -dir ' + tmp_path
    print(sign_msg)
    print(crypto_pro + sign_msg)
    subprocess.call(crypto_pro + sign_msg, shell=True)
    try:
        os.rename(tmp_path + file + '.sgn', tmp_path + file + '.sig')
    except:
        logging.error(u'Ошибка переименования файла. Причина:' + str(sys.exc_info()))
    zip_add(tmp_path, file + '.zip')
    return file + '.zip'


# Шифрование файла
def encrypt(file):
    if file.find('csv') != -1:
        csv_name = file[:file.find('.csv')] + file[file.find('.csv') + 4:]
        encrypt_msg = '-encr -dn "Оператор НБКИ - 2017,support@nbki.ru" -dn "Смирнов Алексей Сергеевич,smirnov@in-bank.ru" -der ' + tmp_path + file + ' ' + tmp_path + csv_name + '.enc'
        file = csv_name
    else:
        encrypt_msg = '-encr -dn "Оператор НБКИ - 2017,support@nbki.ru" -dn "Смирнов Алексей Сергеевич,smirnov@in-bank.ru" -der ' + tmp_path + file + ' ' + tmp_path + file + '.enc'
    subprocess.call(crypto_pro + encrypt_msg, shell=True)  # подписываем файлы
    return file + '.enc'


# Расшифровка файла
def decrypt(file):
    out_file = file[:-4]
    decrypt_msg = '-decr -dn "Смирнов Алексей Сергеевич,smirnov@in-bank.ru" ' + tmp_path + file + ' ' + tmp_path + out_file
    print(decrypt_msg)
    subprocess.call(crypto_pro + decrypt_msg, shell=True)  # подписываем файлы
    zip_ext(tmp_path, out_file, tmp_path)


# Отправка сообщений по почте
def send_mail(text, subject, recipient, attach):
    o = win32com.client.Dispatch("Outlook.Application")
    Msg = o.CreateItem(0)
    Msg.To = recipient
    Msg.Subject = subject
    Msg.Body = text
    attachment1 = attach
    for f in attachment1:
        Msg.Attachments.Add(f)
    Msg.Send()


# Проверка почты через Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort('Отправлено', True)
for z in messages:
    if (z.ReceivedTime).strftime("%Y%m%d") == (time.strftime('%Y%m%d')):
        if z.SenderEmailAddress == 'credithistory@nbki.ru':
            try:
                for x in z.Attachments:
                    print(str(x.FileName))
                    x.SaveAsFile(in_path + x.FileName)
            except:
                logging.error(u'Ошибка проверки почты. Причина:' + str(sys.exc_info()))
    else:
        break

# Расшифровка/Шифровка сообщений из папки in

in_file = os.listdir(in_path)
file_name = ''
if in_file != []:
    for x in in_file:
        try:
            for y in os.listdir(tmp_path):
                os.remove(tmp_path + y)
        except:
            print(sys.exc_info())
        try:
            os.replace(in_path + x, tmp_path + x)
        except:
            print(sys.exc_info())
        if x.find('.') == -1:
            file_name = encrypt(sign(x))  # формирование сообщения
            send_mail('send to nbki', 'to_nbki', 'credithistory@nbki.ru', [tmp_path + file_name, ])
        if x.find('csv') != -1:
            file_name = encrypt(sign(x))  # формирование сообщения
            send_mail('send to nbki', 'to_nbki', 'cancelcredithistory@nbki.ru', [tmp_path + file_name, ])
        else:
            list_out = os.listdir(folder_today(out_path))
            try:
                list_out.index(x)
                print(list_out)
            except:
                decrypt(x)  # расшифровка сообщения
                for y in os.listdir(tmp_path):
                    if y[-6:] == 'ticket':
                        send_mail('Квитанция из НБКИ', 'Квитанция из НБКИ', 'reglament@in-bank.ru', [tmp_path + y, ])
        try:
            for y in os.listdir(tmp_path):
                os.replace(tmp_path + y, folder_today(out_path) + y)
        except:
            print(sys.exc_info())
