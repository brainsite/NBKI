# coding: utf8
__author__ = 'smirnov'

import os
import site
import subprocess
import sys
import time

import win32com.client

sys.stdout = open('C:\\nbki\\data\\log.txt', 'w')

true_crypt = 'C:\\Program Files\\TrueCrypt\\truecrypt.exe '
mount_volume = ' /lz /m rm /a /p ****pass**** /q'

# Образы дисков#
disk_nbki = '"C:\\nbki\\data\\nbki.tc"'


def time_now():
    return time.strftime('%Y-%m-%d %H.%M.%S') + '   '


def mount():
    try:
        os.listdir('z:\\')
        print(time_now() + ' диск на месте')
    except BaseException:
        subprocess.call(true_crypt + disk_nbki + mount_volume)
        print(time_now() + ' подмонтирован nbki')


print(time_now() + '         begin                 ')

site.addsitedir('c:\\nbki\\')

BASE_DIR = 'c:\\nbki'
from zip import zip_ext, zip_add

print(BASE_DIR)
print(time_now() + ' назначение переменных')

crypto_pro = 'cryptcp.exe '
tmp_path = BASE_DIR + '\\tmp\\'
in_path = BASE_DIR + '\\in\\'
out_path = BASE_DIR + '\\out\\'

send_list = ['reglament@in-bank.ru',
             'petrova@in-bank.ru; sidorova@in-bank.ru; ivanova@in-bank.ru']
# send_list=['reglament@in-bank.ru; a.smirnov@in-bank.ru']

print(time_now() + ' переменные назначили')


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


# подпись и архивация файла
def sign(file):
    sign_msg = '-signf -dn "Петрова Наталья Николаевна,nbki-exchange@in-bank.ru" -der -cert ' + tmp_path + file + ' -dir ' + tmp_path
    print(sign_msg)
    print(crypto_pro + sign_msg)
    subprocess.call(crypto_pro + sign_msg, shell=True)
    try:
        os.rename(tmp_path + file + '.sgn', tmp_path + file + '.sig')
    except:
        print('Ошибка переименования файла.')
    zip_add(tmp_path, file + '.zip')
    return file + '.zip'


# шифрование файла
def encrypt(file):
    if file.find('csv') != -1:
        csv_name = file[:file.find('.csv')] + file[file.find('.csv') + 4:]
        encrypt_msg = '-encr -dn "Оператор НБКИ - 2016,support@nbki.ru" -dn "Смирнов Алексей Сергеевич,smirnov@in-bank.ru" -der ' + tmp_path + file + ' ' + tmp_path + csv_name + '.enc'
        file = csv_name
    else:
        encrypt_msg = '-encr -dn "Оператор НБКИ - 2016,support@nbki.ru" -dn "Смирнов Алексей Сергеевич,smirnov@in-bank.ru" -der ' + tmp_path + file + ' ' + tmp_path + file + '.enc'
    subprocess.call(crypto_pro + encrypt_msg, shell=True)  # подписываем файлы
    return file + '.enc'


# расшифровка файла
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


# Проверка почты
print(time_now() + ' проверка почты')

try:
    print(time_now() + ' обращение к Outlook')
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    print(time_now() + ' r Outlook обратились')
    print(time_now() + ' r обращение к ящику Inbox')
    inbox = outlook.GetDefaultFolder(6)
    print(inbox.Name)
    messages = inbox.Items
    # print(messages)
    messages.Sort('Отправлено', True)
    print('step1')
    for z in messages:
        print('step2')
        if (z.ReceivedTime).strftime("%m%d%Y") == (time.strftime('%m%d%Y')):
            print(z.SenderEmailAddress)
            print(z.Subject)
            # if z.SenderEmailAddress=='credithistory@nbki.ru':
            # if z.SenderEmailAddress=='smirnov_as@delfaro.ru':
            if str(z.Subject).replace(' ', '').lower() == 'отправить в нбки'.replace(' ',
                                                                                     '').lower() or z.SenderEmailAddress == 'credithistory@nbki.ru':
                print('smirnov')
                try:
                    print(z.SenderEmailAddress)
                    for x in z.Attachments:
                        print(str(x.FileName))
                        x.SaveAsFile(in_path + x.FileName)
                except:
                    print(sys.exc_info())
                    pass
        else:
            break
except:
    print('Ошибка       ' + time_now())
    print(sys.exc_info())

print(time_now() + ' Расшифровка/Шифровка сообщений из папки in')

# Расшифровка/Шифровка сообщений из папки in
in_file = os.listdir(in_path)
file_name = ''
if in_file != []:
    mount()
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
            ls = os.listdir('C:\\nbki\\out\\' + time.strftime('%Y%m%d') + '\\')
            if x in ls:
                print('Файл отправлен')
            else:
                file_name = encrypt(sign(x))  # формирование сообщения

                try:
                    print('Пробуем отправить почту       ' + time_now())
                    send_mail('send to nbki', 'to_nbki', 'credithistory@nbki.ru', [tmp_path + file_name, ])
                    # send_mail('send to nbki','to_nbki','a.smirnov@in-bank.ru',[tmp_path+file_name,])
                except:
                    print('Ошибка       ' + time_now())
                    print(sys.exc_info())
        if x.find('csv') != -1:
            ls = os.listdir('C:\\nbki\\out\\' + time.strftime('%Y%m%d') + '\\')
            if x in ls:
                print('Файл отправлен')
            else:
                file_name = encrypt(sign(x))  # формирование сообщения
                try:
                    print('Пробуем отправить почту       ' + time_now())
                    send_mail('send to nbki', 'to_nbki', 'cancelcredithistory@nbki.ru', [tmp_path + file_name, ])
                except:
                    print('Ошибка       ' + time_now())
                    print(sys.exc_info())

        else:
            list_out = os.listdir(folder_today(out_path))
            try:
                list_out.index(x)
                print(list_out)
            except:
                decrypt(x)  # расшифровка сообщения
                for y in os.listdir(tmp_path):
                    if y[-6:] == 'ticket':
                        for ls in send_list:
                            try:
                                print('Пробуем отправить почту       ' + time_now())
                                send_mail('Квитанция из НБКИ', 'Квитанция из НБКИ', ls, [tmp_path + y, ])
                            except:
                                print('Ошибка       ' + time_now())
                                print(sys.exc_info())
        try:
            for y in os.listdir(tmp_path):
                os.replace(tmp_path + y, folder_today(out_path) + y)
        except:
            print(sys.exc_info())
