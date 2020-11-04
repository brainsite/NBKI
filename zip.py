# coding: utf8
import subprocess


def zip_ext(patch_arhiv, filename, out_zip):
    extract_7zip = '"C:\\Program Files\\7-Zip\\7z.exe" e '
    # patch_arhiv='C:\\ASTRA\\Mailbox\\MGTU\\A\\2015-05-15\\R\\'

    # filename='письма.zip'
    print(extract_7zip + patch_arhiv + filename + ' o' + out_zip)
    p = subprocess.Popen(extract_7zip + patch_arhiv + filename + ' -o' + out_zip, shell=True, stdout=subprocess.PIPE)
    out = p.stdout.read()
    print (out)
    result = out.split()
    for z in result:
        print(z)


def zip_add(patch, filename):
    add_7zip = '"C:\\Program Files\\7-Zip\\7z.exe" a '
    print(add_7zip + patch + filename + ' ' + patch + '*')
    p = subprocess.Popen(add_7zip + patch + filename + ' ' + patch + '*', shell=True, stdout=subprocess.PIPE)
    out = p.stdout.read()
    print (out)
    result = out.split()
    for z in result:
        print(z)
