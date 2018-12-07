from datetime import datetime, timedelta
import os
import xlrd
from xlrd import open_workbook
from xlutils.copy import copy

lista = []
czasy = []
suma= timedelta(seconds = 0)

print('\nDostepne pliki:')
for file in os.listdir('Z:/emmegifdd/ETX/log/'):
    print(file[:len(file) - 4])

while True:

    plik = input('\nWybierz plik z listy: ')
    fileLog = open('Z:/emmegifdd/ETX/log/' + plik + '.log', 'r').read()
    lines = fileLog.split('\n')
    
    for i in range(5,len(lines),6):
        lista.append(datetime.strptime(lines[i][:19], '%Y-%m-%d %H:%M:%S'))
        
    for i in range(1, len(lista)):
        czasy.append(lista[i] - lista[i-1])
    
    for i in range(len(czasy)):
        if czasy[i] < timedelta(hours = 8) and czasy[i] > timedelta(seconds = 0):
            suma += czasy[i]

    print('\nSredni czas:             ' + str(suma/len(czasy)))
    print('\nIlosc ucietych profili:  ' + str(len(czasy)+1))
    print('\nData rozpoczecia:        ' + str(min(lista)))
    print('\nData zakonczenia:        ' + str(max(lista)))
    print('\nCzas pracy:              ' + str((len(czasy)+1)*(suma/len(czasy))))

    try:
        book_ro = xlrd.open_workbook("raport.xls")
        book = copy(book_ro)
        sheet1 = book.get_sheet(0)
        worksheet = book_ro.sheet_by_name('Arkusz1')
        num_rows = worksheet.nrows - 1

        if(num_rows == (-1)):
            sheet1.write(0, 0, 'Zlecenie:')
            sheet1.write(0, 1, 'Sredni czas:')
            sheet1.write(0, 2, 'Ilosc ucietych profili:')
            sheet1.write(0, 3, 'Data rozpoczecia:')
            sheet1.write(0, 4, 'Data zakonczenia:')
            sheet1.write(0, 5, 'Czas pracy:')

            sheet1.write(num_rows + 2, 0, plik)
            sheet1.write(num_rows + 2, 1, (str(suma/len(czasy)))[0:7])
            sheet1.write(num_rows + 2, 2, str(len(czasy)+1))
            sheet1.write(num_rows + 2, 3, str(min(lista)))
            sheet1.write(num_rows + 2, 4, str(max(lista)))
            sheet1.write(num_rows + 2, 5, (str((len(czasy)+1)*(suma/len(czasy))))[0:7])
        else:
            sheet1.write(num_rows + 1, 0, plik)
            sheet1.write(num_rows + 1, 1, (str(suma / len(czasy)))[0:7])
            sheet1.write(num_rows + 1, 2, str(len(czasy) + 1))
            sheet1.write(num_rows + 1, 3, str(min(lista)))
            sheet1.write(num_rows + 1, 4, str(max(lista)))
            sheet1.write(num_rows + 1, 5, (str((len(czasy) + 1) * (suma / len(czasy))))[0:7])


        try:
            book.save("raport.xls")
            print('Dane zapisane do pliku raport.xls')
        except:
            print('Blad zapisu / plik jest uzywany przez inna osobe!')
    except:
        print('Brak pliku raport.xls, utworz pusty plik o nazwie raport.xls')

    lista = []
    czasy = []
    suma = timedelta(seconds=0)