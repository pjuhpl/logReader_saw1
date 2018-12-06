from datetime import datetime, timedelta
import os

lista = []
czasy = []
suma= timedelta(seconds = 0)

print('\nDostepne pliki:')
for file in os.listdir('Z:/emmegifdd/ETX/log/'):
    print(file[:len(file)-4])

while True:
    plik = input('\nWybierz plik: ')
    fileLog = open('Z:/emmegifdd/ETX/log/' + plik + '.log', 'r').read()
    lines = fileLog.split('\n')
    
    for i in range(5,len(lines),6):
        lista.append(datetime.strptime(lines[i][:19], '%Y-%m-%d %H:%M:%S'))
        
    for i in range(1, len(lista)):
        czasy.append(lista[i] - lista[i-1])
    
    for i in range(len(czasy)):
        if czasy[i] < timedelta(hours = 8) and czasy[i] > timedelta(seconds = 0):
            suma += czasy[i]
    
    print('\nSredni czas:            ' + str(suma/len(czasy)))   
    print('\nIlosc ucietych profili: ' + str(len(czasy)))
    lista = []
    czasy = []
    suma= timedelta(seconds = 0)