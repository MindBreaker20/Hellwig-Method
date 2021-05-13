import pandas as pd
import numpy as np
import math
import statistics
from itertools import product
from pandas.core.arrays.sparse import dtype

#Deklaracja używanych danych
file_name = 'example.xlsx' #nazwa dokumentu Excel o rozszerzeniu xslx
sheet = 'Arkusz1' #nazwa arkusza z tabelą

data1  = pd.read_excel(file_name, sheet_name = sheet, header = None, engine = 'openpyxl') #pobrane dane z arkusza przy czym 1 wiersz to nazwy,
                                                                                          #ponieważ z funkcją read_excel są trudne do pozyskania
data2  = pd.read_excel(file_name, sheet_name = sheet, engine = 'openpyxl') #faktyczne dane 

y = data1.iloc[0][0] #nazwa zmiennej objasnianej
var_number = len(data1.loc[0]) #liczba wszystkich zmiennych
variables = [] #wszystkie zmienne objasniajace spelnaijace Vj > 10%
var_cor = [] #wszystkie współczynniki korelacji miedzy zmiennymi objaśniajacymi a objaśnianą
H = [] #integralne wskaźniki integralności informacyjnej
var_win = [] #najlesze zmienne objaśniające

#ETAP 1 - wspolczynnik zmiennosci Vj > 10%
i = 1
for c in range(1, var_number):  #dla każdej kolumny poza kolumną 1 ze zienna objaśnianą
    xi = data1.loc[:,i]   #wszystkie wartosci liczbowe z kolumny
    xi = xi.loc[1:,]     #nie branie pod uwagę 1 wiersza z wartościami nieliczbowymi (nazwy kolumn)
    mean = statistics.mean(xi) #średnia zmiennej
    sd = statistics.stdev(xi)  #odchylenie standardowe zmiennej
    vj = sd/mean  #współczynnik zmienności
    atr = data1.iloc[0][i] #atrybut zmiennej, która spełnia warunek Vj > 10%
    if vj > 0.1:
        variables.append(atr) #dodawania atrybutów poprawnych zmiennych do listy 
    i += 1

#ETAP 2 - macierz binarna reprezentująca możliwe kombinacje
m = len(variables) #m jest liczbą zmiennych objaśniajacych, kolumny w macierzy binarnej
S = 2**m -1 #liczba kombinacji przepuszczonych zmiennych objaśniających, liczba wierszy macierzy binarnej
binary_matrix = [i for i in product(range(2), repeat = m)] #dwie linijki kodu generujące dowolną macierz z wszystkimi kombinacjami 0-1
binary_matrix = np.array(binary_matrix)
binary_matrix = np.delete(binary_matrix, 0, axis = 0) #usunięcie 1 wiersza z wygenerowanymi 0

#ETAP 3 - macierz korelacji reprezentująca bezwzgledne korelacje między wszystkimi zmiennymi odniesienie do |rij|
data3 = pd.DataFrame(data2)
corr_frame = np.absolute(data3.corr()) #wartości bezwglądne współczynników korelacji
corr_matrix = corr_frame.values #zamiana ramki na macierz
corr_matrix = corr_matrix.round(decimals = 6, out = None)

for i in range(1, m+1): #przekazanie współczynników korelacji między xi a y rj - potrzebne do obliczania h
    var_cor.append(corr_matrix[0][i])

corr_matrix = np.delete(corr_matrix, 0, axis = 0) #usuniecie pierwszego wiersza i kolumny z korelacją x i y, ponieważ potem jest                                 
corr_matrix =np.delete(corr_matrix, 0, axis = 1) #obliczanie sumy rij

#ETAP 4 - obliczanie indywidualnych wskaźników pojemności informacyjnej h
final_matrix = binary_matrix
final_matrix = final_matrix.astype('float64') 
for i in range(0, S):
    for j in range(0, m):
        if final_matrix[i][j] == 1 :
            h = np.square(var_cor[j]) / np.sum(final_matrix[i].dot(corr_matrix[j]))
            h = h.round(decimals = 6, out = None)
            final_matrix[i][j] = h

#ETAP 5 - obliczanie integralnych wskaźników pojemnosci informacyjnej
for i in range(0, S): #kolejne H są dodawane do listy, potem na podstawie pozycji znajdzie się index
    s = np.sum(final_matrix[i]) #najlepszej kombinacji zmiennych objaśniających
    s = round(s, 6)
    H.append(s)

max_H = max(H) #najwyzsze H
idx = 0 #poszukiwany index
while True: #przeszukiwanie listy H w poszukiwaniu pozycji max H
    if H[idx] == max_H:
        break
    idx += 1

#ETAP 6 - znalezienie najlepszej kombinacji i prezentacja wyników
binary_matrix = binary_matrix.astype('float64')
for i in range(0, m): #wypełnianie listy nazwami zmiennych z najlepszej kombinacji
    if binary_matrix[idx][i] == 1:
        var_win.append(variables[i])

print(f"Optymalnym zbiorem zmiennych objaśniających jest kombinacja C{idx}")
for i in range(len(var_win)):
    print(var_win[i])
    












