import pandas as pd
import numpy as np
import math
import statistics
from itertools import product
from pandas.core.arrays.sparse import dtype

#Declaration of used data
file_name = 'example.xlsx' #the name of the Excel document with the extension xlsx
sheet = 'Arkusz1' #name of the sheet with the table

data1  = pd.read_excel(file_name, sheet_name = sheet, header = None, engine = 'openpyxl') #name of the sheet with the table, because with the read_excel function they are difficult to get
                                                                                          
data2  = pd.read_excel(file_name, sheet_name = sheet, engine = 'openpyxl') #actual data

y = data1.iloc[0][0] #the name of the target variable
var_number = len(data1.loc[0]) #number of independent variables
variables = [] #all variables satisfying requirement Vj> 10%
var_cor = [] #all correlation coefficients between the explanatory and dependent variables
H = [] #integral indicators of information integrity
var_win = [] #the best explanatory variables

#STEP 1 - coefficient of variation Vj> 10%
i = 1
for c in range(1, var_number):  #for each column except column 1 with the explained variable
    xi = data1.loc[:,i]   #all numeric values in the column
    xi = xi.loc[1:,]     #ignoring 1 row with non-numeric values (column names)
    mean = statistics.mean(xi) #the mean of the variable
    sd = statistics.stdev(xi)  #standard deviation of the variable
    vj = sd/mean  #coefficient of variation
    atr = data1.iloc[0][i] #attribute of a variable that meets the condition Vj> 10%
    if vj > 0.1:
        variables.append(atr) #adding attributes of valid variables to the list
    i += 1

#STEP 2 - binary matrix representing possible combinations
m = len(variables) #m is the number of explanatory variables, columns in a binary matrix
S = 2**m -1 #the number of combinations of the passed explanatory variables, the number of rows of the binary matrix
binary_matrix = [i for i in product(range(2), repeat = m)] #two lines of code generating any matrix with all combinations 0-1
binary_matrix = np.array(binary_matrix)
binary_matrix = np.delete(binary_matrix, 0, axis = 0) #delete 1 row with generated multiple 0

#STEP 3 - correlation matrix representing absolute correlations between all variables reference to | rij |
data3 = pd.DataFrame(data2)
corr_frame = np.absolute(data3.corr()) #absolute values of the correlation coefficients
corr_matrix = corr_frame.values #change frame to matrix
corr_matrix = corr_matrix.round(decimals = 6, out = None)

for i in range(1, m+1): #passing the correlation coefficients between xi and y rj - needed for computing h
    var_cor.append(corr_matrix[0][i])

corr_matrix = np.delete(corr_matrix, 0, axis = 0) #removing the first row and column with the x and y correlation, because then there is                                 
corr_matrix =np.delete(corr_matrix, 0, axis = 1) #calculating sum rij

#STEP 4 - calculation of individual indicators of information capacity h
final_matrix = binary_matrix
final_matrix = final_matrix.astype('float64') 
for i in range(0, S):
    for j in range(0, m):
        if final_matrix[i][j] == 1 :
            h = np.square(var_cor[j]) / np.sum(final_matrix[i].dot(corr_matrix[j]))
            h = h.round(decimals = 6, out = None)
            final_matrix[i][j] = h

#STEP 5 - calculation of integral indicators of information capacity
for i in range(0, S): #subsequent Hs are added to the list, then index will be found based on the position
    s = np.sum(final_matrix[i]) #the best combination of explanatory variables
    s = round(s, 6)
    H.append(s)

max_H = max(H) #the highest H
idx = 0 #index you are looking for
while True: #searches the H list for max H items
    if H[idx] == max_H:
        break
    idx += 1

#STEP 6 - finding the best combination and presenting the results
binary_matrix = binary_matrix.astype('float64')
for i in range(0, m): #filling the list with variable names from the best combination
    if binary_matrix[idx][i] == 1:
        var_win.append(variables[i])

print(f"The optimal set of explanatory variables is the combination C{idx}")
for i in range(len(var_win)):
    print(var_win[i])
    












