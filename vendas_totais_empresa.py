# -*- coding: utf-8 -*-
"""
Created on Tue Aug 11 19:06:43 2020

@author: User-PC
"""

# -*- coding: utf-8 -*-
"""
Spyder Editor

Este é um arquivo de script temporário.
"""

import glob
import xlrd
import xlsxwriter
   
def escrever_linha_xlsx(new_worksheet,n,linha_dados):
    
    new_worksheet.write(n,0,'01/01/2020 a 31/01/2020')
    #new_worksheet.write(n,1,empresa)
    
    i = 0
    for dado in linha_dados:
        new_worksheet.write(n,1+i,dado)
        i = i + 1
    

#arquivo fonte dos dados
mes = "11-2019"        
#mes = "05-2020"
file_name = "GRN/Venda de Itens por Categoria " + mes +".xlsx" 
workbook = xlrd.open_workbook(file_name)
sheet = workbook.sheet_by_index(0)
    
    #arquivo a ser gerado
new_file_name = "xt2/media_vendas_total " + mes + ".xlsx"   
new_workbook = xlsxwriter.Workbook(new_file_name)
new_worksheet = new_workbook.add_worksheet()
    
L = 0
cont = 0
linha = []
linha_comb = []
    
while(L != sheet.nrows):
        
    cell = sheet.cell(L,0)
    if(('Total Categoria: COMBUSTIVEIS' in cell.value) == True):
            linha_comb.append(sheet.cell(L,4).value)#venda bruta total
            linha_comb.append(sheet.cell(L,9).value)#venda liquida total
            
    elif(('Total Empresa: ' in cell.value) == True):
            #categoria
            linha.append(cell.value.strip('Total Empresa: '))
            #totais
            linha.append(sheet.cell(L,4).value - linha_comb[0])#venda bruta total - venda de combustiveis
            linha.append(sheet.cell(L,9).value - linha_comb[1])#venda liquida total - venda de combustiveis
            escrever_linha_xlsx(new_worksheet, cont, linha)
            print(linha)
            linha = []
            linha_comb = []
            cont = cont + 1
            
    L  = L + 1    
         
new_workbook.close()