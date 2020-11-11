# -*- coding: utf-8 -*-
"""
Created on Tue Aug 11 19:06:43 2020

@author: User-PC
"""

# -*- coding: utf-8 -*-
import pandas as pd
import glob
import csv



def ler_xlsx(file_name,linha,coluna):       
    file = pd.read_excel(file_name, 'Page 1', index_col=None, usecols = coluna, header =linha, nrows=0)
    data = file.columns.values[0]
    return data

#faturamento produto = total item
def escrever_csv(writer, atendente, qtde, fat_prod, total_com, dia, item):       
    writer.writerow({'Atendente':atendente, 'Qtde':qtde, 'Faturamento Produto':fat_prod,
                     'Total Comissao': total_com, 'Data':dia, 'Item':item})
    
def executar():
    xlsxfiles = []
    files = glob.glob("tipo2/*.xlsx")
    print("Total de Arquivos: ",len(files))
    
    for file in files:
        xlsxfiles.append(file)

        file_name = xlsxfiles.pop(0); 
    
        #Obtendo o tipo de posto
        f = file_name.replace('tipo2\\','')
        f = f.replace('.xlsx','')
        s = list(f)
        if(s[0] == "0"):
            tipo_posto = s[1]
        else:
            tipo_posto = s[0] + s[1]
        
        #colunas a usar
        coluna = ["A", "B", "C", "I"]
        
        
        #Obtendo a data
        linha = 1    
        dia = ler_xlsx(file_name,linha,coluna[0])
        dia = dia.split(" ")
        
        if(dia[1] == dia[3]):
            dia = dia[1];
        else:
            dia = dia[1] + " a " + dia[3];
      
        #Obtendo o restante dos dados
        stop = 0;
        dados = []    
        linha = 5; 
        d = ""
        
        while (True):
            for c in coluna:
                d = ler_xlsx(file_name, linha, c)
                
                if(("Total da Empresa" in str(d)) == True):
                    stop = 1;
                    break;
                else:    
                    dados.append(d)
         
            if(stop == 1):
                break;
            else:    
                #print(dados)    
                escrever_csv(writer, dados[0],dados[1],dados[2],dados[3],dia, tipo_posto)
                dados = []
                linha = linha + 1;    
        
        
    
with open('posto1_func.csv', mode='w', newline = '') as csv_file:
    fieldnames = ['Atendente', 'Qtde', 'Faturamento Produto', 'Total Comissao', 'Data', 'Item']
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
    writer.writeheader()
    executar();
    
      
#obs: total liq = total item + acrescimo + desconto

