# -*- coding: utf-8 -*-
"""
Created on Sat Aug 15 19:05:23 2020

@author: User-PC
"""

import xlsxwriter
import xlrd
import glob

def escrever_linha_xlsx(new_worksheet,n,medias):
    
    i = 0
    for dado in medias:
        new_worksheet.write(n,i,dado)
        i = i + 1

def medias_da_empresa(files, empresa):
    qtd_meses = 11
    mvb = mvl = 0
    
    for file_name in files:
        workbook = xlrd.open_workbook(file_name)
        sheet = workbook.sheet_by_index(0)
        L = 0
        while(L != sheet.nrows):
            if(sheet.cell(L,1).value == empresa):
                mvb = mvb + sheet.cell(L,2).value
                mvl = mvl + sheet.cell(L,3).value
 #               print(empresa,mvb,mvl)
                break;
             
            L = L + 1   
    #print('\n')   
    dados = ['06-2019 a 05-2020', empresa, mvb/qtd_meses,mvl/qtd_meses]
    return dados
    

workbook = xlsxwriter.Workbook('medias_de_vendas_totais_empresa.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold':True})

#headers
worksheet.write('A1','Periodo',bold)
worksheet.write('B1','Empresa',bold)
worksheet.write('C1','Méd. Venda Bruta Total',bold)
worksheet.write('D1','Méd. Venda Líquida Total',bold)

#arquivos de fonte
files = glob.glob("xt2/*.xlsx")

empresas = ['POSTO FORMOSA', 'POSTO GRUTA DA LAPA', 'POSTO IBOTIRAMA', 'POSTO JAJA', 'POSTO MACAUBENSE II - BARREIRAS',
          'POSTO MACAUBENSE III -MG', 'POSTO MACAUBENSE I - MACAUBAS', 'POSTO MACAUBENSE IV - RAFAEL JAMBEIRO',
          'POSTO MACAUBENSE VI - VIT. CONQUISTA',  'POSTO RODA VELHA', 'POSTO SABRINA II - BRUMADO',
          'POSTO SABRINA III - VIT CONQUISTA', 'POSTO SABRINA I - LIVRAMENTO', 'POSTO SABRINA IV - BARREIRAS',
          'POSTO SABRINA IX - CAETI','POSTO SABRINA VI - GUANAMBI', 'POSTO SABRINA VII - GUANAMBI',
          'POSTO SABRINA VIII - VIT CONQUISTA', 'POSTO SABRINA V - PARAMIRIM', 'POSTO SABRINA XII - BARREIRAS', 'POSTO SABRINA X (NOVO)', 'POSTO SIGA BEM' ]

print(len(empresas))
#categorias = ['ADITIVOS', 'FILTROS', 'LUBRIFICANTES', 'OMBUSTIVEIS', 'DIVERSOS']
        
n = 1;
for empresa in empresas:
    medias = medias_da_empresa(files, empresa)
    escrever_linha_xlsx(worksheet,n, medias)     
    n = n + 1;



workbook.close();