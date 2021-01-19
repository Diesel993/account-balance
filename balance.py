import os
import numpy as np
#import path
import xml.etree.ElementTree as ET
import linecache 
import re
import xlsxwriter #Download at http://xlsxwriter.readthedocs.io/getting_started.html

##Path direction
# test
year = '2017'
months = ['Maanden', 'Januari', 'Februari', 'Maart', 'April', 'Mei', 'Juni', 'Juli', 'Augustus', 'September', 'Oktober', 'November', 'December']
# dr = r'I:\Uitgaven\Gezamenlijke rekening\201710_Rabo.txt'
dr = 'CSV_A_20180130_221733.csv'

# os.remove(r'\Uitgaven_GR_'+year+'.xlsx')
path_workbook =
workbook =  xlsxwriter.Workbook(r'\Uitgaven_GR_'+year+'_1.xlsx')
worksheetJan = workbook.add_worksheet('Januari')
worksheetFeb = workbook.add_worksheet('Februari')
worksheetMar = workbook.add_worksheet('Maart')
worksheetApr = workbook.add_worksheet('April')
worksheetMei = workbook.add_worksheet('Mei')
worksheetJun = workbook.add_worksheet('Juni')
worksheetJul = workbook.add_worksheet('Juli')
worksheetAug = workbook.add_worksheet('Augustus')
worksheetSep = workbook.add_worksheet('September')
worksheetOkt = workbook.add_worksheet('Oktober')
worksheetNov = workbook.add_worksheet('November')
worksheetDec = workbook.add_worksheet('December')

m_static = 0
l = []
worksheets = ['worksheets', worksheetJan, worksheetFeb, worksheetMar, worksheetApr, worksheetMei, worksheetJun, worksheetJul, worksheetAug, worksheetSep, worksheetOkt, worksheetNov, worksheetDec]
with open(dr, 'r') as file:
        for line in file:
            l.append(line.split('\n'))
            date = [int(s) for s in re.findall(r'\d{8}', line)]
            inout = [str(b) for b in re.findall(r'\w+', line)]
            m_temp = str(date[1])
            if m_temp[4] == 0:
                m = int(m_temp[5])
            else:
                m = int(m_temp[4:6])
            month = months[m]
            i = 4
            if m != m_static:
                m_static = m
                iB = 9
                iV = 9
                iO = 9
                iIn = 9
            
            
##Create Excel document


            format1 = workbook.add_format()
            format2 = workbook.add_format({'num_format':'€  #,##0.00'})    
            format3 = workbook.add_format()
            formatGreen = workbook.add_format({'bg_color':   '#C6EFCE', 'font_color': '#006100', 'num_format':'€  #,##0.00'} )
            formatRed = workbook.add_format({'bg_color':   '#FFC7CE','font_color': '#9C0006', 'num_format':'€  #,##0.00'})
            
            format1.set_bold() 
            format1.set_font_size(font_size = 18)
            format1.set_font_color('#005c99')
            
            format3.set_bold()
            
            worksheets[m].set_column('A:Z', 20)
            format1.set_bottom() 
            format1.set_bottom_color('#005c99') 
            
            worksheets[m].conditional_format('B5', {'type':     'cell', 'criteria': '<','value':    0, 'format':   formatRed})
            worksheets[m].conditional_format('B5', {'type':     'cell','criteria': '>=','value':    0,'format':   formatGreen})
           
            #Declare headers
            worksheets[m].write('A1', 'Uitgaven', format1)
            worksheets[m].write('B1', '', format1)
            worksheets[m].write('C1', month, format1)
            worksheets[m].write('D1', year, format1)
            worksheets[m].write('E1', '', format1)
            worksheets[m].write('F1', '', format1)
            worksheets[m].write('G1', '', format1)
            worksheets[m].write('H1', '', format1)
            
            worksheets[m].write('A8', 'Boodschappen', format3)
            worksheets[m].write_formula('B7', "=SUM(B9:B999)", format2)
            worksheets[m].write('D8', 'Overig', format3)
            worksheets[m].write_formula('E7', '=SUM(E9:E999)', format2)
            worksheets[m].write('G8', 'Vaste lasten', format3)
            worksheets[m].write_formula('H7', "=SUM(H9:H999)",format2)
            worksheets[m].write('J8', 'Inkomsten', format3)
            worksheets[m].write_formula('K7', "=SUM(K9:K999)", format2)
            
            worksheets[m].write('A7', 'Totaal (€)', format3)
            worksheets[m].write('A3', 'Totale uitgaven (€)', format3)
            worksheets[m].write_formula('B3', '=SUM(B7,E7,H7)', formatRed)
            worksheets[m].write('A4', 'Totale inkomsten (€)', format3)
            worksheets[m].write_formula('B4', "=K7", formatGreen)
            worksheets[m].write('A5', 'Geld beschikbaar (€)', format3)
            worksheets[m].write_formula('B5', '=B4-B3', format2) 
            
            cost = [float(x) for x in re.findall(r'\d+\.\d*', line)]
            shop = [str(y) for y in re.findall(r'"(.*?)"', line)]
            person = [str(y) for y in re.findall(r'"(.*?)"', line)]
            
            if inout[3] == 'C': #Inkomsten
                worksheets[m].write('J%i' %iIn, person[6])
                worksheets[m].write('K%i' %iIn, cost[0], format2)
                worksheets[m].write('L%i' %iIn, person[10])
                iIn += 1
                
            #Boodschappen
            elif 'Albert Heijn' in line or 'Lidl' in line or 'AH' in line or 'Jumbo' in line or 'COOP' in line or 'Spar ' in line:
                worksheets[m].write('A%i' %iB, shop[10])
                worksheets[m].write('B%i' %iB, cost[0], format2)
                
                iB += 1
            #Vaste lasten
            elif 'INTERPOLIS' in line or 'Woningstichting'in line or 'Autoverzekering' in line or 'Belastingkantoor' in line or 'Rabo' in line or 'KPN' in line or 'VITENS' in line or 'Nuon' in line or 'Sparen' in line or 'BELASTINGDIENST' in line:
                worksheets[m].write('G%i' %iV, person[6])
                worksheets[m].write('H%i' %iV, cost[0], format2)
                iV += 1
            #Overig
            else:
                worksheets[m].write('D%i' %iO, shop[10])
                worksheets[m].write('E%i' %iO, cost[0], format2)
                iO += 1
        
## Jaaropgave
        
worksheetJO = workbook.add_worksheet('Jaaropgave')
worksheetJO.set_column('A:Z', 20)

#Declare headers
worksheetJO.write('A1', 'Jaaropgave', format1)

worksheetJO.write('A3', 'Maand', format3)
worksheetJO.write('B3', 'Uitgaven', format3)
worksheetJO.write('C3', 'Inkomsten', format3)

worksheetJO.write('A18', 'Totaal', format3)
worksheetJO.write_formula('B18', '=SUM(B4:B15)', format2)
worksheetJO.write_formula('C18', '=SUM(C4:C15)', format2)
worksheetJO.write('A20', 'Balans', format3)
worksheetJO.write_formula('B20', '=C18-B18', format2)

k = 4
mon = 1
for i in range(len(months)-1):
    worksheetJO.write('A%i' %k, months[mon])
    k += 1
    mon += 1

#Get values from other sheets
worksheetJO.write_formula('B4', '=Januari!B3', format2)
worksheetJO.write_formula('C4', '=Januari!B4', format2)
worksheetJO.write_formula('B5', '=Februari!B3', format2)
worksheetJO.write_formula('C5', '=Februari!B4', format2)
worksheetJO.write_formula('B6', '=Maart!B3', format2)
worksheetJO.write_formula('C6', '=Maart!B4', format2)
worksheetJO.write_formula('B7', '=April!B3', format2)
worksheetJO.write_formula('C7', '=April!B4', format2)
worksheetJO.write_formula('B8', '=Mei!B3', format2)
worksheetJO.write_formula('C8', '=Mei!B4', format2)
worksheetJO.write_formula('B9', '=Juni!B3', format2)
worksheetJO.write_formula('C9', '=Juni!B4', format2)
worksheetJO.write_formula('B10', '=Juli!B3', format2)
worksheetJO.write_formula('C10', '=Juli!B4', format2)
worksheetJO.write_formula('B11', '=Augustus!B3', format2)
worksheetJO.write_formula('C11', '=Augustus!B4', format2)
worksheetJO.write_formula('B12', '=September!B3', format2)
worksheetJO.write_formula('C12', '=September!B4', format2)
worksheetJO.write_formula('B13', '=Oktober!B3', format2)
worksheetJO.write_formula('C13', '=Oktober!B4', format2)
worksheetJO.write_formula('B14', '=November!B3', format2)
worksheetJO.write_formula('C14', '=November!B4', format2)
worksheetJO.write_formula('B15', '=December!B3', format2)
worksheetJO.write_formula('C15', '=December!B4', format2)

workbook.close()