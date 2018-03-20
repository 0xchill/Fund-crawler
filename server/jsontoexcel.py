# -*- coding: utf-8 -*-
import xlsxwriter
from xlsxwriter.utility import xl_range
import json
from datetime import datetime

def create_excel(json_file):
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('../static/Fonds.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.hide_gridlines(2)

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Add a number format for cells with money.
    percent = workbook.add_format({'num_format': '0.00 %'})

    # Add an Excel date format.
    date_format = workbook.add_format({'num_format': 'd-m-yyyy'})


    # Create a format to use in the merged range.
    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})

    right = workbook.add_format({
        'right': True,
        })

    left = workbook.add_format({
        'left': True,
        })

    top = workbook.add_format({
        'top': True,
        })

    percent_right = workbook.add_format({'num_format': '0.00 %', 'right':True})

    percent_left = workbook.add_format({'num_format': '0.00 %','left':True})

    # Write some data headers.
    worksheet.merge_range('A1:C1','Fond',merge_format)
    worksheet.write('A2', 'ISIN', merge_format)
    worksheet.write('B2', 'Nom', merge_format)
    worksheet.write('C2', 'VL (m €)', merge_format)

    worksheet.merge_range('D1:H1','Performances',merge_format)
    worksheet.write('D2', 'Date', merge_format)
    worksheet.write('E2', '1janv', merge_format)
    worksheet.write('F2', '3ans', merge_format)
    worksheet.write('G2', '5ans', merge_format)
    worksheet.write('H2', '10ans', merge_format)

    worksheet.merge_range('I1:L1','Répartition',merge_format)
    worksheet.write('I2', 'Actions', merge_format)
    worksheet.write('J2', 'Obligations', merge_format)
    worksheet.write('K2', 'Liquidités', merge_format)
    worksheet.write('L2', 'Autres', merge_format)

    worksheet.merge_range('M1:Q1','Régions',merge_format)
    worksheet.write('M2', 'Top1', merge_format)
    worksheet.write('N2', 'Top2', merge_format)
    worksheet.write('O2', 'Top3', merge_format)
    worksheet.write('P2', 'Top4', merge_format)
    worksheet.write('Q2', 'Top5', merge_format)

    worksheet.merge_range('R1:V1','Secteurs',merge_format)
    worksheet.write('R2', 'Top1', merge_format)
    worksheet.write('S2', 'Top2', merge_format)
    worksheet.write('T2', 'Top3', merge_format)
    worksheet.write('U2', 'Top4', merge_format)
    worksheet.write('V2', 'Top5', merge_format)

    # Read Json
    row = 2
    col = 0

    for line in json_file:
        worksheet.write(row,col,line['isin'])
        worksheet.write(row,col+1,line['nom'],bold)
        worksheet.write(row,col+2,line['vl'],right)

        date = datetime.strptime(line['performance']['date'], "%d/%m/%Y")
        worksheet.write_datetime(row,col+3,date,date_format)

        idx = 4
        for key in list(line["performance"]):
            if key != 'date':
                nb = line["performance"][key]
                if nb !="-":
                    nb = float(nb.replace(',','.'))/100
                    worksheet.write_number(row,col+idx,nb,percent)
                else:
                    worksheet.write(row,col+idx,nb)
                idx = idx + 1
        try:
            nb = float(line['actions']['total'].replace(",","."))/100
            worksheet.write_number(row,col+8,nb,percent_left)
        except (KeyError,TypeError):
            worksheet.write(row,col+8,'-',left)
        try:
            nb = float(line['obligations']['total'].replace(',','.'))/100
            worksheet.write_number(row,col+9,nb,percent)
        except (KeyError,TypeError):
            worksheet.write(row,col+9,'-')
        try:
            nb = float(line['liquidites']['total'].replace(',','.'))/100
            worksheet.write_number(row,col+10,nb,percent)
        except (KeyError,TypeError):
            worksheet.write(row,col+10,'-')
        try:
            nb = float(line['autres']['total'].replace(',','.'))/100
            worksheet.write_number(row,col+11,nb,percent_right)
        except (KeyError,TypeError):
            worksheet.write(row,col+11,'-',right)

        keys=['first','second','third','fourth','fifth']
        for i in range(1,6):
            s = keys[i-1]
            worksheet.write(row,col+11+i,line['regions'][s])

        worksheet.write(row,col+17,line['secteurs'][keys[0]],left)
        for i in range(2,5):
            s = keys[i-1]
            worksheet.write(row,col+16+i,line['secteurs'][s])
            worksheet.write(row,col+21,line['secteurs'][keys[4]],right)

        row = row + 1
        col = 0

    worksheet.set_column("A:A",15)
    worksheet.set_column("B:B",84)
    worksheet.set_column("C:C",13)
    worksheet.set_column("D:D",10)
    worksheet.set_column("E:E",7)
    worksheet.set_column("F:F",7)
    worksheet.set_column("G:G",7)
    worksheet.set_column("H:H",7)
    worksheet.set_column("I:I",8)
    worksheet.set_column("J:J",11)
    worksheet.set_column("K:K",9)
    worksheet.set_column("L:L",8)
    worksheet.set_column("M:M",30)
    worksheet.set_column("N:N",30)
    worksheet.set_column("O:O",30)
    worksheet.set_column("P:P",29)
    worksheet.set_column("Q:Q",29)
    worksheet.set_column("R:R",34)
    worksheet.set_column("S:S",34)
    worksheet.set_column("T:T",34)
    worksheet.set_column("U:U",34)
    worksheet.set_column("V:V",34)

    for c in range(0,23):
        worksheet.write(row,c,'',top)

    workbook.close()
