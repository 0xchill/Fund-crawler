# -*- coding: utf-8 -*-
import xlsxwriter
from xlsxwriter.utility import xl_range
from xlsxwriter.utility import xl_rowcol_to_cell
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
    percent = workbook.add_format({'num_format': '0.00 %'})

    percent_left = workbook.add_format({'num_format': '0.00 %','left':True})

    # Write some data headers.
    worksheet.merge_range('A1:C1','Fond',merge_format)
    worksheet.write('A2', 'ISIN', merge_format)
    worksheet.write('B2', 'Nom', merge_format)
    worksheet.write('C2', 'VL (devise)', merge_format)
    worksheet.write('D2', 'Devise', merge_format)

    worksheet.merge_range('E1:I1','Performances',merge_format)
    worksheet.write('E2', 'Date', merge_format)
    worksheet.write('F2', '1janv', merge_format)
    worksheet.write('G2', '3ans', merge_format)
    worksheet.write('H2', '5ans', merge_format)
    worksheet.write('I2', '10ans', merge_format)

    worksheet.merge_range('J1:M1','Répartition',merge_format)
    worksheet.write('J2', 'Actions', merge_format)
    worksheet.write('K2', 'Obligations', merge_format)
    worksheet.write('L2', 'Liquidités', merge_format)
    worksheet.write('M2', 'Autres', merge_format)

    sectors = [
        'technologie',
        'industriels',
        'consommation_cyclique',
        'consommation_defensive',
        'materiaux_de_base',
        'sante',
        'services_financiers',
        'services_de_communication',
        'services_publics',
        'immobilier',
        'energie',
        'autres'
    ]

    geo =[
        'eurozone',
        'royaume_uni',
        'europe_sauf_euro',
        'etats_unis',
        'amerique_latine',
        'japon',
        'asie_emergente',
        'asie_pays_developpes',
        'canada'
    ]

    worksheet.merge_range('N1:V1','Régions',merge_format)

    for c in range(13,22):
        cell = xl_rowcol_to_cell(1, c)
        worksheet.write(cell, geo[13-c], merge_format)

    worksheet.merge_range('W1:AH1','Secteurs',merge_format)
    for c in range(22,34):
        cell = xl_rowcol_to_cell(1, c)
        worksheet.write(cell, sectors[22-c], merge_format)

    # Read Json
    row = 2
    col = 0

    for line in json_file:
        # get last item downloaded
        last_index = len(line['data'])-1
        data = line['data'][last_index]

        worksheet.write(row,col,line['isin'])
        worksheet.write(row,col+1,line['nom'],bold)
        worksheet.write(row,col+2,data['vl']['vl'])
        worksheet.write(row,col+3,data['vl']['devise'],right)

        date = datetime.strptime(data['performance']['date'], "%d/%m/%Y")
        worksheet.write_datetime(row,col+4,date,date_format)

        # Performance
        idx = 5
        for key in list(data["performance"]):
            if key != 'date':
                nb = data["performance"][key]
                if nb !="-":
                    nb = float(nb.replace(',','.'))/100
                    worksheet.write_number(row,col+idx,nb,percent)
                else:
                    worksheet.write(row,col+idx,nb)
                idx = idx + 1

        # Type d'actif
        try:
            nb = data['repartition_actifs']['actions']['total']/100
            worksheet.write_number(row,col+9,nb,percent_left)
        except (KeyError,TypeError):
            worksheet.write(row,col+9,'-',left)
        try:
            nb = data['repartition_actifs']['obligations']['total']/100
            worksheet.write_number(row,col+10,nb,percent)
        except (KeyError,TypeError):
            worksheet.write(row,col+10,'-')
        try:
            nb = data['repartition_actifs']['liquidites']['total']/100
            worksheet.write_number(row,col+11,nb,percent)
        except (KeyError,TypeError):
            worksheet.write(row,col+11,'-')
        try:
            nb = data['repartition_actifs']['autres']['total']/100
            worksheet.write_number(row,col+12,nb,percent_right)
        except (KeyError,TypeError):
            worksheet.write(row,col+12,'-',right)

        # Regions
        for i in range(1,10):
            s = geo[i-1]
            try:
                worksheet.write(row,col+12+i,data['repartition_geo'][s]/100,percent)
            except KeyError:
                worksheet.write(row,col+12+i,'-')

        #Secteurs
        for i in range(1,13):
            s = sectors[i-1]
            try:
                worksheet.write(row,col+21+i,data['repartition_secteurs'][s]/100,percent)
            except KeyError:
                worksheet.write(row,col+21+i,'-')

        row = row + 1
        col = 0


    # Set column width
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
    worksheet.set_column("K:K",15)
    worksheet.set_column("L:L",15)
    worksheet.set_column(12,35,30)
    worksheet.set_column('V:V', 30, right)

    # for c in range(0,23):
    #     worksheet.write(row,c,'',top)

    workbook.close()
