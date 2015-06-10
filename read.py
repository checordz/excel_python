# -*- coding: utf-8 -*-
import xlrd

def compare():
	dataset_path = "/Users/sergio/Developer/python/excel_test/data_set.xlsx"
	workbook = xlrd.open_workbook(dataset_path)
	print workbook.sheet_names()
	worksheet = workbook.sheet_by_name('Sheet 1')
	dataset_rows = worksheet.nrows

	extract_path = "/Users/sergio/Developer/python/excel_test/extract.xlsx"
	workbook_extract = xlrd.open_workbook(extract_path)
	worksheet_extract = workbook_extract.sheet_by_name('Sheet 1')
	extract_rows = worksheet_extract.nrows
	fila = 0
	columna = 0
	if dataset_rows >= extract_rows:
		fila, columna = find_data(worksheet_extract, worksheet)
	else:
		fila, columna = find_data(worksheet, worksheet_extract)

	print("fila", fila)
	print("col", columna)
	print(worksheet.cell_xf_index(fila, columna).background.pattern_colour_index)


def find_data(origen, destino):
	print(origen.nrows, destino.nrows)
	data_origen = []
	origenrows = origen.nrows
	origencoumns = origen.ncols
	print range(origenrows), range(origencoumns)
	for row in range(origenrows):
		temp_row = []
		for cell in range(origencoumns):
			temp_row.append(origen.cell(row, cell).value)
		data_origen.append(temp_row)

	data_destino = []
	destinorows = destino.nrows
	destinocoumns = destino.ncols
	print range(destinorows), range(destinocoumns)
	for row in range(destinorows):
		temp_row = []
		for cell in range(destinocoumns):
			temp_row.append(destino.cell(row, cell).value)
		data_destino.append(temp_row)

	print data_origen, data_destino
	fila = 0
	columna = 0
	
	for indexr, row in enumerate(data_destino):
		for indexc, cell in enumerate(row):
			if data_origen[2][0] == cell:
				fila, columna = indexr, indexc

	return fila, columna
if __name__ == '__main__':
	compare()