import openpyxl
import os


def append(infile, wsout):
	wb = openpyxl.load_workbook(infile)
	ws = wb.get_active_sheet()
	for row in ws.rows:
		rowdata = []
		for cell in row:
			rowdata.append(cell.value)
		wsout.append(rowdata)

if __name__ == '__main__':
	outfile = "out.xlsx"
	wbout = openpyxl.Workbook()
	wsout = wbout.get_active_sheet()
	for infile in os.listdir("."):
		if len(infile) > 4 and infile[-5:] == '.xlsx':
			print infile
			append(infile, wsout)

	wbout.save(outfile)

