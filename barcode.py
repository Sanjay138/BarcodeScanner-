from datetime import datetime
from xlutils.copy import copy
from xlrd import open_workbook 
from xlwt import easyxf
def getTime():
    now = datetime.now();
    return  now.strftime("%H:%M:%S")
def getDate():
    now = datetime.now();
    return  now.strftime("%Y-%m-%d")
readbook = open_workbook("barcodes.xls",formatting_info=True);
readsheet =readbook.sheet_by_index(0);
writebook = copy(readbook);
w_sheet = writebook.get_sheet(0)
while True:
	done=False
	barcode=input("Scan a barcode right now:");
	if readsheet.nrows>2:
		for r in range(1, readsheet.nrows):
			b=int(readsheet.cell(r,1).value);
			if(b==int(barcode)):#if we have existing barcode in xl then update time t2
				w_sheet.write(r, 4, getTime())
				writebook.save("barcodes.xls");
				done=True
	if done==False:
		w_sheet.write(readsheet.nrows, 0, readsheet.nrows)
		w_sheet.write(readsheet.nrows, 1, int(barcode))
		w_sheet.write(readsheet.nrows, 2, getDate())
		w_sheet.write(readsheet.nrows, 3, getTime())
		writebook.save("barcodes.xls");
	done=False
