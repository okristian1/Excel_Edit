from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook # http://pypi.python.org/pypi/xlrd
from xlwt import easyxf # http://pypi.python.org/pypi/xlwt
import xlwt

font_size_style = xlwt.easyxf('font: name Calibri, bold on, height 280;')

rb = open_workbook('faktura.xls', formatting_info=True)
r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)

sheets = rb.nsheets

for each in range(sheets):
	w_sheet = wb.get_sheet(each)
	w_sheet.write(4,3, label = 'SpareBank1 Bygget Trondheim AS', style = font_size_style)
	
	
	
wb.save('faktura1.xls')

