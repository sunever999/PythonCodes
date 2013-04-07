import xlrd

book = xlrd.open_workbook("myfile.xls")
print "The number of worksheets is", book.nsheets
print "Worksheet name(s):", book.sheet_names()

for s in xrange(0, book.nsheets):

	table = book.sheets()[s]
	nrows = table.nrows
	ncols = table.ncols

	for x in xrange(0, ncols):
		for y in xrange(1, nrows):
			#colorInt = int(table.cell(y, x).value)
			colorStr = str(table.cell(y, x).value).strip()
			if colorStr != "":
				print int(table.cell(y, x).value)
		print "-----------------------------"