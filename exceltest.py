import xlrd

def loopSheets():	
	book = xlrd.open_workbook("/Users/Sam/Documents/color.xls")
	#print "The number of worksheets is", book.nsheets
	#print "Worksheet name(s):", book.sheet_names()

	colorMatchDic = {}

	for s in xrange(0, book.nsheets):

		key = book.sheet_names()[s]

		table = book.sheets()[s]
		nrows = table.nrows
		ncols = 3   #table.ncols

		oneSheetList = []

		for x in xrange(0, ncols):

			oneColList = []
			
			for y in xrange(1, nrows):
				#colorInt = int(table.cell(y, x).value)

				colorStr = str(table.cell(y, x).value).strip()
				if colorStr != "":
					#print int(table.cell(y, x).value)
					oneColList.append(int(table.cell(y, x).value))
			#print "-----------------------------"
			oneSheetList.append(oneColList)
		#print (oneSheetList)

		colorMatchDic[key] = oneSheetList

	#print(colorMatchDic)
	return colorMatchDic

def writeFile(dataDic):
	print "write file"

	header = '''\
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>'''

	end = '\n</dict>\n</plist>'

	f = file('/Users/Sam/Documents/color_matches.plist', 'w')
	f.write(header)

	# out put data
	for name, colorArray in dataDic.items():
		#print 'Color:', name, ' data:', colorArray

		theKey = '\n\t<key>'+name+'</key>'
		f.write(theKey)

		f.write('\n\t<array>')

		for item in colorArray:
			f.write('\n\t\t<array>')
			for colorCode in item:
				f.write('\n\t\t\t<string>'+str(colorCode)+'</string>')
			f.write('\n\t\t</array>')

		f.write('\n\t</array>')

	f.write(end)
	f.close()



def main():
	print "main here"

	dic = loopSheets()
	#print(dic)
	writeFile(dic)

if __name__=="__main__":
	main()
