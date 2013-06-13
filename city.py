import xlrd
import sys


def loopSheets():	
    book = xlrd.open_workbook("/Users/Sam/Downloads/city.xls")
    #print "The number of worksheets is", book.nsheets
    #print "Worksheet name(s):", book.sheet_names()

    table = book.sheets()[1]
    nrows = table.nrows
    #ncols = 5

    #weatherIconDic = {}

    initialList = []

    for r in xrange(1, nrows):

        #for x in xrange(3, ncols):

        initialList.append(str(table.cell(r, 3).value).lower())
        
        #print table.cell(r, 1).value+" | "+table.cell(r, 2).value+" | "+table.cell(r, 3).value+" | "+table.cell(r, 4).value
        #weatherIconDic[table.cell(y, 3).value] = table.cell(y, 4).value[:2]
	
    initialSet = set(initialList)
    print(initialSet)

    # add city to array, add array to dic

    cityDic = {}
    for l in initialSet:
        cityDic[l] = []

    for r in xrange(1, nrows):
        letter = str(table.cell(r, 3).value).lower()

        oneCity = {}
        oneCity['code'] = str(table.cell(r, 1).value)
        oneCity['name'] = str(table.cell(r, 2).value)
        oneCity['letter'] = str(table.cell(r, 3).value).lower()
        oneCity['province'] = str(table.cell(r, 4).value)

        letterList = cityDic[letter]
        letterList.append(oneCity)

    #print(cityDic)
    return cityDic


def writeFile(cityDic):
    header = '''\
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>'''
    header = '''\
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>'''

    end = '\n</dict>\n</plist>'

    f = file('/Users/Sam/Documents/city.plist', 'w')
    f.write(header)

	# out put data
    for letter, citys in cityDic.items():

        theKey = '\n\t<key>'+letter+'</key>'
        f.write(theKey)

        f.write('\n\t<array>')
        for city in citys:
            f.write('\n\t\t<dict>')
            
            for ckey, cvalue in city.items():
                f.write('\n\t\t\t<key>'+ckey+'</key>')
                f.write('\n\t\t\t<string>'+cvalue+'</string>')

            f.write('\n\t\t</dict>')            

        f.write('\n\t</array>')

    f.write(end)
    f.close()


def main():
    reload(sys)
    sys.setdefaultencoding('utf-8')

    print "main here"

    dic = loopSheets()
    writeFile(dic)


if __name__=="__main__":
	main()
