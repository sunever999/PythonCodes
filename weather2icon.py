# coding=utf8

import xlrd
import sys

def loopSheets():	
    book = xlrd.open_workbook("/Users/Sam/Downloads/w.xls")
    #print "The number of worksheets is", book.nsheets
    #print "Worksheet name(s):", book.sheet_names()

    table = book.sheets()[0]
    nrows = table.nrows
    ncols = 5

    weatherIconDic = {}

    for y in xrange(1, nrows):

        #for x in xrange(3, ncols):
        
        #print table.cell(y, 3).value+" | "+table.cell(y, 4).value[:2]
        weatherIconDic[table.cell(y, 3).value] = table.cell(y, 4).value[:2]
	
    print(weatherIconDic)
    return weatherIconDic


def writeFile(dataDic):
    print "write file"

    reload(sys)
    sys.setdefaultencoding('utf-8')

    header = '''\
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>'''

    end = '\n</dict>\n</plist>'

    f = file('/Users/Sam/Documents/weatherIcon.plist', 'w')
    f.write(header)

	# out put data
    for weather, icon in dataDic.items():
        #print 'Color:', name, ' data:', colorArray

        theKey = '\n\t<key>'+weather+'</key>'
        f.write(theKey)

        f.write('\n\t<string>'+icon+'</string>')

    f.write(end)
    f.close()



def main():
    print "main here"

    dic = loopSheets()
    writeFile(dic)


if __name__=="__main__":
	main()
