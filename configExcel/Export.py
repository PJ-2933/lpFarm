#-*- coding=utf-8 -*-

import os, os.path
import re, xlrd, math, glob, shutil, sys, csv

# reload(sys)
# sys.setdefaultencoding( "utf-8" )

######
# ***_P is path
#
######

ProductRoot_P = "D:/Project/LP/lpFarm/configExcel/"
ClientRoot_P = "D:/Project/LP/lpFarm/Assets/GameMain/"

## ***** test*****
FinalClientPath_Table = ClientRoot_P + "DataTables"
#FinalClientPath_Note = ClientRoot_P + "Notes"

#FinalClientPath_Table = ProductRoot_P + "Tables"
#FinalClientPath_Note = ProductRoot_P + "Notes"


FinalServerPath = ProductRoot_P + "Game/Xml/"

TABLE_P = ProductRoot_P + "excel/"
#NOTE_P = ProductRoot_P + "Note/"

TABLE_EXPORT = ProductRoot_P + "Export/"
#NOTE_EXPORT = ProductRoot_P + "Export/Note/"

XML_P = ProductRoot_P + "Export/Xml/"

#ClientTableExport_P = FinalClientPath
#ServerTableExport_P = FinalServerPath

def MakeDir(path):
	if not os.path.exists(path):
		os.makedirs(path)

def DelDir(path):
	if os.path.exists(path):
		shutil.rmtree(path)

def ReSetupDir(path):
	DelDir(path)
	MakeDir(path)

def CopyFiles(src_dir, dest_dir, filepattern):
	files = glob.iglob(os.path.join(src_dir, filepattern))
	for file in files:
		if os.path.isfile(file):
			#print("copy file: " + file + " to " + dest_dir)
			shutil.copy2(file, dest_dir)

def IteratorFiles(re_exp, sourceDir, exportDir, funcOnFile):
	for f in os.listdir(sourceDir):
		sourceF = os.path.join(sourceDir, f)

		if os.path.isfile(sourceF):
			if not re.match(re_exp, sourceF):
				continue

			if f.startswith('.'):
				continue
				
			if f.startswith('~'):
				continue

			print("file: " + f)
			if funcOnFile:
				funcOnFile(f, exportDir, sourceF)

		if os.path.isdir(sourceF):
			IteratorFiles(re_exp, sourceF, exportDir, funcOnFile)

def isString(value):
		if isNumber(value):
				return False
		return type(value) is str or type(value) is unicode

def isNumber(value):
		try:
				a = float(value)
		except:
				return False;
		return True

def is_chinese(uchar):  
		if uchar >= u'\u4e00' and uchar<=u'\u9fa5':  
				return True  
		else:  
				return False  

def isfloat(value):
		if isint(value):
				return False;
		try:
				a = float(value)
		except:
				return False;
		return True;

def isint(value):
		try:
				f = float(value);
				if math.floor(f) == f:
						return True;
				else:
						return False;
		except:
				return False;
		return True;

def overwriteFile(filePath, content):
	file_ = open(filePath, 'w')
	file_.write(content)
	file_.close()

def open_excel(filePath):
	try:
		data = xlrd.open_workbook(filePath, on_demand=True)
		return data
	except Exception, e:
		print str(e)

def excel_table_byindex(book, by_index=0):
	table = book.sheets()[by_index]
	nrows = table.nrows
	# ncols = table.ncols
	list = []
	for rownum in xrange(0,nrows):
		row = table.row_values(rownum)
		list.append(row)
	return list

def paser_table(table):
	retString = ''
	for r, row in enumerate(table):
		for i, cell in enumerate(row):
			# 2th cell is note, dont export
			if i == 1:
				continue

			try:
				if isString(cell):
					retString += cell.encode('utf-8')
				elif isint(cell):
						retString += str(int(cell))
				elif isfloat(cell):
					retString += str(float(cell))
				else:
					retString += str(cell)

				if i < len(row):
					retString += '\t'

			except Exception, e:
				print 'Err: %s in Row %d Clum %d' % (e, r+1, i+1)
				os.system("pause") 

		retString += '\n'

	return retString

def ExportTable(f, exportDir, sourceF):
	book = open_excel(sourceF)

	for i in range(book.nsheets):
		sheet = book.sheets()[i]
		tableName = sheet.name
		print "Exporting " + tableName

		table = excel_table_byindex(book, i)
		content = paser_table(table)

		overwriteFile(exportDir + tableName + '.txt', content)


def convert_csv_to_xml(f, sourceF):
	reader = csv.reader(open(sourceF, 'r'), delimiter='\t')
	print "Converting " + f

	content = "<TABLE>\n" ####### Start Table

	for row in reader:
		if reader.line_num == 1:
			continue

		# XML child elements
		if reader.line_num == 2:
			header = row
		else:
			content += "\t<DATA>\n"	####### Start Data
			for i, field in enumerate(row):
				if not field:
					continue
				content += "\t\t<%s>" % header[i] + field + "</%s>\n" % header[i]
			content += "\t</DATA>\n"	####### End Data

	content += "</TABLE>\n" ####### End Table

	xmlname = os.path.basename(os.path.splitext(f)[0])
	overwriteFile(XML_P + xmlname + '.xml', content)


if __name__ == '__main__':
	ReSetupDir(TABLE_EXPORT)
	#ReSetupDir(NOTE_EXPORT)

	# ReSetupDir(XML_P)

	# convert xlsx to csv for client
	IteratorFiles(".*\.xlsx", TABLE_P, TABLE_EXPORT, ExportTable)
	CopyFiles(TABLE_EXPORT, FinalClientPath_Table, "*.txt")

	#IteratorFiles(".*\.xlsx", NOTE_P, NOTE_EXPORT, ExportTable)
	#CopyFiles(NOTE_EXPORT, FinalClientPath_Note, "*.txt")

	# conver csv to xml for server
	# IteratorFiles(".*\.txt", CSV_P, convert_csv_to_xml)
	# CopyFiles(XML_P, ServerTableExport_P, "*.xml")

	os.system("pause") 

	# ExportTable(TABLE_P + "CloseUpClip.xlsx", TABLE_P)
