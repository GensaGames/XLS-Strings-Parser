#!/usr/bin/python



import sys, getopt
import os.path
import xlrd

# -------------------------------------------------
# --------------- Main Consts here ---------------
# -------------------------------------------------
CONST_RESULT_FILE = "string.xml"
CONST_FILE_HEADER = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
CONST_STRING_RES ="<string name=\"{key}\">{val}</string>\n"
CONST_RESOURCE_START="<resources>\n"
CONST_RESOURCE_END="</resources>\n"

def main(argv):
   values = ['', -1, -1, -1]
   try:
      opts, args = getopt.getopt(argv, 'he:f:s:k:v:', \
      	['help', 'file=', 'sheet=', 'sheet_key=', 'sheet_val='])
   except getopt.GetoptError:
      printError()
      sys.exit(2)
   for opt, arg in opts:
      if opt in ('-h', '--help'):
         printHelp()
         sys.exit()
      elif opt in ('-f','--file'):
         values[0] = arg
      elif opt in ('-s','--sheet'):
         values[1] = arg
      elif opt in ('-k','--sheet_key'):
         values[2] = arg
      elif opt in ('-v','--sheet_val'):
         values[3] = arg

   if not os.path.isfile(values[0]) \
   or values[1] < 0 or values[2] < 0 or values[3] < 0:
      printError()
      sys.exit(2)

   printValues(values)
   resolve(values)


def resolve(values):
	workbook = xlrd.open_workbook(values[0], formatting_info=True)
	worksheet = workbook.sheet_by_index(int(values[1]))
	dict_list = dict()
	for index in range(0, worksheet.nrows):
		dict_list[worksheet.cell(index, int(values[2])).value] =\
		worksheet.cell(index, int(values[3])).value

	printFile(dict_list)


def printFile(dict_list):
	file = open(CONST_RESULT_FILE, 'w+')
	file.write(CONST_FILE_HEADER)
	file.write(CONST_RESOURCE_START)
	for key, value in dict_list.iteritems():
		valuesString = CONST_STRING_RES.format(key = key.encode('ascii','ignore'), \
                                                       val = value.encode('ascii','ignore'))
		file.write(valuesString)

	file.write(CONST_RESOURCE_END)
	file.close()


# -------------------------------------------------
# --------------- Text Messages -------------------
# -------------------------------------------------
def printValues(values):
	print ''
	print ' Proceed with next values: '
	print ' -------------------------------- '
	print ' File Location: ', values[0]
	print ' Sheet Index ', values[1]
	print ' Index String Key: ', values[2]
	print ' Index String  Value:', values[3]
	print ' -------------------------------- '
	print ' .... '
	print ''


def printError():
	print """
	[ERROR] Exceptions during resolving objects! 
	Write --help for more information.
	"""


def printHelp():
	print """
	To read and write XLS files, you would need to install two helpers.
	xlrd and xlwt make sure you are have that modules
	-------------------------------------------------------------
	"""
	print """
	Scripts requires input XLS file and columns index for sheets.
	Correct order for script: 
	--file <path/to/file.xls>	
	--sheet <sheet index>
	--sheet_key <key sheet index>
	--sheet_val <value sheet index>
	"""


if __name__ == "__main__":
   main(sys.argv[1:])
