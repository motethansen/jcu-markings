#!/usr/bin/python

import sys, getopt
from openpyxl import Workbook
from openpyxl import load_workbook

def markCleaner(inputfile, outputfile):
    workbook = load_workbook(filename=inputfile)
    outputfilename = outputfile

    sheet = workbook.active

    for row in sheet.rows:
        for cell in row:
            value = cell.value
            if (type(value) == str):
                if (value.find('Needs Grading') !=-1):
                    start = value.find("Needs Grading(") + len("Needs Grading(")
                    end = value.find(")")
                    substring = value[start:end]
                    cell.value = substring
                if (value.find('In Progress') !=-1):
                    start = value.find("In Progress(") + len("In Progress(")
                    end = value.find(")")
                    substring = value[start:end]
                    cell.value = substring
                if (value.find('Reply to Post') !=-1):
                    start = value.find("Reply to Post(") + len("Reply to Post(")
                    end = value.find(")")
                    substring = value[start:end]
                    cell.value = substring
        #row = [float(cell) if cell.isdigit() else cell for cell in row]
    workbook.save(filename=outputfilename)

def main(argv):
   inputfile = ''
   outputfile = 'updatedmarksheets.xlsx'
   try:
      opts, args = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
   except getopt.GetoptError:
      print ('clean-marksheet.py -i <inputfile> -o <outputfile>')
      sys.exit(2)
   for opt, arg in opts:
      if opt == '-h':
         print ('clean-marksheet.py -i <inputfile> -o <outputfile>')
         sys.exit()
      elif opt in ("-i", "--ifile"):
         inputfile = arg
      elif opt in ("-o", "--ofile"):
         outputfile = arg
   print ('Input file is : ', inputfile)
   print ('Output file is : ', outputfile)
   if (inputfile!=''):
       markCleaner(inputfile, outputfile)
   else:
       print('Please set the input file clean-marksheet.py -i <inputfile> -o <outputfile>')

if __name__ == "__main__":
   main(sys.argv[1:])