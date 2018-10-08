#####
#Input : CAMP Data File
#Output : That same file split into chunks
#####

import argparse
from openpyxl import Workbook, load_workbook

def getLineSplits():
    """Return a list of tuples showing where the lines should break."""
    lineSplits = []

    start = 2 #presume header row
    while(start <= max_row):
        end = start + args.chunkSize
        if (end >= max_row):
            lineSplits.append((start,max_row))
            break
        lineSplits.append((start,end))
        start = end + 1

    return lineSplits

def createOutputFile(filename,startRow,endRow):
    """Creates an output file with a heading row and rows from startRow to endRow from inputFile"""
    #create output sheet
    outputBook = Workbook()
    outputSheet = outputBook.active
    outputSheet['A1'] = "JWTO"
    outputBook.save("dm-test.xlsx")
    #copy in header
    #copy in rows
    #save the output file
    

#Get the input file Name
#Get the output base
#Get the chunksize 
parser = argparse.ArgumentParser()
parser.add_argument("inputFilename")
parser.add_argument("outputFilename")
parser.add_argument("chunkSize", type=int)
args = parser.parse_args()

showSplits = True
log = True

#Read in input file
inputBook = load_workbook(args.inputFilename)
inputSheet = inputBook.active

#Get max row count
max_row = inputSheet.max_row

#get max column count
max_column = inputSheet.max_column




print(getLineSplits())

print("Creating output file")
createOutputFile("blarg",2,30)


