from openpyxl import Workbook
from openpyxl import load_workbook
import time

#Open the workbook, get the sheet in question
wb = load_workbook('testKojo_Atlantic Engineering CC Budgets for 2012_2017_JunQtr Update_9 16 13_Aug Actuals updated.xlsx')
ws = wb.get_sheet_by_name('57260')

for c in ws.columns[5]:
    if c.value > 1000:
        print c
        
    
for c in ws.columns[5]:
    print c, type(c.value)

for c in ws.columns[5]:
    if type(c.value) == int:
        print ws.title, c, c.value
'''
In [45]: type(ws.title)
Out[45]: str
'''
#Since I couldn't process ALL the sheets at once (Memory Error), I broke it into slices
#I have to clear the memory between each run. Now wouldn't be a bad time to better
#understand Profiler. :-/
sheets = wb.get_sheet_names()[1:8] #Skipping first sheet, which is a summary
sheets = wb.get_sheet_names()[8:9]
sheets = wb.get_sheet_names()[9:]

def fc_list(sheets):
	'''
	list --> dictionary

	Takes in a list of Cost Centers (strings) from a spreadsheet, goes to each spreadsheet
	and creates a list of the cells in column 5 (usally F, the feature code column) with
	a cell value of type int; These are the Feature code numbers that are 1000 or greater
	'''
	
	time0 = time.time()
	fc ={}
	for s in sheets:
		ws = wb.get_sheet_by_name(s)
		cells = []
		for cell in ws.columns[5]:
			if type(cell.value) == int:
				cells.append(cell)
		fc.update({s: cells})
	time1 = time.time()
	print "Processing time:", time1 - time0, "seconds"

	print "FC has", len(fc), "Cost Centers."
	for k in fc.keys():
		print k, len(fc[k])
	return fc

##This needs to become some sort of "Exception Finder" function
for c in range(len(fc['52P02'])-1):
    #ci = fc['52P02'][0].index(c); using 'range' above means I'm working with indexes. No need for this line
    #also, as I'm using range, but comparing one index to the next, I've taken the range of (len(x)-1)
    #This avoids 'IndexError: list index out of range'

    ##defining variables I'll use for comparision; should make code easier to read & maintain later
    c_row = fc['52P02'][c].row #spreadsheet row of cell at Index C
    c_value = fc['52P02'][c].value #value IN cell at Index C
    c_row_next = fc['52P02'][c].row + 1 #row for Cell c, +1
    next_c_row = fc['52P02'][c+1].row #row for cell ONE index after c; used for comparison

    print c, c_row, c_value, c_row_next - next_c_row
    	    

for i in fc['52P02']:
    ndx = fc['52P02'].index(i)
    print ndx
print "Length is", len(fc['52P02'])