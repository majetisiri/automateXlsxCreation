import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles import Font
import string

def loadXlsxFile():
	wb = openpyxl.load_workbook('specs1.xlsx')
	return wb

def getSpecsSheet(wb):
	# print wb.get_sheet_names()
	sheet=wb['specs1']
	return sheet

def deleteOutputSheetIfExists(wb):
	if 'output' in wb.sheetnames:
		std=wb['output']
		wb.remove(std)
	
def createOutputSheet(wb):
	outputSheet=wb.create_sheet(title='output')
	sheet=wb['output']
	return sheet

def copyColsFromSheets(wb,specsSheet,outputSheet):
	rows=specsSheet.max_row
	columns=specsSheet.max_column
	listab=[]
	for i in range(1, rows+1):
		listab.append([])
	for r in range(1,rows+1):
		for c in range(1,columns+1):
			e=specsSheet.cell(row=r,column=c)
			listab[r-1].append(e.value)
	# print listab	
	for r in range(1,rows+1):
		for c in range(1,columns+1):
			j=outputSheet.cell(row=r,column=c)
			if c ==1:
				j.value=listab[r-1][3]
			elif c==2:
				j.value=listab[r-1][5]
			elif c==3:
				new_str=listab[r-1][5].replace(" ","")
				j.value=new_str
			elif c==4:
				j.value=listab[r-1][4]
	wb.save('specs1.xlsx')

def changeBackgroundColor(wb):
	# Change background color 
	sheet=wb['output']
	for cell in sheet["1:1"]:
		cell.fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
		cell.font = Font(color="FFFFFFFF",italic=True,name="Arial")
	wb.save('specs1.xlsx')


wb=loadXlsxFile()
specsSheet=getSpecsSheet(wb)
deleteOutputSheetIfExists(wb)
outputSheet=createOutputSheet(wb)
wb.save('specs1.xlsx')
copyColsFromSheets(wb,specsSheet,outputSheet)
changeBackgroundColor(wb)