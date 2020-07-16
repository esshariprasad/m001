import xlrd

loc = ("chits\FEB2019.xlsx")

wb=xlrd.open_workbook(loc)
sheet=wb.sheet_by_index(0)


print(sheet.cell_value(1,0))
# 1,0 means 2nd row q col

print(sheet.nrows)
print(sheet.ncols)

sno=sheet.cell_value(1,0)
name=sheet.cell_value(1,1)

new_mno=sheet.cell_value(1,2)  #new_month new one that is newer then other
middle_mno=sheet.cell_value(1,3) #middle then both
old_mno=sheet.cell_value(1,4)

f=open("newChitformat.txt","w+")
f.write("db.newChitsformat.insertMany([")
f=open("newChitformat.txt","a+")
j=0
for i in range(0,sheet.nrows-2):
	
	f.write("{""\n")
	if(sheet.cell_value(i+2,j)=="" ):
		f.write("\""+"sno"+"\":"+"\""+"NULL"+"\""+","+"\n")
	else:

		f.write("\""+"sno"+"\":"+str(sheet.cell_value(i+2,j))+","+"\n")
	f.write("\""+"name\":"+"\""+str(sheet.cell_value(i+2,j+1))+"\""+","+"\n")
	if("X2" in str(sheet.cell_value(i+2,j+2))):
		f.write("\""+"newENROLLED"+"\":"+"2"+","+"\n")
	elif("X3" in str(sheet.cell_value(i+2,j+2))):
		f.write("\""+"newENROLLED"+"\":"+"3"+","+"\n")
	elif("X4" in str(sheet.cell_value(i+2,j+2))):
		f.write("\""+"newENROLLED"+"\":"+"4"+","+"\n")
	elif("X5" in str(sheet.cell_value(i+2,j+2))):
		f.write("\""+"newENROLLED"+"\":"+"5"+","+"\n")
	elif("X6" in str(sheet.cell_value(i+2,j+2))):
		f.write("\""+"newENROLLED"+"\":"+"6"+","+"\n")			
	elif("X7" in str(sheet.cell_value(i+2,j+2))):
		f.write("\""+"newENROLLED"+"\":"+"7"+","+"\n")
	elif((sheet.cell_value(i+2,j+2)=="")):
		f.write("\""+"newENROLLED"+"\":"+"\""+"NULL"+"\""+","+"\n")
	else: 
		f.write("\""+"newENROLLED"+"\":"+"1"+","+"\n")
	
	if("X2" in str(sheet.cell_value(i+2,j+3))):
		f.write("\""+"middle_month_ENROLLED"+"\":"+"2"+","+"\n")
	elif("X3" in str(sheet.cell_value(i+2,j+3))):
		f.write("\""+"middle_month_ENROLLED"+"\":"+"3"+","+"\n")

	elif("X4" in str(sheet.cell_value(i+2,j+3))):
		f.write("\""+"middle_month_ENROLLED"+"\":"+"4"+","+"\n")
	elif("X5" in str(sheet.cell_value(i+2,j+3))):
		f.write("\""+"middle_month_ENROLLED"+"\":"+"5"+","+"\n")
	elif("X6" in str(sheet.cell_value(i+2,j+3))):
		f.write("\""+"middle_month_ENROLLED"+"\":"+"6"+","+"\n")			
	elif("X7" in str(sheet.cell_value(i+2,j+3))):
		f.write("\""+"middle_month_ENROLLED"+"\":"+"7"+","+"\n")	
	elif((sheet.cell_value(i+2,j+3)=="")):
		f.write("\""+"middle_month_ENROLLED"+"\":"+"\""+"NULL"+"\""+","+"\n")
	
	else: 
		f.write("\""+"middle_month_ENROLLED"+"\":"+"1"+","+"\n")
	
	if("X2" in str(sheet.cell_value(i+2,j+4))):
		f.write("\""+"last_month_ENROLLED"+"\":"+"2"+","+"\n")
	elif("X3" in str(sheet.cell_value(i+2,j+4))):
		f.write("\""+"last_month_ENROLLED"+"\":"+"3"+","+"\n")
	
	elif("X4" in str(sheet.cell_value(i+2,j+4))):
		f.write("\""+"last_month_ENROLLED"+"\":"+"4"+","+"\n")
	elif("X5" in str(sheet.cell_value(i+2,j+4))):
		f.write("\""+"last_month_ENROLLED"+"\":"+"5"+","+"\n")
	elif("X6" in str(sheet.cell_value(i+2,j+4))):
		f.write("\""+"last_month_ENROLLED"+"\":"+"6"+","+"\n")			
	elif("X7" in str(sheet.cell_value(i+2,j+4))):
		f.write("\""+"last_month_ENROLLED"+"\":"+"7"+","+"\n")
	elif((sheet.cell_value(i+2,j+4)=="")):
		f.write("\""+"last_month_ENROLLED"+"\":"+"\""+"NULL"+"\""+","+"\n")
		
	else: 
		f.write("\""+"last_month_ENROLLED"+"\":"+"1"","+"\n")
	
	if(sheet.cell_value(i+2,j+5)==""):
		f.write("\""+"Balance"+"\":"+"\""+"NULL"+"\""+","+"\n")	
	else:
		f.write("\""+"Balance"+"\":"+ str(sheet.cell_value(i+2,j+5))+","+"\n")
	if(sheet.cell_value(i+2,j+6)==""):
		f.write("\""+"TOTAL"+"\":"+"\""+"NULL"+"\""+"\n")
	else:
		f.write("\""+"TOTAL"+"\":"+str(sheet.cell_value(i+2,j+6))+"\n")
	f.write("},"+"\n")
	j=0

f.write("]);")	





