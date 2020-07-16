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

f=open("test_names.txt","w+")
f.write("db.ChitsTest.insertMany([")
f=open("test_names.txt","a+")
j=0
for i in range(0,sheet.nrows-2):
	
	f.write("{""\n")
	f.write("\""+"sno"+"\":"+str(sheet.cell_value(i+2,j))+","+"\n")
	f.write("\""+"name\":"+"\""+str(sheet.cell_value(i+2,j+1))+"\""+","+"\n")
	f.write("\""+new_mno+"\":" +"\"" +str(sheet.cell_value(i+2,j+2))+"\""+","+"\n")
	f.write("\""+middle_mno+"\":" +"\"" +str(sheet.cell_value(i+2,j+3))+"\""+","+"\n")
	f.write("\""+old_mno+"\":"+ "\""+ str(sheet.cell_value(i+2,j+4))+"\""+","+"\n")
	f.write("\""+"Balance"+"\":"+ "\""+ str(sheet.cell_value(i+2,j+5))+"\""+","+"\n")
	f.write("\""+"TOTAL:"+"\":"+"\"" +str(sheet.cell_value(i+2,j+6))+"\""+"\n")
	f.write("},"+"\n")
	j=0

f.write("]);")	





