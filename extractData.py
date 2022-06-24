from openpyxl import load_workbook
 
wb = load_workbook("temp.xlsx")   # open an Excel file and return a workbook
univName = "Indiana University-Purdue University" 
courseCodeString = "CSCI"  
if univName  in wb.sheetnames:
    print(univName+' exists')
else:
    wb.create_sheet(univName)

activesheet = wb[univName]
if(activesheet["A1"]!="Course No"):
    #activesheet.append(("Course No","Course Name"))
    activesheet["A1"]="Course No"
    activesheet["B1"]="Course Name"

prevline=""
courseCode=[]
duplicateCount=0
tempcode=""
tempcode=""
with open('temp.txt',encoding="utf8") as f:
    lines = f.readlines()
    for i in range(0, len(lines)):
        line = lines[i]   
        if(courseCodeString in line[:6]):
                
                if("CSCI-N" in line[:7]):
                    tmpstr = line[5:]
                    hipposi = tmpstr.find("-")+5
                    tempcode = "CSCI-"+line[5:hipposi]
                    tempName = line[hipposi+1:]
                elif("CSCI-C" in line[:7]):
                    tmpstr = line[5:]
                    hipposi = tmpstr.find("-")+5
                    tempcode = "CSCI-"+line[5:hipposi]
                    tempName = line[hipposi+1:]
                else:
                    tempcode = line[:line.find("-")]
                    tempName = line[line.find("-")+1:]
                if(tempcode  in courseCode):
                    print("duplicate for : "+line+" total duplicate "+str(duplicateCount))
                    duplicateCount+=1
                    continue
                courseCode.append(tempcode)
                activesheet.append((tempcode,tempName))

            # break
                print(tempcode+" "+tempName)
        prevline = line
        
f.close()

wb.save('temp.xlsx')
wb.close()