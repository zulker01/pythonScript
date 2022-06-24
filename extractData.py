from openpyxl import load_workbook
 
wb = load_workbook("temp.xlsx")   # open an Excel file and return a workbook
univName = "Howard University" 
courseCodeString = "CS"  
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
with open('temp.txt',encoding="utf8") as f:
    lines = f.readlines()
    for i in range(0, len(lines)):
        line = lines[i]   
        if(line[line.find("-")+1:line.find("-")+4].isnumeric()):
                tempcodestr = line[:line.find("-")]+" "+line[line.find("-")+1:line.find("-")+4]
                tempcode = tempcodestr
                tempName = lines[i+1]

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