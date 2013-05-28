import sys
import xlrd

# Parse Parameters (Filename in, Filename out, list of columns to use for sunburst from center to outer rings)
filein=sys.argv[1]
fileout=sys.argv[2]
sheetName=sys.argv[3]
columnsin=sys.argv[4]
rowsToSkip=int(sys.argv[5])
centreLabel=sys.argv[6]

print ('\n\n\n**************************************')
print ('Excel File   : '+filein)
print ('JSon File    : '+fileout)
print ('Work Sheet   : '+sheetName)
print ('Column List  : '+columnsin)
print ('Rows to Skip : ',rowsToSkip)
print ('Centre Label : '+centreLabel)

columns = columnsin.split(',')

#Open the file
workbook = xlrd.open_workbook(filein) # formatting_info only works on XLS files, not XLSX!!!
print('\n\nWorksheets in Excel file:')
print (workbook.sheet_names())

#Parse Excel into Data Structure
worksheet = workbook.sheet_by_name(sheetName)
print('worksheet.nrows=',worksheet.nrows)
num_rows = worksheet.nrows - 1
num_cells = len(columns)
excelData = []
curr_row = -1+rowsToSkip

def isRowEmpty(rw):
    empty=1
    for dt in rw:
        if(dt[0]!=0 and dt[0]!=6):
            empty=0
            break
    return empty


while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    #print ('Row:', curr_row)
    curr_cell = -1
    rowData=[]
    for curr_cell_label in columns:
        curr_cell=(ord(curr_cell_label[0])-ord('A')) #Only supports the first 26 columns
        # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
        cell_type = worksheet.cell_type(curr_row, curr_cell)
        cell_value = worksheet.cell_value(curr_row, curr_cell)
        #print ('	', cell_type, ':' , cell_value)
        rowData.append([cell_type,cell_value])
    if not isRowEmpty(rowData):
        excelData.append(rowData)

print ('\n')       
#print (excelData)       

#ExcelData only contains the Data Rows so worksheet.nrows-rowsToSkip

for i in range(0, len(excelData)):
    rowdt=excelData[i]
    for j in range(0, num_cells):
        #print(i,j)
        if(excelData[i][j][0]==0 or excelData[i][j][0]==6):
            k=j+1
            found=0
            while (k<num_cells and found==0):
                #print('i,k=',i,k);
                if(excelData[i][k][0]!=0 and excelData[i][k][0]!=6):
                    found=1
                    #print('found=',found)
                k+=1
            if found:        
                #print(i,j)
                #print(excelData[i-1][j])
                if(excelData[i-1][j][0]!=0 and excelData[i-1][j][0]!=6):
                    excelData[i][j]=excelData[i-1][j]

#At this stage we have the full tree in the data structure without blanks.
print ('\n')                    
for i in range(0,len(excelData)):
    print (excelData[i])       
    
print ('\n')                    
print ('\n')                    
print ('\n')                    

JSonString=''
n='name'
c='children'
s='size'
sv=300       

JSonStringPreamble='{\"'+n+'\":\"'+centreLabel+'\", \"'+c+'\":[\n'
JSonStringEnd=']\n}'
JSonCloseChildrenList=']'
JSonCloseParentNode='}'
JSonSiblingSeparator=',\n'

def prettyIndent(stck):
    ind='    '
    return ind*len(stck) 
    
def getJSonIntermediateLeaf(st):
    return '{\"'+n+'\":\"'+str(st)+'\", \"'+c+'\":[]}'

def getJSonOutermostLeaf(st):
    return '    {\"'+n+'\":\"'+str(st)+'\", \"'+s+'\": '+str(sv)+'}'

def hasChildren(data, i,j):
    if(j>=len(data[i])-1):
        return 0
    if(data[i][j][0]==0 or data[i][j][0]==6):
        return 0
    #The child could be on the same row
    elif(data[i][j+1][0]!=0 and data[i][j+1][0]!=6):
        #print('same row child')
        return 1
    #Or on the row below...
    #First of all check if we are on the last row already and if we are, return 0
    if(i>=len(data)-1):
        return 0
    else:
        #print(data[i][0:j+1])
        #print(data[i+1][0:j+1])
        if(data[i][0:j+1]==data[i+1][0:j+1]): #same branch
            if(data[i+1][j+1][0]!=0 and data[i+1][j+1][0]!=6):
                #print('next row child')
                return 1 #The child is on the next row.
            else: 
                return 0 #This should not happen... two identical rows
        else:
            return 0 #It's a different branch: no child
        
def hasNextSibling(data, i,j):
    if(i>=len(data)-1):
        return 0
    if(data[i][0:j]==data[i+1][0:j]): #same branch
        if(data[i][j]!=data[i+1][j]): #there is a next sibling
            return 1
        else:
            return 0
    else:
        return 0

def hasPreviousSibling(data, i,j):
    if(i<=0):
        return 0
    if(data[i-1][0:j]==data[i][0:j]): #same branch
        if(data[i-1][j]!=data[i][j]): #there is a previous sibling
            return 1
        else:
            return 0
    else:
        return 0
        
#We now need to traverse the tree: "depth first"
JSonString=JSonString+JSonStringPreamble 
print(JSonStringPreamble)
parentstack=[]
for i in range(0, len(excelData)):
    rowdt=excelData[i]
    for j in range(0, len(excelData[i])):
        #print('parentstack lenght:',len(parentstack),', j=',j)
        #print(i,j)
        #if cell is empty then ignore it
        if(excelData[i][j][0]==0 or excelData[i][j][0]==6):
            continue

        #If cell is on Stack at the right depth, then ignore it: it has already been written
        if(len(parentstack)>j): 
            if(parentstack[j]==excelData[i][j][1]):
                continue
        
        #else it may be a new branch!
            else:
                #print ('maybe a new branch.....',len(parentstack),j )
                while (len(parentstack)>j):
                    #print('parentstack:',parentstack)
                    #print('popping one from parentstack')
                    #pop one
                    parentstack.pop()
                    #print('parentstack:',parentstack)
                    #Write closing of a list of children 
                    JSonString=JSonString+JSonCloseChildrenList
                    print(JSonCloseChildrenList)
                    #write closing of of a parent node
                    JSonString=JSonString+JSonCloseParentNode
                    print(JSonCloseParentNode)
                if(hasPreviousSibling(excelData, i,j)): 
                    JSonString=JSonString+JSonSiblingSeparator
                    print(JSonSiblingSeparator)
                
        #Has children? 
        if(hasChildren(excelData,i,j)):
            #print ('##### hasChildren=1')
            #Write opening of a parent node
            JSonString=JSonString+prettyIndent(parentstack)
            JSonString=JSonString+'{\"'+n+'\":\"'+str(excelData[i][j][1])+'\", \"'+c+'\":[\n'
            print ('{\"'+n+'\":\"'+str(excelData[i][j][1])+'\", \"'+c+'\":[')
            #add to stack
            parentstack.append(excelData[i][j][1])
        elif(j<len(excelData[i])-1):
            #write intermediate leaf
            JSonString=JSonString+prettyIndent(parentstack)
            JSonString=JSonString+getJSonIntermediateLeaf(excelData[i][j][1])
            print(getJSonIntermediateLeaf(excelData[i][j][1]))
            if(hasNextSibling(excelData, i,j)): 
                JSonString=JSonString+JSonSiblingSeparator
                print(JSonSiblingSeparator)
        else: 
            #write outermost leaf
            JSonString=JSonString+prettyIndent(parentstack)
            JSonString=JSonString+getJSonOutermostLeaf(excelData[i][j][1])
            print(getJSonOutermostLeaf(excelData[i][j][1]))
            if(hasNextSibling(excelData, i,j)): 
                JSonString=JSonString+JSonSiblingSeparator
                print(JSonSiblingSeparator)

#Tidies up at the end of the cycle
while (len(parentstack)>0):
    parentstack.pop()
    #Write closing of a list of children 
    JSonString=JSonString+JSonCloseChildrenList
    print(JSonCloseChildrenList)
    #write closing of of a parent node
    JSonString=JSonString+JSonCloseParentNode
    print(JSonCloseParentNode)

                
print(JSonStringEnd)        
JSonString=JSonString+JSonStringEnd        

print ('---------------------')
#print (JSonString)
  
f = open(fileout, 'w')
f.write(JSonString)
f.close()








    

print ('\n\n\n**************************************')
    

#Write JSon to outfile



