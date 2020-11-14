import os, time, openpyxl

print("Follow these steps to format your spreadsheet")
input("Press Enter when finished with the step")

print("\nStep One: Copy your spreadsheet into the folder that's just about to open up (Python's active directory)")
path = os.getcwd()
realpath = os.path.realpath(path)
time.sleep(5)
os.startfile(realpath)
input("Press Enter when finished with the step")

print("\nStep Two: Name your sheet containing the data as 'Input' (capitalized, without quotes).\n Make sure that the data is in this format: \n" + """
BDE47 gene | BDE47 symbol | BDE47 logFC | BDE47 FDR | IPP gene | IPP symbol | etc
----------------------------------------------------------------------------------
ENSDAR0394 | itga4        | 1.3948      | 0.00283   | ENSDAR029| apba1b     | etc
----------------------------------------------------------------------------------
ENSDAR034 | mcu           | 1.4948      | 0.00863   | ENSDAR024| numed1     | etc
""")
input("Press Enter when finished with the step")


fileName = input("\nStep Three: Enter the full name of your spreadsheet as it appears in Python's directory (e.g. Example.xlsx). Press Enter to enter the name: ")

print("\n Step Four: Enter every chemical in your sheet (BDE47, TCPP, etc) by entering a name and pressing Enter. When finished, press Enter on a blank line.")
chemList = []
while "" not in chemList:
    entered = input("Enter a chemical here: ")
    chemList.append(entered)
del chemList[-1]
print(chemList)

print("\nStep Five: Make sure to close the spreadsheet before continuing")
input("Press Enter when the spreadsheet is closed")

print("\nPlease enjoy this hold music while your spreadsheet is being formatted")
print("*Elevator Music*")

#code

print("\nYour code should now be formatted in the 'Output' sheet.\n If you want to thank the creator of this script, mail an uninsulated raw fish to 1600 Pennsylvania Avenue NW, Washington, DC 20500")
print("\nIf there are any issues, ask Konoha. Or if you're feeling adventurous, open the code using IDLE and make your own tweaks.")

wb = openpyxl.load_workbook(fileName)
outputFileName = 'out.xlsx'

#Code to add sought-after genes to a target list
#sheet = wb['Gene List']
#targetList = []
#for i in range(1, sheet.max_row + 1):
#    targetList.append(sheet.cell(i, 1).value)
sheet = wb['Input']
targetList = []
for column in range(2,40,4):
    for row in range(2, sheet.max_row + 1):
        checkingGene = sheet.cell(row, column).value
        if checkingGene not in targetList:
            targetList.append(checkingGene)


#Code to create a label line on Row 1
sheet = wb['Output']
sheet.cell(1,1).value = "Gene"
sheet.cell(1,2).value = "Symbol"
c = 3
for chem in chemList:
    sheet.cell(1, c).value = chem + " logFC"
    sheet.cell(1, c + 1).value = chem + " FDR"
    c += 2


#Code to locate genes in Input matching targetList
inputSheet = wb['Input']
outputRowCounter = 2
totalHitCounter = 0
#removed targetList loop
hitList = [] #List of hits for a specific gene, contains columnNumber, gene, symbol, logFC and FDR. Blanks out after each target
for column in range(2, 40, 4):
    print(str(sheet.max_column+1))
    foundInColumn = False #Resets at the start of each column
    for row in range(2, inputSheet.max_row + 1): #change 1 to two bc it started on the wrong row
        currentGene = inputSheet.cell(row, column).value
        if currentGene in targetList:
            #print(currentGene + " found in row " + str(row) + " and column " + str(column))
            foundInColumn = True
            totalHitCounter += 1
            gene = inputSheet.cell(row, column - 1).value
            symbol = currentGene
            logFC = inputSheet.cell(row, column + 1).value
            fdr = inputSheet.cell(row, column + 2).value
            hitList.append([column, gene, symbol, logFC, fdr])
            #Break out of row for-loop for increased efficiency
            print("column is: " + str(column))
            print("row is: " + str(row))
            print("logFC is: " + str(logFC))
            print("FDR is: " + str(fdr))
    if foundInColumn == False: #If the gene isn't found for a specific chem, fill hitList with an empty list
        hitList.append([column, " ", " ", " ", " "])

#print("hitList contains:")
#print(hitList)



#Code to fill Output
sheet = wb['Output']

# sheet.cell(outputRowCounter, 1).value = hitList[0][1]
# sheet.cell(outputRowCounter, 2).value = hitList[0][2]
# hitListPlace = 0
# for item in hitList:
#     sheet.cell(outputRowCounter, hitListPlace * 4 + 3).value = hitList[hitListPlace][3]
#     sheet.cell(outputRowCounter, hitListPlace * 4 + 4).value = hitList[hitListPlace][4]
#
#     hitListPlace += 1
# outputRowCounter += 1

row = 2
for target in targetList:
    for item in hitList:
        if item[2] == target:
            hitCol = (item[0] / 2) + 2
            sheet.cell(row, 1).value = item[1] #fill gene name
            sheet.cell(row, 2).value = item[2] #fill gene symbol
            #add other important values
            sheet.cell(row, hitCol).value = item[3]
            sheet.cell(row, hitCol+1).value = item[4]
    row += 1

print("All Done, Boss! We found a total of " + str(totalHitCounter) + " hits across " + str(len(targetList)) + " genes")
wb.save(fileName)
