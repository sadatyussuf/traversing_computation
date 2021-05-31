import openpyxl
import math

# lOADING THE EXCEL WORKBOOK
wb = openpyxl.load_workbook('traverse.xlsx')

sheet = wb['Sheet1']
sheet.title = 'Original'


numberOfSides = 6
sumOfActualAngles = 180 * (numberOfSides - 2)
Misclose = 3 * (10 / 3600) * math.sqrt(numberOfSides)

if Misclose < ((1 / 60) + (50 / 3600)):
    Misclose = (1 / 60)
else:
    Misclose= Misclose

adjPerAngle = -(Misclose/numberOfSides)

# Computing for the Adjusted misclose per Angle in Cell D
for i in range(2,2+numberOfSides):
    sheet.cell(row=i, column=4).value = adjPerAngle


# Computing for the sum total of the Angles in Cell B
count = 0
for i in range(2,2+numberOfSides):
    angles = sheet.cell(row=i,column=2).value
    angle = eval(angles[1:])
    count += angle
sheet['B8'].value = count

# Computing for the Adjusted Angles in Cell E
count = 0
for i in range(2,2+numberOfSides):
    angles = sheet.cell(row=i, column=2).value
    sheet.cell(row=i,column=5).value = eval(angles[1:]) + float(sheet.cell(row=i,column=4).value)
    count += sheet.cell(row=i,column=5).value
sheet['E8'].value = count




wb.save('test1.xlsx')