import openpyxl
import math

# lOADING THE EXCEL WORKBOOK
wb = openpyxl.load_workbook('Book1.xlsx')
sheet = wb['Sheet1']
sheet.title = 'Original'



numberOfSides = 4
pageNum = numberOfSides+3
sheet[f'A{pageNum}'].value = 'Total'

# Computing for the sum total of the Angles in Cell B
count = 0
for i in range(2,2+numberOfSides):
    angles = sheet.cell(row=i,column=2).value
    angle = eval(angles[1:])
    count += angle

sheet[f'B{pageNum}'].value = count
# sheet['B8'].value = count



# ------------------------------------------------
# Adjustment of angular error
sumObservedAngles = count 
sumActualAngles = 180 * (numberOfSides - 2)
totalError =-(sumObservedAngles - sumActualAngles) 

sheet[f'D{pageNum}'].value =totalError

adjPerAngle  = totalError/numberOfSides
# -----------------------------------------------------



# ----------------------------------------------------------------
# Computing for the Total Length in Cell C
sheet.cell(row=1, column=3).value = 'Distance'
count = 0
for i in range(2,2+numberOfSides):
    distance = sheet.cell(row=i, column=3).value
    count += distance

sheet[f'C{pageNum}'].value = count
# -------------------------------------------------------------



# --------------------OLd Way Of Adjusting Error--------------------------------
#  Miscole to correct the included angle
# Misclose = 3 * (10 / 3600) * math.sqrt(numberOfSides)

# if Misclose < ((1 / 60) + (50 / 3600)):
#     Misclose = (1 / 60)
# else:
#     Misclose= Misclose

# adjPerAngle = -(Misclose/numberOfSides)
# ----------------------------------------------------------------



# -------------------------------------------------------------
# Computing for the Adjusted misclose per Angle in Cell D
sheet.cell(row=1, column=4).value = 'Correction'
for i in range(2,2+numberOfSides):
    sheet.cell(row=i, column=4).value = adjPerAngle
# ----------------------------------------------------------------------



# Computing for the Adjusted Angles in Cell E
sheet.cell(row=1, column=5).value = 'Corrected Angles'
count = 0
for i in range(2,2+numberOfSides):
    angles = sheet.cell(row=i, column=2).value
    sheet.cell(row=i,column=5).value = eval(angles[1:]) + float(sheet.cell(row=i,column=4).value)
    count += sheet.cell(row=i,column=5).value

sheet[f'E{pageNum}'].value = math.ceil(count)
# sheet['E8'].value = math.ceil(count)



# ----------------------------------------------------------------
# Computing for the Bearings in Cell F
sheet.cell(row=1, column=6).value = 'W.C.B'


# -------------------------------------------------------------

wb.save('test1.xlsx')