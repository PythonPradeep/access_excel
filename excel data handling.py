#import needed libraries
import openpyxl
#select the file path
file_path = "/home/pradeepthamizh/Desktop/data handling/task.xlsx"
#use load_workbook for use our file (the spreadsheet)
file = openpyxl.load_workbook(file_path)
#use active for select the available sheet in spreadsheet
new_sheet = file.active

#select the value of the student in the file
#student1
cell_obj = new_sheet['A2': 'F2'] #select which column to which column
for cell1, cell2, cell3, cell4, cell5, cell6 in cell_obj: #In here using boundaries for the iteration
     marks = cell2.value + cell3.value + cell4.value + cell5.value + cell6.value #adding the 1st student marks

#student2
cell_obj = new_sheet['A3': 'F3']
for cell1, cell2, cell3, cell4, cell5, cell6 in cell_obj:
    marks2 = cell2.value + cell3.value + cell4.value + cell5.value + cell6.value #adding the 2nd student marks

#student3
cell_obj = new_sheet['A4': 'F4']
for cell1, cell2, cell3, cell4, cell5, cell6 in cell_obj:
    marks3 = cell2.value + cell3.value + cell4.value + cell5.value + cell6.value #adding the 3rd student marks

#student4
cell_obj = new_sheet['A5': 'F5']
for cell1, cell2, cell3, cell4, cell5, cell6 in cell_obj:
    marks4 = cell2.value + cell3.value + cell4.value + cell5.value + cell6.value #adding the 4th student marks

#student5
cell_obj = new_sheet['A6': 'F6']
for cell1, cell2, cell3, cell4, cell5, cell6 in cell_obj:
    marks5 = cell2.value + cell3.value + cell4.value + cell5.value + cell6.value #adding the 5th student marks

#insert the calculated values in the cells
new_sheet['G2'] = marks
new_sheet['G3'] = marks2
new_sheet['G4'] = marks3
new_sheet['G5'] = marks4
new_sheet['G6'] = marks5

#save the changes
file.save('task.xlsx')
