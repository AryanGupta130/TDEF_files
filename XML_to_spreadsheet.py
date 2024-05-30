import xml.etree.ElementTree as ET ## allows me to pull info from xml file
import openpyxl as op
from openpyxl import load_workbook ## will use this later
import os

## have to put in for loop for every tdef
tree = ET.parse('TDEF_files/AB12.tdef') ## parses through tree in XML file
root = tree.getroot()

def find_namespace_declaration(root): ## this is the namespace helper function 
    name = root.tag.split('}')[0][1:]
    return name

## has to go in for loop as every tdef can have diff name space
namespace_name = find_namespace_declaration(root) ## call to the helper funciton 
namespace = {'ns': str(namespace_name)}

## loads open the excel file that i want to edit
wb = op.load_workbook('C:\\Users\\z004ymfp\\Downloads\\TDEF_TESTER_common1.xlsx')
sheet = wb.active

##Dictionary to match the column header name with the val to get out of TDEF file
## No need for 'Analyte\nStability' since it is left blank
header_val = {'Test\nName': 'TestName', 
              'Test\nType': 'TestTypeID',
              'Test\nVersion': "TestVersion", 
              'Display\nName':"TestName",
              'Print\nName':"TestName",
              'LIS\nCode': "TestName", 
              'LOINC': "LOINC", 
              'Status': "StatusID", 
              'Result\nReview\nMode': "ResultReviewMode", 
              'Reuse\nResult': "ReuseResult", 
              'Result\nTime\nLimit': "ResultTimeLimit", }

## this will give me a list of all the tdef files
tdef_files = []
files_in_folder = os.listdir('TDEF_files')
for file in files_in_folder:
    # Check if the file has the ".tdef" extension
    if file.endswith('.tdef'):
        # If it does, add the file to the list of TDEF files
        tdef_files.append(file)
                  
count = 0
tdef_files_len = len(tdef_files)
headers = ['Test\nName', 'Test\nType', 'Test\nVersion', 'Display\nName', 'Print\nName', 'LIS\nCode',
           'LOINC', 'Status', 'Result\nReview\nMode', 'Reuse\nResult', 'Result\nTime\nLimit', 'Analyte\nStability']
tdef_files_iter = iter(tdef_files)  # convert tdef_files to an iterator

## need to iterate through all the xml files and see if it is luminometer
## need to also check if it is system monitoring
filtered_tdefs = []
systemMonitoring_true = []
systemMonitoring_false = []
for tdef_file in tdef_files:
    ## need to look at the detection type to see if it is a luminometer
    tree = ET.parse(f'TDEF_files/{tdef_file}')
    root = tree.getroot()
    namespace_name = find_namespace_declaration(root)
    namespace = {'ns': str(namespace_name)}
    detection_type = root.findtext(f'.//ns:DetectionType', namespaces=namespace) #indicates luminometer or not
    system_monitoring = root.findtext(f'.//ns:IsSystemMonitoringTest', namespaces=namespace) # indicates whether it is system monitoring or not
    print(system_monitoring, tdef_file)
    if detection_type =='Luminometer': ## if not luminometer, we don't consider it 
        filtered_tdefs.append(tdef_file)
        
        if system_monitoring == 'false':
            systemMonitoring_false.append(tdef_file)
        else:
            systemMonitoring_true.append(tdef_file)
        
    
    
    

row_num = 3  # start from 3rd row
while row_num <= tdef_files_len:  # use the number of XML files
    xml_file = next(tdef_files_iter)  # get the next XML file
    tree = ET.parse(f'TDEF_files/{xml_file}')
    root = tree.getroot()
    namespace_name = find_namespace_declaration(root)
    namespace = {'ns': str(namespace_name)}
    
    row = sheet[row_num]
    for col_num, cell in enumerate(row, start=1):
        print(f"Processing row {row_num}, column {col_num}, count{count}")

        header_value = headers[count]

        if header_value == 'Analyte\nStability':
            value_to_add = ""
        else:
            print(header_value)
            header = header_val[header_value]
            value_to_add = root.findtext(f'.//ns:{header}', namespaces=namespace) 
            
            if header_value == "Reuse\nResult":
                if value_to_add == "true":
                    value_to_add = 'X'
                else:
                    value_to_add = ''
        if count == 11:
            count = 0
            row_num += 1 
        else:
            count += 1 
        if row_num == len(tdef_files) + 2:
            break
        # Write the value to the cell
        print(value_to_add)
        c = sheet.cell(row = row_num, column = col_num)
        c.value = value_to_add


wb.save('C:\\Users\\z004ymfp\\Downloads\\TDEF_TESTER_common1.xlsx')



