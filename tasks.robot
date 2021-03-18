*** Settings ***
Documentation   A robot to handle sales report excel from downloads folder.
# the purpose is to make sales report anonymous and calculate total sum
# first the file need to be saved in Downloads folder as salesreport.xlsx
# the final report will be saved to operating path

Library   Collections
Library   RPA.Excel.Files
Library   RPA.FileSystem
Library   RPA.Tables
Library   RPA.Excel.Application

*** Variables ***
${path}=   C:\\Users\\ATJ\\Downloads\\salesreport.xlsx

*** Keywords ***
Get File From Folder
   [Arguments]   ${path}
   RPA.Excel.Files.Open Workbook    ${path}
   ${content}=   Read Worksheet As Table   Taul1   header=True
   [Return]   ${content}
   Close Workbook

*** Keywords ***
Handle Excel
   [Arguments]   ${content}
   Pop Table Column   ${content}   column=Customer Name   as_list=False
   Pop Table Column   ${content}   column=Customer ID   as_list=False
   Pop Table Column   ${content}   column=Invoice Date   as_list=False
   Pop Table Column   ${content}   column=Invoice Due Date   as_list=False
   Pop Table Column   ${content}   column=Invoice Creator ID  as_list=False
   Pop Table Column   ${content}   column=Invoice Approver Name  as_list=False

*** Keywords ***
Calculate Sums
   [Arguments]   ${content}
   @{sums}=   Get Table Column   ${content}    column=EUR   as_list=True
   ${invoiceSums}   Set variable   ${0}
   FOR  ${sum}   IN   @{sums}
      ${invoiceSums}=   Evaluate    ${invoiceSums} + ${sum}
   END
   [Return]   ${invoiceSums}

*** Keywords ***
Save New Excel File
   [Arguments]   ${content}   ${invoiceSums}
   Add Table Row    ${content}   TOTAL
   Add Table Row    ${content}   ${invoiceSums}
   Create Workbook   salesreport_final.xlsx
   Create Worksheet   Tulokset
   Append Rows To Worksheet    ${content}   Tulokset   header=True
   Save Workbook

*** Tasks ***
Read excel and handle it
    ${content}=   Get File From Folder   ${path}
    Handle Excel   ${content}
    ${invoiceSums}=   Calculate Sums   ${content}
    Save New Excel File    ${content}    ${invoiceSums}
