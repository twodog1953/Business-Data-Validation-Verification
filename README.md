# Data-Validation-Verification
Process data for work purposes. Mostly aim for security scheduling and AP/AR data. Will add in functions and codes incrementally throughout my work process. 

Description: The two original datasets are pulled from the Winteam database and the Sage 50 database. These two db are not automatically synced, and the dispatch department, upon receiving the call from employees asking for an info change, mostly just update the data on Winteam without informing other departments. Therefore, it requires constant manual comparison and examination to ensure key data fields (such as address, name) of employees to be updated. Without adjustments, many employees who requested an info change on phone would still have outdated data in the Sage 50 database, creating potential problems for receiving paystubs mailing or payroll tax calculation. 


Requirements: Following requirements are needed to be satisfied to run this code:

Packages to be installed:
1. pandas
2. openpyxl
3. Levenshtein

Running the code:
The code for 'Employees Records Verification.py' could be run to create a finalized excel sheet "check_out.xlsx" after data cleaning, restructuring and merging, which contains all designated columns from both database, and Levenshtein scores for certain columns. 

The sample of end result is as following: 

![Screenshot 2024-10-28 144949](https://github.com/user-attachments/assets/f9249390-4022-48b1-a665-4b0b2323221d)


As shown on the screenshot, the Levenshtein scores between key columns of the two datasets are calculated, and the color scale are set in Excel based on the scores to offer a visual comparison between the discrepancies of the two datasets. 
