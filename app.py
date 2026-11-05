from analyticskit import *
import pandas as pd
import os
import numpy as np
import openpyxl
Alloc_path = select_file('Please Select the Student Report file.')
Heirarchy_path = select_file('Please Select the Updated Program Heirarchy file.')
Tableau_path = select_file('Please Select the Tableau file.')
directory = os.getcwd()
Namelist = ['Name']
masks = [False]
file_name = input_form('Input what the file should be called',inputs=Namelist,masks=masks)
Allocdf = pd.read_excel(Alloc_path)
Tableaudf = pd.read_excel(Tableau_path)
Heirarchydf = pd.read_excel(Heirarchy_path)
GroupedAlloc = Allocdf.groupby('BYU ID')['ISP Allocation'].sum().reset_index()
GroupedTableau = Tableaudf.groupby(['BYUID','Program Operating Unit Code'])['Paid Amount'].sum().reset_index()
Merged1 = GroupedAlloc.merge(GroupedTableau,how='inner',left_on='BYU ID',right_on='BYUID')
Merged2 = GroupedAlloc.merge(GroupedTableau,how='left_anti',left_on='BYU ID',right_on='BYUID')
Merged3 = GroupedAlloc.merge(GroupedTableau,how='right_anti',left_on='BYU ID',right_on='BYUID')
Merged2 = pd.concat([Merged2,Merged3],axis=0,ignore_index=True)
del Merged1['BYUID']
del Merged1['Paid Amount']
AllocByCode = Merged1.groupby('Program Operating Unit Code')['ISP Allocation'].sum().reset_index()
ProgramSummary = Heirarchydf.merge(AllocByCode,how='inner',left_on='Operating Unit',right_on='Program Operating Unit Code')
del ProgramSummary['Project Hierarchy Name']
del ProgramSummary['Project Code']
del ProgramSummary['Operating Unit']
del ProgramSummary['ISP Admin Name']
del ProgramSummary['Program Operating Unit Code']
del ProgramSummary['Manager']
Finalname = file_name['Name']+'.xlsx'
filepath = os.path.join(directory, Finalname)
with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
    ProgramSummary.to_excel(writer, sheet_name='Program Summary', index=False)
    Merged2.to_excel(writer, sheet_name='Unmatched Students', index=False)
    Allocdf.to_excel(writer, sheet_name='alloc_isp', index=False)
    Tableaudf.to_excel(writer, sheet_name='Tableau Report', index=False)
    Heirarchydf.to_excel(writer, sheet_name='Program Hierarchy', index=False)

