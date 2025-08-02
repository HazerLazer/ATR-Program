
from openpyxl import load_workbook
import shutil
import os


# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
#                                                                       #
#                                                                       #
#                                                                       #
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #

# label_row     = row that holds the labels of the columns (in data)
# minimum_row   = row that determines the start of the data (and not the headings)
# data          = sheet that contains the input data
# destination   = sheet that contains the output data
# column_labels = mapped list that contains the lowercase names of important columns
# column_map    = mapped list that contains the location of the columns in the output
def populate(label_row, minimum_row, data, destination, column_labels, column_map):
    
    # Locate column indices
    for cell in data[label_row]:
        label = str(cell.value).lower() if cell.value else None
        if label in column_labels:
            column_labels[label] = cell.column


    d_row = 2
    for row_cells in data.iter_rows(min_row=minimum_row, values_only=False):
        row_values = [cell.value for cell in row_cells]
        
        # if row_values[column_labels['tifnum'] - 1] not in (0, None):
        #     destination.cell(row=d_row, column=1).value = row_values[column_labels['tifnum'] - 1]
        # destination.cell(row=d_row, column=2).value = reporting_year

        # Copy values conditionally
        for dest_col, label in column_map.items():
            col_index = column_labels[label] - 1
            value = row_values[col_index] if col_index >= 0 else None
            if value not in (0, None):
                destination.cell(row=d_row, column=dest_col).value = value

        d_row += 1
    return

def Data_Tables(reporting_year):
    shutil.copy('2023 ATR Data Tables - Final 2025 01 03.xlsx', 'Data Tables Copy.xlsx')  # CHANGE IF FILE CHANGES
    copied_table = load_workbook('Data Tables Copy.xlsx')
    
    def section_1():
        data = data_file['Section 1']
        destination = copied_table['Section 1']
        
        column_labels = {
            'tifnum': None, 'tifname': None, 'approvedate': None, 'expiredate': None, 'filename': None
        }
        column_mapping = {
            1: 'tifnum', 2: 'tifname', 3: 'approvedate', 4: 'expiredate', 5: 'filename'
        }
        
        populate(1, 2, data, destination, column_labels, column_mapping)
        return
    
    def section_2():
        data = data_file['Section 2']
        destination = copied_table['Section 2']
        
        column_labels = {
            'TIFnum': None, 'ReptYear': None, 'PrimaryUse': None, 'ComboMix': None, 'IJRL': None
        }
        column_mapping = {
            1: 'TIFnum', 2: 'ReptYear', 3: 'PrimaryUse', 4: 'ComboMix', 5: 'IJRL'
        }
        
        populate(1, 2, data, destination, column_labels, column_mapping)
        
        # add more columns based on criteria (ask)
        
        return
    
    def section_3():
        def section_3_1():
            destination = copied_table['Section 3.1 Other']
            data = data_file['Section 3.1 Other']
            
            column_labels = {
                'TIFnum': None, 'reptyear': None, 'TaxAllocationFundBalance': None, 
                'PropTaxIncr-previous': None, 'PropTaxIncr-current': None, 'PropTaxIncr-cum': None, 
                'Interest-previous': None, 'Interest-current': None, 'Interest-cum': None, 
                'Land/bldg-previous': None, 'Land/bldg-current': None, 'Land/bldg-cum': None, 
                'Bond-previous': None, 'Bond-current': None, 'Bond-cum': None, 
                'Municipal-previous': None, 'Municipal-current': None, 'Municipal-cum': None, 
                'Private-previous': None, 'Private-current': None, 'Private-cum': None, 
                'totalExp/Cash': None, 'DistributionOfSurplus': None, 'Transfers--municipal': None
            }
            
            column_mapping = {
                'TIFnum', 'reptyear', 'TaxAllocationFundBalance', 'PropTaxIncr-previous', 
                'PropTaxIncr-current', 'PropTaxIncr-cum', 'Interest-previous', 
                'Interest-current', 'Interest-cum', 'Land/bldg-previous', 
                'Land/bldg-current', 'Land/bldg-cum', 'Bond-previous', 'Bond-current', 
                'Bond-cum', 'Municipal-previous', 'Municipal-current', 'Municipal-cum', 
                'Private-previous', 'Private-current', 'Private-cum', 
                'totalExp/Cash', 'DistributionOfSurplus', 'Transfers--municipal'
            }
            
            populate(1, 4, data, destination, column_labels, column_mapping)
            return
        
        def section_3_1_Other():
            destination = copied_table['Section 3.1 Other']
            data = data_file['Section 3.1 Other']
            
            column_labels = {
                'TIFnum': None, 'ReptYear': None, 'priorYearsCumulative': None, 
                'NoteProceedsCurrentYear': None, 'noteProceedsCumulative': None, 
                'nonCompliancePayment': None, 'nonComplianceCum': None, 
                'excessReserveRequirement': None, 'excessReserveCum': None, 
                'BABrebate': None, 'BABrebateCum': None, 'collectionReturns': None, 
                'collectionReturnsCum': None, 'creditsExpenditures': None, 
                'creditsExpendituresCum': None
            }
            column_mapping = {
                'TIFnum', 'ReptYear', 'priorYearsCumulative', 
                'NoteProceedsCurrentYear', 'noteProceedsCumulative', 
                'nonCompliancePayment', 'nonComplianceCum', 
                'excessReserveRequirement', 'excessReserveCum', 
                'BABrebate', 'BABrebateCum', 'collectionReturns', 
                'collectionReturnsCum', 'creditsExpenditures', 
                'creditsExpendituresCum'
            }
            
            return
        
        def section_3_2_A():
            destination = copied_table['Section 3.2A']
            data = data_file['Section 3.2a']
            return
        
        def section_3_2_B():
            return
        
        def section_3_3():
            destination = copied_table['Section 3.3']
            data = data_file['Section 3.3']
            return
        
        data_file = load_workbook('2023 ATR - Section 3.1 3.1 Schedule 3.2A and 3.3 Template 7.10.24')
        
        section_3_1()    
        section_3_1_Other()
        section_3_2_A()
        section_3_2_B()
        section_3_3()
        return
    
    def section_4():
        return
    
    def section_5():
        data = data_file['Section 5']
        destination = copied_table['Section 5']
        
        
        column_labels = {
            'project#': None, 'project / iga': None, 'type': None, 'tifnum': None, 
            'rda name normalized': None, 'annual report name': None, 
            'currentyearnewdeals': None, 'ongoing': None, 'complete': None, 
            'currentyearpmts': None, 'estsubsequentyearpmts': None, 
            'pvt 12-31-99 to yr end': None, 'pvt to completion': None, 
            'public 11-1-99 to yearend': None, 'public to completion': None
        }

        # Locate column indices
        label_row = 3
        for cell in data[label_row]:
            label = str(cell.value).lower() if cell.value else None
            if label in column_labels:
                column_labels[label] = cell.column

        d_row = 2
        for row_cells in data.iter_rows(min_row=4, values_only=False):
            row_values = [cell.value for cell in row_cells]
            
            # if row_values[column_labels['tifnum'] - 1] not in (0, None):
            #     destination.cell(row=d_row, column=1).value = row_values[column_labels['tifnum'] - 1]
            # destination.cell(row=d_row, column=2).value = reporting_year

            # Define a mapping between destination column and label
            column_mapping = {
                3: 'project / iga', 4: 'type', 5: 'project#', 6: 'rda name normalized', 
                7: 'annual report name', 10: 'currentyearnewdeals', 11: 'ongoing', 
                12: 'complete', 13: 'currentyearpmts', 14: 'estsubsequentyearpmts', 
                15: 'pvt 12-31-99 to yr end', 16: 'public 11-1-99 to yearend', 
                17: 'pvt to completion', 18: 'public to completion'
            }

            # Copy values conditionally
            for dest_col, label in column_mapping.items():
                col_index = column_labels[label] - 1
                value = row_values[col_index] if col_index >= 0 else None
                if value not in (0, None):
                    destination.cell(row=d_row, column=dest_col).value = value

            d_row += 1
        
        
        label_row = 3
        for cell in data[label_row]:
            if cell.value and str(cell.value).lower() == 'project#':
                proj_num_col = cell.column
            if cell.value and str(cell.value).lower() == 'project / iga':
                project_col = cell.column
            if cell.value and str(cell.value).lower() == 'type':
                type_col = cell.column
            if cell.value and str(cell.value).lower() == 'tifnum':
                num_col = cell.column
            if cell.value and str(cell.value).lower() == 'rda name normalized':
                rda_col= cell.column
            if cell.value and str(cell.value).lower() == 'annual report name':
                ar_name_col = cell.column
            if cell.value and str(cell.value).lower() == 'currentyearnewdeals':
                cynd_col = cell.column
            if cell.value and str(cell.value).lower() == 'ongoing':
                ong_col = cell.column
            if cell.value and str(cell.value).lower() == 'complete':
                com_col = cell.column
            if cell.value and str(cell.value).lower() == 'currentyearpmts':
                cyp_col = cell.column
            if cell.value and str(cell.value).lower() == 'estsubsequentyearpmts':
                esyp_col = cell.column
            if cell.value and str(cell.value).lower() == 'pvt 12-31-99 to yr end':
                pvt_col = cell.column
            if cell.value and str(cell.value).lower() == 'pvt to completion':
                pvt_com_col = cell.column
            if cell.value and str(cell.value).lower() == 'public 11-1-99 to yearend':
                pub_col = cell.column
            if cell.value and str(cell.value).lower() == 'public to completion':
                pub_com_col = cell.column

        d_row = 2
        for row in data.iter_rows(values_only=True):        
            if row < 4:
                continue
            
            num = data.cell(row=row, column=num_col).value     
            if num != 0 and num != None:
                destination.cell(row=d_row, column=1).value = num
            
            destination.cell(row=d_row, column=2).value = reporting_year
            
            project = data.cell(row=row, column=project_col).value     
            if project != 0 and project != None:
                destination.cell(row=d_row, column=3).value = project
            
            type = data.cell(row=row, column=type_col).value     
            if type != 0 and type != None:
                destination.cell(row=d_row, column=4).value = type
            
            proj_num = data.cell(row=row, column=proj_num_col).value     
            if proj_num != 0 and proj_num != None:
                destination.cell(row=d_row, column=5).value = proj_num
            proj_num_col
            
            rda = data.cell(row=row, column=rda_col).value     
            if rda != 0 and rda != None:
                destination.cell(row=d_row, column=6).value = rda
            
            ar_name = data.cell(row=row, column=ar_name_col).value     
            if ar_name != 0 and ar_name != None:
                destination.cell(row=d_row, column=7).value = ar_name
            
            cynd = data.cell(row=row, column=cynd_col).value     
            if cynd != 0 and cynd != None:
                destination.cell(row=d_row, column=10).value = cynd
            
            ong = data.cell(row=row, column=ong_col).value     
            if ong != 0 and ong != None:
                destination.cell(row=d_row, column=11).value = ong
            
            com = data.cell(row=row, column=com_col).value     
            if com != 0 and com != None:
                destination.cell(row=d_row, column=12).value = com
            
            cyp = data.cell(row=row, column=cyp_col).value     
            if cyp != 0 and cyp != None:
                destination.cell(row=d_row, column=13).value = cyp
            
            esyp = data.cell(row=row, column=esyp_col).value     
            if esyp != 0 and esyp != None:
                destination.cell(row=d_row, column=14).value = esyp
            
            pvt = data.cell(row=row, column=pvt_col).value     
            if pvt != 0 and pvt != None:
                destination.cell(row=d_row, column=15).value = pvt
            
            pvt_com = data.cell(row=row, column=pvt_com_col).value     
            if pvt_com != 0 and pvt_com != None:
                destination.cell(row=d_row, column=17).value = pvt_com
            
            pub = data.cell(row=row, column=pub_col).value     
            if pub != 0 and pub != None:
                destination.cell(row=d_row, column=16).value = pub
            
            pub_com = data.cell(row=row, column=pub_com_col).value     
            if pub_com != 0 and pub_com != None:
                destination.cell(row=d_row, column=18).value = pub_com
            
            d_row += 1
        
        return
    
    def section_6():
        return
    
    def section_7():
        return
    
    destination = copied_table['Section 3.1']
    data = data_file['Section 3.1']
    
    label_row = 1
    for cell in data[label_row]:
        if cell.value and str(cell.value).lower() == '': 
            _col = cell.column
        if cell.value and str(cell.value).lower() == 'creditsexpenditurescum': 
            ce_cum_col = cell.column
    
    
    
    # Populate Section 5 of the Data Table output
    data = data_file['ATR 2023']
    destination = copied_table['Section 5']
    
    data_file = load_workbook('Master Input File for Python 2025 01 15.xlsx')
    return
    