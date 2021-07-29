#########################################################
# Rapid Fuzzy Version
#########################################################
import pandas as pd
import numpy as np
import openpyxl
from rapidfuzz.fuzz import token_set_ratio as rapid_token_set_ratio
from rapidfuzz import process as process_rapid
from rapidfuzz import utils as rapid_utils
import time

def excel_sheet_to_dataframe(path):
    '''
        Loads sheet from Excel workbook using openpyxl
    '''
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    data = ws.values
     # Get the first line in file as a header line
    columns = next(data)[0:]
    
    return pd.DataFrame(data, columns=columns)

def process_rapid_fuzz(data):
    '''
        Process using rapid fuzz rather than fuzz_wuzzy
    '''
    series = (rapid_utils.default_process(d) for d in data)       # Pre-process to make lower-case and remove non-alphanumeric 
                                                                   # characters (generator)
    processed_data = pd.Series(series)   

    clean_rapid = []
    threshold = 80 
   
    for query in processed_data:
        scores = process_rapid.extract(query, processed_data, scorer=rapid_token_set_ratio, score_cutoff=threshold)
        
        m = max(scores[:2], key = lambda k:len(k[0]))                # Of up to two matches above threshold, takes longest
        clean_rapid.append(m[-1])                                    # Saving the match index
        
    clean_rapid = set(clean_rapid)                                   # remove duplicate indexes

    return data[clean_rapid]                                         # Get actual values by indexing to Pandas Series

################ Testing
t0 = time.time()
df = excel_sheet_to_dataframe('Duplicates1.xlsx')   # Using Excel file in working folder

# Desired data in body column
data = df['Body'].dropna()                                           # Dropping None rows (few None rows at end after Excel import)

result_fuzzy_rapid = process_rapid_fuzz(data)
print(f'Elapsed time {time.time() - t0}')
