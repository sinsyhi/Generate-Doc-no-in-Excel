import xlsxwriter
import pandas as pd

def generate_text_range(start, end, prefix):
    numbers_array = [f"{prefix}{i:04d}" for i in range(start, end + 1)]
    return numbers_array

prefix = input('Enter product number: ')
start = int(input('Enter start number: '))
end = int(input('Enter end number: '))

#Part of creating excel file
dataframe = pd.DataFrame({'Product number': generate_text_range(start, end, prefix)}) #Create dataframe
writer = pd.ExcelWriter(f'Product number of {prefix}.xlsx', engine='xlsxwriter') #Create excel file
dataframe.to_excel(writer, sheet_name='Sheet1', index=False) #Write dataframe to excel file
writer.close()  # Changed from writer.save() to writer.close()
