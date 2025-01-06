import xlsxwriter
import pandas as pd

def generate_text_range(start, end, prefix):
    numbers_array = [f"{prefix}{i:04d}" for i in range(start, end + 1)]
    return numbers_array

station = input('Enter station name: ')
prefix = input('Enter product number prefix: ')
start = int(input('Enter start number: '))
end = int(input('Enter end number: '))

#Part of creating excel file
dataframe = pd.DataFrame({'Product number': generate_text_range(start, end, prefix)}) #Create dataframe
writer = pd.ExcelWriter(f'Product_number_List_{station}.xlsx', engine='xlsxwriter') #Create excel file
dataframe.to_excel(writer, sheet_name='Sheet1', index=False) #Write dataframe to excel file
writer.close() #Save excel file