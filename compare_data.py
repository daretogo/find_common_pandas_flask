import pandas, pandas_usaddress, pdb, flask

#read in the new visitor list
df_newv = pandas.read_excel(r'C:\Users\Aaron\OneDrive\findDuplicateXLSXpandasflask\data.xlsx', sheet_name='newv')

#parse the new visitor list into standardized address bits based on the existing address fields
df_newv_parsed = pandas_usaddress.tag(df_newv, ['Mailing Street', 'Mailing City', 'Mailing State', 'Mailing Zip'], granularity='medium', standardize=True)

#read in the home buyers list
df_homebuy = pandas.read_excel(r'C:\Users\Aaron\OneDrive\findDuplicateXLSXpandasflask\data.xlsx', sheet_name='homebuy')

#parse homebuyer list into standardized address bits too. 
df_homebuy_parsed = pandas_usaddress.tag(df_homebuy, ['Mailing Street', 'Mailing City', 'Mailing State', 'Mailing Zip'], granularity='medium', standardize=True)

#the magic line - perform the inner join matching on address number, city name (placename) and the street name.  This 
#effectively eliminates missed matches due to Cir vs Cr or St vs Street comparisons in the address as we have extracted those values away. 
df_merged = pandas.merge(df_homebuy_parsed, df_newv_parsed, on=['AddressNumber', 'PlaceName', 'StreetName'], how='inner')
#df_merged.dropna(inplace=True)

writer = pandas.ExcelWriter(r'C:\Users\Aaron\OneDrive\findDuplicateXLSXpandasflask\result.xlsx', engine='xlsxwriter')
df_merged.to_excel(writer, sheet_name='Parsed and Merged Raw Data')
# Close the Pandas Excel writer and output the Excel file.
writer.save()
