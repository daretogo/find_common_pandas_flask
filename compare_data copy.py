import pandas, pandas_usaddress, pdb, flask



newvfilename = r'C:\Users\Aaron\OneDrive\find_common_pandas_flask\newvisitor\NewV_Nov_Dec_Jan.xlsx'
newvsheetname = 'Sheet1'
newvstreetcolumn = 'Mailing Street'
newvcitycolumn = 'Mailing City'
newvzipcolumn = 'Mailing Zip'

homebuyersfilename = r'C:\Users\Aaron\OneDrive\find_common_pandas_flask\homebuyer\Homebuyer_Hotlist_Nov_11_to_Dec_8_2018.xlsx'
homebsheetname = 'Parker Only'
homebstreetcolumn = 'Situs Street Address'
homebcitycolumn = 'Situs City'
homebzipcolumn = 'Situs Zip Code'

resultfilename = r'C:\Users\Aaron\OneDrive\find_common_pandas_flask\result.xlsx'

###################################################################################################################################
##################^^^^^^^^^^^^^^  Edit the variables above this line ^^^^^^^^^^####################################################
###################################################################################################################################

#read in the new visitor list
df_newv = pandas.read_excel(newvfilename, sheet_name=newvsheetname)

#parse the new visitor list into standardized address bits based on the existing address fields
df_newv_parsed = pandas_usaddress.tag(df_newv, [newvstreetcolumn, newvcitycolumn, newvzipcolumn], granularity='medium', standardize=True)

#read in the home buyers list
df_homebuy = pandas.read_excel(homebuyersfilename, sheet_name=homebsheetname)

#parse homebuyer list into standardized address bits too. 
df_homebuy_parsed = pandas_usaddress.tag(df_homebuy, [homebstreetcolumn, homebcitycolumn, homebzipcolumn], granularity='medium', standardize=True)

#the magic line - perform the inner join matching on address number, city name (placename) and the street name.  This 
#effectively eliminates missed matches due to Cir vs Cr or St vs Street comparisons in the address as we have extracted those values away. 
df_merged = pandas.merge(df_homebuy_parsed, df_newv_parsed, on=['AddressNumber', 'PlaceName', 'StreetName'], how='inner')
#df_merged.dropna(inplace=True)

writer = pandas.ExcelWriter(resultfilename, engine='xlsxwriter')
df_merged.to_excel(writer, sheet_name='Parsed and Merged Raw Data')
# Close the Pandas Excel writer and output the Excel file.
writer.save()
