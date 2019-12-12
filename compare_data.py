import pandas, pandas_usaddress, pdb, flask
from flask import Flask, request, render_template


from flask import Flask
app = Flask(__name__)

#############################################################################################################################
def Comparison(newv_filename, newv_sheet, newv_street_column, newv_city_column, newv_zip_column, homeb_filename, homeb_sheet, homeb_street_column, homeb_city_column, homeb_zip_column, result_filename):

    #read in the new visitor list
    df_newv = pandas.read_excel(newv_filename, sheet_name=newv_sheet)

    #parse the new visitor list into standardized address bits based on the existing address fields
    df_newv_parsed = pandas_usaddress.tag(df_newv, [newv_street_column, newv_city_column, newv_zip_column], granularity='medium', standardize=True)

    #read in the home buyers list
    df_homebuy = pandas.read_excel(homeb_filename, sheet_name=homeb_sheet)

    #parse homebuyer list into standardized address bits too. 
    df_homebuy_parsed = pandas_usaddress.tag(df_homebuy, [homeb_street_column, homeb_city_column, homeb_zip_column], granularity='medium', standardize=True)

    #the magic line - perform the inner join matching on address number, city name (placename) and the street name.  This 
    #effectively eliminates missed matches due to Cir vs Cr or St vs Street comparisons in the address as we have extracted those values away. 
    df_merged = pandas.merge(df_homebuy_parsed, df_newv_parsed, on=['AddressNumber', 'PlaceName', 'StreetName'], how='inner')

    writer = pandas.ExcelWriter(result_filename, engine='xlsxwriter')
    df_merged.to_excel(writer, sheet_name='Parsed and Merged Raw Data')
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    return(df_merged)
#####################################################################################################################################

@app.route('/pdb')
def pdb():
   """Enter python debugger in terminal"""

   import sys
   print("\n'/pdb' endpoint hit. Dropping you into python debugger. globals:")
   print("%s\n" % dir(sys.modules[__name__]))
   import pdb; pdb.set_trace()

   return 'After PDB debugging session, now execution continues...'

@app.route("/", methods=['GET', 'POST'])
def index():
    
    if request.method == 'GET':
        return render_template('index.html')

    if request.method == 'POST':
        #this creates a dictionary of form_inputs with a kvp from the form in the template
        form_inputs = request.form
        
        #this executes the comparison function with the supplied data from the user
        Comparison(form_inputs['newvfilename'],form_inputs['newvsheetname'],form_inputs['newv_street'],form_inputs['newv_city'],form_inputs['newv_zip'],form_inputs['homebuyersfilename'],form_inputs['homebsheetname'],form_inputs['homeb_street'],form_inputs['homeb_city'],form_inputs['homeb_zip'],form_inputs['resultpath'])

        return render_template(
        'result.html')

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=80)