from flask import Flask, jsonify
from flask import json
from flask import request
from flask_cors import CORS
import pandas as pd
from datetime import datetime
from flask import send_file
from oauth2client.service_account import ServiceAccountCredentials
import gspread


app = Flask(__name__)
CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'

@app.route('/preview', methods=['POST'])
def application():
    global parsed_data
    parsed_data = json.loads(request.data)

    #simulation of code barcode
    df_artikel = pd.read_csv('data/database_artikel.csv')
    df_code = pd.DataFrame(parsed_data,columns=['Code'])
    df_code["Quantity"] = 1
    df_code["Quantity"] = df_code.groupby(["Code"]).transform(sum)
    df_code = df_code.drop_duplicates(subset=["Code"], keep="first").reset_index(drop=True)

    df_merge = pd.merge(df_code , df_artikel, on='Code', how='left')
    df_merge = df_merge[['Code','Artikel','Quantity','Size']]
    df_merge['Artikel'] = df_merge['Artikel'].fillna('TIDAK ADA DATA')
    df_merge['Size'] = df_merge['Size'].fillna('TIDAK ADA DATA')
    df_merge = df_merge.sort_values(by=['Artikel','Code'], ascending=True)
    dict = df_merge.to_dict('records')
    return dict

@app.route('/download', methods=['POST'])
def download():
    global parsed_data
    parsed_data = json.loads(request.data)
    
    current_datetime = datetime.now().strftime("%y%m%d_%H%M_%S")
    file_name = f"stockopname/Stock Opname {current_datetime}.xlsx"


    #simulation of code barcode
    df_artikel = pd.read_csv('data/database_artikel.csv')
    df_code = pd.DataFrame(parsed_data,columns=['Code'])
    df_code["Quantity"] = 1
    df_code["Quantity"] = df_code.groupby(["Code"]).transform(sum)
    df_code = df_code.drop_duplicates(subset=["Code"], keep="first").reset_index(drop=True)
    df_merge = pd.merge(df_code , df_artikel, on='Code', how='left')
    df_merge = df_merge[['Code','Artikel','Quantity','Size']]
    df_merge['Artikel'] = df_merge['Artikel'].fillna('TIDAK ADA DATA')
    df_merge['Size'] = df_merge['Size'].fillna('TIDAK ADA DATA')
    df_merge = df_merge.sort_values(by=['Artikel','Code'], ascending=True)
    df_merge.to_excel(file_name,index=False)

    return send_file(file_name, as_attachment=True)


@app.route('/updatedatabase', methods=['GET'])
def update():
    try:
        myscope = ['https://spreadsheets.google.com/feeds', 
                    'https://www.googleapis.com/auth/drive']
        mycred = ServiceAccountCredentials.from_json_keyfile_name('cred/credentiangoogle.json',myscope)

        client =gspread.authorize(mycred)
        mysheet = client.open("Database_Artikel").sheet1
        list_of_row = mysheet.get_all_records()
        df_artikel = pd.DataFrame(list_of_row)
        df_artikel = df_artikel[['Code','Artikel','Size']]
        df_artikel.to_csv('data/database_artikel.csv',index=False)
        status = 'success'
    except :
        status = 'failed'

    return status

if __name__ == '__main__':
    app.run()