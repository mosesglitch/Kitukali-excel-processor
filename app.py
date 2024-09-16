from flask import Flask, request, send_file, jsonify
import pandas as pd
import io
from flask_cors import CORS
import os
app = Flask(__name__)
CORS(app)

# def process_and_combine_excel(workbook_path):
#     # Read the Excel workbook
#     order_df = pd.read_excel(workbook_path, sheet_name='Orders')
#     pesapal_df = pd.read_excel(workbook_path, sheet_name='pesapal Export')

#     # Process 'Location' column in the 'Order' dataset
#     if 'Location' in order_df.columns:
#         order_df['Location'] = order_df['Location'].apply(lambda x: x.replace('Kitu Kali', '') if pd.notna(x) else x)

#     # Process 'Description' column in the 'pesapal Export' dataset
#     if 'Description' in pesapal_df.columns:
#         pesapal_df['Description'] = pesapal_df['Description'].apply(lambda x: x.split(',')[0] if pd.notna(x) else x).apply(lambda x: x.split(':')[1] if pd.notna(x) else x)
    
#     # Select and rename columns from each dataset according to the specified mapping
#     order_df = order_df.rename(columns={
#         'Created at': 'Date',
#         'Name': 'Code',
#         'Payment Method': 'Payment Method',
#         'Location': 'Location',
#         'Total': 'Total'
#     })
# def process_and_combine_excel(workbook_path):
#     # Read the Excel workbook
#     order_df = pd.read_excel(workbook_path, sheet_name='Orders')
#     pesapal_df = pd.read_excel(workbook_path, sheet_name='pesapal Export')

#     # Process 'Location' column in the 'Order' dataset
#     if 'Location' in order_df.columns:
#         order_df['Location'] = order_df['Location'].apply(lambda x: x.replace('Kitu Kali', '') if pd.notna(x) else x)

#     # Process 'Description' column in the 'pesapal Export' dataset
#     if 'Description' in pesapal_df.columns:
#         pesapal_df['Description'] = pesapal_df['Description'].apply(lambda x: x.split(',')[0] if pd.notna(x) else x).apply(lambda x: x.split(':')[1] if pd.notna(x) else x)
    
#     # Select and rename columns from each dataset according to the specified mapping
#     order_df = order_df.rename(columns={
#         'Created at': 'Date',
#         'Name': 'Code',
#         'Payment Method': 'Payment Method',
#         'Location': 'Location',
#         'Total': 'Total'
#     })

#     pesapal_df = pesapal_df.rename(columns={
#         'Confirmation Code': 'Code',
#         'Amount': 'Total',
#         'Date': 'Date',
#         'Payment Method': 'Payment Method',
#         'Description': 'Location'
#     })

#     # Ensure that columns exist before combining
#     combined_columns = ['Date', 'Code', 'Payment Method', 'Location', 'Total']

#     if set(combined_columns).issubset(order_df.columns) and set(combined_columns).issubset(pesapal_df.columns):
#         # Select required columns
#         order_df = order_df[combined_columns]
#         pesapal_df = pesapal_df[combined_columns]
        
#         # Combine the two datasets
#         combined_df = pd.concat([order_df, pesapal_df], ignore_index=True)

#         # Save the combined dataset to a BytesIO object
#         output = io.BytesIO()
#         combined_df.to_excel(output, index=False)
#         output.seek(0)

#         return output
#     else:
#         return None 
  
# import pandas as pd
# import io

def process_and_combine_excel(workbook_path):
    # Read the entire workbook
    xls = pd.ExcelFile(workbook_path)
    combined_df = pd.DataFrame()

    # Define the column mappings and required columns
    order_columns = ['Created at', 'Name', 'Payment Method', 'Location', 'Total']
    pesapal_columns = ['Confirmation Code', 'Amount', 'Date', 'Payment Method', 'Description']
    combined_columns = ['Date', 'Code', 'Payment Method', 'Location', 'Total']
    
    order_col_mapping = {
        'Created at': 'Date',
        'Name': 'Code',
        'Payment Method': 'Payment Method',
        'Location': 'Location',
        'Total': 'Total'
    }

    pesapal_col_mapping = {
        'Confirmation Code': 'Code',
        'Amount': 'Total',
        'Date': 'Date',
        'Payment Method': 'Payment Method',
        'Description': 'Location'
    }

    # Iterate through all sheets in the workbook
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)

        if set(order_columns).issubset(df.columns):
            # Process 'Location' column in the order dataset
            df['Location'] = df['Location'].apply(lambda x: x.replace('Kitu Kali', '') if pd.notna(x) else x)
            df['Location'] = df['Location'].apply(lambda x: x.strip() if pd.notna(x) else x)
            
            # Rename columns according to the order mapping
            df = df.rename(columns=order_col_mapping)
            df = df[combined_columns]
            
        elif set(pesapal_columns).issubset(df.columns):
            # Process 'Description' column in the pesapal dataset
            df['Description'] = df['Description'].apply(lambda x: x.split(',')[0] if pd.notna(x) else x).apply(lambda x: x.split(':')[1] if pd.notna(x) else x)
            df['Description'] = df['Description'].apply(lambda x: x.strip() if pd.notna(x) else x)
            
            # Rename columns according to the pesapal mapping
            df = df.rename(columns=pesapal_col_mapping)
            df = df[combined_columns]
            
        else:
            # If the sheet does not contain the relevant columns, skip it
            continue
        
        # Append the processed dataframe to the combined dataframe
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    if not combined_df.empty:
        # Save the combined dataset to a BytesIO object
        output = io.BytesIO()
        combined_df.to_excel(output, index=False)
        output.seek(0)
        return output
    else:
        return None

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if file and file.filename.endswith('.xlsx'):
        output = process_and_combine_excel(file)
        
        if output:
            return send_file(output, download_name='Combined_Output.xlsx', as_attachment=True)
        else:
            return jsonify({"error": "Required columns are missing in the uploaded file"}), 400
    else:
        return jsonify({"error": "Invalid file type. Please upload an Excel file."}), 400

if __name__=='__main__':
     app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)),debug=True)
