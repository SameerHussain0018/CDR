from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
from io import BytesIO, StringIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    if file.filename == '':
        return "No selected file"

    try:
        # Read Excel file into DataFrame
        df = pd.read_excel(file)

        # Convert DataFrame to CSV
        csv_data = df.to_csv(index=False)

        # Create a BytesIO object to store the CSV data
        csv_io = BytesIO()
        csv_io.write(csv_data.encode())

        # Set the BytesIO object to the beginning
        csv_io.seek(0)

        return send_file(csv_io, as_attachment=True, download_name='converted_file.csv', mimetype='text/csv')

    except Exception as e:
        return f"Error: {str(e)}"

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return redirect(request.url)

    file = request.files['file']

    if file.filename == '':
        return redirect(request.url)

    if file:
        # Read CSV file and perform the operations
        df = process_csv(file)

        # Save the modified DataFrame to a new CSV file
        output_file_path = 'output_file.csv'
        df.to_csv(output_file_path, index=False)

        # Create a BytesIO object to store the Excel file
        excel_io = BytesIO()
        df.to_excel(excel_io, index=False)
        excel_io.seek(0)
        
        # Send the BytesIO object as an Excel file
        return send_file(
            excel_io,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='modified_data.xlsx'
        )
def download_excel():
    file_path = request.args.get('file', '')
    
    if not file_path:
        return "File path not provided", 400
    # Create a BytesIO object to store the Excel file
    excel_data = process_csv(request.files['file']).to_excel(index=False)
    excel_io = StringIO()
    excel_io.write(excel_data)
    excel_io.seek(0)

    # Send the BytesIO object as an Excel file
    return send_file(
        excel_io,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='modified_data.xlsx'
    )
def process_csv(file):
    # Read the CSV file
    file_content = file.stream.read().decode('utf-8')
    df = pd.read_csv(StringIO(file_content))

    # Convert 'Data Volume (KB)' column to numeric and divide by 1000
    df['Data Volume (KB)'] = pd.to_numeric(df['Data Volume (KB)'].str.replace(',', ''), errors='coerce') / 1000

    # Create a new column with the converted values
    df['Data Volume (MB)'] = df['Data Volume (KB)'].round(3)  # Round to three decimal places

    # Round off the 'Data Volume (KB)' column and ensure it's at least 1
    df['Rounded Data Volume'] = df['Data Volume (KB)'].round().clip(lower=1)

    return df

if __name__ == '__main__':
    app.run(debug=True)
