from flask import Flask, send_file
import openpyxl
import io
from supabase import create_client, Client
import os

# Initialize Flask app
app = Flask(__name__)

# Initialize Supabase client using environment variables
url = os.getenv("SUPABASE_URL")
key = os.getenv("SUPABASE_KEY")
supabase: Client = create_client(url, key)

@app.route('/download_excel')
def download_excel():
    # Query data from Supabase (replace with your actual table and query)
    data = supabase.table("activity").select("*").execute()

    # Create an in-memory Excel file using openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active

    # Add column headers (assuming data is a list of dictionaries)
    headers = data['data'][0].keys()  # Get the column names from the first row
    ws.append(list(headers))

    # Add data rows
    for row in data['data']:
        ws.append(list(row.values()))

    # Save to an in-memory file
    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)

    # Send the file as a response
    return send_file(excel_file, as_attachment=True, download_name='data.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    # Use the port from the environment variable
    port = int(os.getenv("PORT", 5000))  # Default to 5000 if no PORT is provided
    app.run(host="0.0.0.0", port=port, debug=True)
