from flask import Flask, send_file
import pandas as pd
import io
from supabase import create_client, Client

# Initialize Flask app
app = Flask(__name__)

# Initialize Supabase client
url = "YOUR_SUPABASE_URL"  # Replace with your Supabase URL
key = "YOUR_SUPABASE_KEY"  # Replace with your Supabase API Key
supabase: Client = create_client(url, key)

@app.route('/download_excel')
def download_excel():
    # Query data from Supabase (replace with your actual query)
    data = supabase.table("your_table_name").select("*").execute()

    # Convert to a pandas DataFrame
    df = pd.DataFrame(data['data'])  # Extract the 'data' from the response

    # Create an in-memory Excel file
    excel_file = io.BytesIO()
    df.to_excel(excel_file, index=False, engine='openpyxl')
    excel_file.seek(0)
    
    # Send the file as a response
    return send_file(excel_file, as_attachment=True, download_name='data.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    app.run(debug=True)
