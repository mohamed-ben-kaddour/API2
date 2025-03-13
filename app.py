from flask import Flask, send_file, jsonify
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
    try:
        # Call the stored procedure (make sure it exists in Supabase)
        response = supabase.rpc('get_monthly_attendance', {}).execute()
        
        if 'data' not in response:
            return jsonify({"error": "Invalid response format from Supabase"}), 500
        
        data = response['data']
        
        if not data:
            return jsonify({"error": "No data found"}), 404

        # Create an in-memory Excel file using openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Monthly Attendance"

        # Add column headers
        headers = ["Activity Name", "Activity ID", "Month", "Gender", "Total Attendance"]
        ws.append(headers)

        # Add data rows
        for row in data:
            ws.append([
                row.get("activity_name", "N/A"),
                row.get("activity_id", "N/A"),
                row.get("year_month", "N/A"),
                row.get("sexe", "N/A"),
                row.get("total_attendance", "N/A")
            ])

        # Save to an in-memory file
        excel_file = io.BytesIO()
        wb.save(excel_file)
        excel_file.seek(0)

        # Send the file as a response
        return send_file(
            excel_file,
            as_attachment=True,
            download_name='attendance_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))  # Default to 5000 if no PORT is provided
    app.run(host="0.0.0.0", port=port, debug=True)
