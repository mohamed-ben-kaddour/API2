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
        response = supabase.functions.invoke("hello-world", invoke_options={'body':{}})
        print(response)
        if response.data:
            for row in response.data:
                print(row)
        else:
            print("No data returned.")
        
        # Use the data from the response
        data = response.data

        # Create an in-memory Excel file using openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Monthly Attendance"

        # Add column headers
        headers = ["Activity ID", "Month", "Male Count", "Female Count"]
        ws.append(headers)

        # Add data rows
        for row in data:
            ws.append([
                row.get("activity_id", "N/A"),
                row.get("month", "N/A"),
                row.get("male_count", "N/A"),
                row.get("female_count", "N/A")
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

# --- Older Code Version (for reference) ---
# If using an older version of the Supabase client that requires .execute(),
# you could call the RPC function like this:
#
# response_old = supabase.rpc("get_monthly_attendance_counts", params={}).execute()
# print("Older method response:")
# if response_old.data:
#     for row in response_old.data:
#         print(row)
# else:
#     print("No data returned.")
#
# And then use `response_old.data` when building your Excel file.

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))  # Default to 5000 if no PORT is provided
    app.run(host="0.0.0.0", port=port, debug=True)
