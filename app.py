from flask import Flask, jsonify, request, render_template
from flask_cors import CORS
import pandas as pd
import re

app = Flask(__name__)
CORS(app)

# Load Excel file
file_path = r'C:\Users\MySurface\Downloads\college-chatbot\college-chatbot\EMPLOYEE MANAGEMENT .xlsx'

try:
    # Load all required sheets
    employee_df = pd.read_excel(file_path, sheet_name='PERSONAL_INFO')
    performance_df = pd.read_excel(file_path, sheet_name='PERFORMANCE')
    project_df = pd.read_excel(file_path, sheet_name='PROJECT_ASSIGNMENTS')      # New sheet
    schedule_df = pd.read_excel(file_path, sheet_name='EMPLOYEE_SCHEDULE')       # New sheet

    # Standardize column names across all sheets
    all_dfs = [employee_df, performance_df, project_df, schedule_df]
    for df in all_dfs:
        df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')

except FileNotFoundError:
    print("❌ Excel file not found. Please check the path.")
    exit()

@app.route('/')
def home():
    return render_template('index.html')

# Utility extractor
def extract_empid_from_question(question):
    match = re.search(r'\b\d{4,}\b', question)  # assuming EMPLOYEE_ID has at least 4 digits
    return match.group(0) if match else None

# ---------------------- DATA FETCH FUNCTIONS ----------------------

def get_employee_info(emp_id):
    emp = employee_df[employee_df['EMPLOYEE_ID'].astype(str) == str(emp_id)]
    if not emp.empty:
        row = emp.iloc[0]
        return {
            "Employee Name": row['NAME_OF_EMPLOYEE'],
            "EMPLOYEE_ID": emp_id,
            "Employee Info": row.drop(labels=['EMPLOYEE_ID']).to_dict()
        }
    return {"response": "Employee not found for the provided EMPLOYEE_ID."}

def get_employee_performance(emp_id):
    perf = performance_df[performance_df['EMPLOYEE_ID'].astype(str) == str(emp_id)]
    if not perf.empty:
        row = perf.iloc[0]
        full_perf = row.drop(labels=['EMPLOYEE_ID', 'NAME_OF_EMPLOYEE']).to_dict()
        return {
            "Employee Name": row['NAME_OF_EMPLOYEE'],
            "EMPLOYEE_ID": emp_id,
            "Performance": full_perf
        }
    return {"response": "Performance not found for the provided EMPLOYEE_ID."}

def get_employee_projects(emp_id):
    projects = project_df[project_df['EMPLOYEE_ID'].astype(str) == str(emp_id)]
    if not projects.empty:
        full_projects = projects.drop(columns=['EMPLOYEE_ID', 'NAME_OF_EMPLOYEE']).to_dict(orient='records')
        return {
            "Employee Name": projects.iloc[0]['NAME_OF_EMPLOYEE'],
            "EMPLOYEE_ID": emp_id,
            "Projects": full_projects
        }
    return {"response": "No projects found for the provided EMPLOYEE_ID."}

def get_employee_schedule(emp_id):
    schedule = schedule_df[schedule_df['EMPLOYEE_ID'].astype(str) == str(emp_id)]
    if not schedule.empty:
        full_schedule = schedule.drop(columns=['EMPLOYEE_ID', 'NAME_OF_EMPLOYEE']).to_dict(orient='records')
        return {
            "Employee Name": schedule.iloc[0]['NAME_OF_EMPLOYEE'],
            "EMPLOYEE_ID": emp_id,
            "Schedule": full_schedule
        }
    return {"response": "No schedule found for the provided EMPLOYEE_ID."}

# ---------------------- RESPONSE FORMATTER ----------------------

def format_response_data(response):
    if isinstance(response, dict):
        if any(isinstance(v, dict) or isinstance(v, list) for v in response.values()):
            formatted = ""
            for key, val in response.items():
                if isinstance(val, dict):
                    formatted += f"\n{key}:\n"
                    for sub_key, sub_val in val.items():
                        formatted += f"  {sub_key}: {sub_val}\n"
                elif isinstance(val, list):
                    formatted += f"\n{key}:\n"
                    for idx, record in enumerate(val, start=1):
                        formatted += f"  Record {idx}:\n"
                        for sub_key, sub_val in record.items():
                            formatted += f"    {sub_key}: {sub_val}\n"
                else:
                    formatted += f"{key}: {val}\n"
            return formatted.strip()
        else:
            return '\n'.join([f"{k}: {v}" for k, v in response.items()])
    elif isinstance(response, str):
        return response
    else:
        return str(response)

# ---------------------- CHAT ENDPOINT ----------------------

@app.route('/ask', methods=['POST'])
def ask():
    data = request.get_json()
    question = data.get('question', '').lower()
    emp_id = extract_empid_from_question(question)

    if emp_id:
        if "performance" in question or "attendance" in question:
            response = get_employee_performance(emp_id)
        elif "project" in question:
            response = get_employee_projects(emp_id)
        elif "schedule" in question or "work timing" in question or "shift" in question:
            response = get_employee_schedule(emp_id)
        elif "employee" in question or "personal info" in question:
            response = get_employee_info(emp_id)
        else:
            response = {"response": "Please ask about Employee info, Performance, Projects, or Schedule."}
    else:
        response = {"response": "Sorry, I couldn't understand your question."}

    return jsonify({'response': format_response_data(response)})

# ---------------------- MAIN ----------------------

if __name__ == '__main__':
    app.run(debug=True)
