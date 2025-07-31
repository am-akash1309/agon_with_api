import requests

def read_timesheet_data(filename: str) -> str:
    """
    Use this to read invoice data from an Excel timesheet file (XLSX format). 
    It returns rows of date, status, and remarks. Input must be the filename (e.g., 'timesheet_july.xlsx').
    """
    try:
        response = requests.get(
            "http://localhost:5000/read_timesheet",
            params={"filename": filename}
        )
        data = response.json()
        return data.get("result") or data.get("error")
    except Exception as e:
        return f"Error calling read_timesheet_data API: {e}"

def create_invoice_document(filename: str, data: dict) -> str:
    """    
    Generates a .docx invoice from a dictionary of data.    
    Use this to generate and save a formatted invoice as a Word (.docx) file. 
    Input must be a dictionary with 'filename' and 'data' keys.
    """
    try:
        response = requests.post(
            "http://localhost:5000/create_invoice",
            json={
                "filename": filename,
                "data": data
            }
        )
        data = response.json()
        return data.get("result") or data.get("error")
    except Exception as e:
        return f"Error calling create_invoice_document API: {e}"

def save_or_update_timesheet(filename: str, date: str, status: str, remarks: str) -> str:
    """
    Saves or updates a single entry in the Excel timesheet.
    If an entry for the given date already exists, it will be UPDATED.
    Otherwise, a new entry will be ADDED.
    The date must be in 'YYYY-MM-DD' format.
    It needs the filename, date, status, and remarks as input. Filename usually be in the format: 'timesheet_<month>.xlsx').
    """
    try:
        response = requests.post(
            "http://localhost:5000/save_or_update_timesheet",
            json={
                "filename": filename,
                "date": date,
                "status": status,
                "remarks": remarks
            }
        )
        data = response.json()
        return data.get("result") or data.get("error")
    except Exception as e:
        return f"Error calling save_or_update_timesheet API: {e}"

def send_message_with_attachments(xlsx_filename: str, docx_filename: str) -> str:
    """
    Sends an email with an XLSX and a DOCX file as attachments from the current directory.
    
    Args:
        xlsx_filename: The filename of the .xlsx file (e.g., 'timesheet_july.xlsx').
        docx_filename: The filename of the .docx file (e.g., 'invoice_july.docx').
        
    Returns:
        A string indicating success or failure.
    """
    try:
        response = requests.post(
            "http://localhost:5000/send_telegram",
            json={
                "xlsx_filename": xlsx_filename,
                "docx_filename": docx_filename
            }
        )
        data = response.json()
        return data.get("result") or data.get("error")
    except Exception as e:
        return f"Error calling send_message_with_attachments API: {e}"

def calculate_salary(present_days: int, pay_per_day: int) -> str:
    try:
        response = requests.get(
            "http://localhost:5000/calculate_salary",
            params={"present_days": present_days, "pay_per_day": pay_per_day}
        )
        data = response.json()
        if "salary" in data:
            return f"Present Days: {data['present_days']}, Pay Per Day: {data['pay_per_day']}, Salary: ₹{data['salary']}"
        else:
            return f"Error: {data.get('error', 'Unknown error')}"
    except Exception as e:
        return f"Error calling calculate_salary API: {e}"