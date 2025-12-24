import os
import random
import calendar
import shutil
import tkinter as tk
from tkinter import filedialog
from datetime import date, timedelta
from docx import Document

# ==========================================
# 1. CONFIGURATION
# ==========================================
class Config:
    VALID_INSP_TYPES = ["RGI", "SAFE", "WSIN", "ENVI"]
    
    KEY_MAPPING = {
        "PI Inspection Form": "location",
        "Project No": "project_no",
        "Inspector": "inspector",
        "Contractor": "contractor",
        "Form No": "form_no",
        "Insp Type": "insp_type",
        "Scheduled": "scheduled",
        "Deadline": "deadline",
        "Performed By": "inspector",   
        "Checked By": "checker",
        "Date": "perform_date",         
    }

# ==========================================
# 2. DOCUMENT UTILITIES (Static Helpers)
# ==========================================
class DocUtils:
    """Static helper methods for low-level Docx manipulation."""
    
    @staticmethod
    def find_next_real_cell(row_cells, current_index):
        """Skips merged cells to find the actual value cell."""
        current_cell = row_cells[current_index]
        next_index = current_index + 1
        while next_index < len(row_cells):
            next_cell = row_cells[next_index]
            if next_cell._element is not current_cell._element:
                return next_cell
            next_index += 1
        return None

    @staticmethod
    def safe_update_cell(cell, new_value):
        """Updates cell text preserving format."""
        if not cell: return
        if cell.paragraphs:
            p = cell.paragraphs[0]
            if p.runs:
                p.runs[0].text = str(new_value)
                for run in p.runs[1:]:
                    run.text = ""
            else:
                p.add_run(str(new_value))
        else:
            cell.text = str(new_value)

# ==========================================
# 3. DATE LOGIC
# ==========================================
class DateEngine:
    """Handles calendar calculations."""
    
    @staticmethod
    def get_random_weekday(year, month, start_day, end_day):
        _, last_day_of_month = calendar.monthrange(year, month)
        actual_start = max(1, start_day)
        actual_end = min(end_day, last_day_of_month)
        
        if actual_start > actual_end:
            return date(year, month, actual_end)

        valid_dates = []
        for day in range(actual_start, actual_end + 1):
            try:
                curr_date = date(year, month, day)
                if curr_date.weekday() < 5: # Mon-Fri
                    valid_dates.append(curr_date)
            except ValueError:
                continue
                
        if valid_dates:
            return random.choice(valid_dates)
        return date(year, month, actual_start)

# ==========================================
# 4. TEMPLATE MODEL
# ==========================================
class InspectionTemplate:
    """Represents the uploaded Word document."""
    
    def __init__(self, filepath):
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        self.type = self._detect_type()
        self.doc_obj = Document(filepath) # Keep a reference for scanning
        self.project_details = self._extract_details()

    def _detect_type(self):
        fname_upper = self.filename.upper()
        for type_code in Config.VALID_INSP_TYPES:
            if type_code in fname_upper:
                return type_code
        return "UNKNOWN"

    def _extract_details(self):
        extracted = {
            "location": "Unknown", "project_no": "Unknown",
            "inspector": "Unknown", "contractor": "Unknown", "checker": "Unknown"
        }
        keys = {
            "PI Inspection Form": "location", "Location": "location",           
            "Project No": "project_no", "Inspector": "inspector",
            "Contractor": "contractor", "Checked By": "checker"
        }
        
        print("\n>> Scanning template for project details...")
        for table in self.doc_obj.tables:
            for row in table.rows:
                cells = row.cells
                for i, cell in enumerate(cells):
                    cell_text = cell.text.strip()
                    for key, field in keys.items():
                        if key in cell_text:
                            if key == "Inspector" and "PI Inspection" in cell_text: continue
                            
                            target = DocUtils.find_next_real_cell(cells, i)
                            if target and target.text.strip():
                                extracted[field] = target.text.strip()
                                print(f"   Found {field.upper()}: {extracted[field]}")
        return extracted

# ==========================================
# 5. GENERATOR ENGINE
# ==========================================
class FormGenerator:
    """Handles the creation of batch files."""
    
    def __init__(self, template_model):
        self.template = template_model

    def generate_batch(self, req_type, year):
        # Setup Output Directory
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        out_name = f"{req_type}_{year}_Forms_Generated"
        out_dir = os.path.join(downloads, out_name)

        if os.path.exists(out_dir): shutil.rmtree(out_dir)
        os.makedirs(out_dir)

        print(f"\n>> Generating {req_type} forms in: {out_dir}")
        
        counter = 1
        for month in range(1, 13):
            # Calculate Dates
            dates = self._calculate_dates_for_month(req_type, year, month)
            
            # Common Data
            month_str = f"{month:02d}"
            _, last_day = calendar.monthrange(year, month)
            
            base_data = self.template.project_details.copy()
            base_data["scheduled"] = f"01/{month_str}/{year}"
            base_data["deadline"] = f"{last_day}/{month_str}/{year}"
            base_data["insp_type"] = req_type

            # Generate Files
            for perform_dt in dates:
                date_str = perform_dt.strftime("%d/%m/%Y")
                form_no = f"IPRJ{req_type}{counter:04d}"
                
                final_data = base_data.copy()
                final_data["form_no"] = form_no
                final_data["perform_date"] = date_str

                self._create_single_file(final_data, out_dir, form_no)
                print(f"Created: {form_no}.docx | {date_str}")
                counter += 1

        print(f"\nSUCCESS: Files saved to Downloads folder.")
        os.startfile(out_dir)

    def _calculate_dates_for_month(self, req_type, year, month):
        _, last_day = calendar.monthrange(year, month)
        if req_type == "SAFE":
            d1 = DateEngine.get_random_weekday(year, month, 1, 7)
            min_gap = d1 + timedelta(days=14)
            d2 = DateEngine.get_random_weekday(year, month, min_gap.day, last_day)
            return [d1, d2]
        else:
            return [DateEngine.get_random_weekday(year, month, 1, last_day)]

    def _create_single_file(self, data, out_dir, fname):
        # Open a fresh copy of the doc for every iteration
        doc = Document(self.template.filepath)
        
        for table in doc.tables:
            for row in table.rows:
                cells = row.cells
                for i, cell in enumerate(cells):
                    text = cell.text.strip()
                    for key, field in Config.KEY_MAPPING.items():
                        if key in text:
                            if key == "Inspector" and "PI Inspection" in text: continue
                            target = DocUtils.find_next_real_cell(cells, i)
                            if target:
                                DocUtils.safe_update_cell(target, data[field])
        
        doc.save(os.path.join(out_dir, f"{fname}.docx"))

# ==========================================
# 6. USER INTERFACE
# ==========================================
class UserInterface:
    """Handles inputs, outputs, and dialogs."""
    
    @staticmethod
    def open_file_dialog():
        root = tk.Tk()
        root.withdraw() 
        root.attributes('-topmost', True) 
        fn = filedialog.askopenfilename(
            title="Select Inspection Form Template",
            filetypes=[("Word Documents", "*.docx")]
        )
        root.destroy()
        return fn

    @staticmethod
    def ask_type_and_year(default_type):
        while True:
            val = input(f"\nWhich Insp Type and year? (e.g., {default_type} 2025) [Press Enter for default]: ").strip()
            if not val:
                print(f">> Using default: {default_type} 2025")
                return default_type, 2025
            
            try:
                parts = val.split()
                if len(parts) < 2: raise ValueError
                
                req_type = parts[0].upper()
                year = int(parts[1])
                
                if not (1990 < year < 2100):
                    print(">> [!] Error: Realistic year required.")
                    continue
                return req_type, year
            except ValueError:
                print(">> [!] Invalid format. Try 'SAFE 2025'.")

    @staticmethod
    def ask_conflict_resolution(uploaded, requested):
        print(f"\n[!] CONFLICT: Uploaded [{uploaded}] vs Requested [{requested}]")
        while True:
            print("1. Re-upload correct file")
            print("2. Change request")
            choice = input("Enter 1 or 2: ").strip()
            if choice in ['1', '2']: return choice
            print("Invalid input.")

    @staticmethod
    def ask_next_step():
        print("\n" + "="*50)
        print("JOB COMPLETE. What next?")
        print("1. Generate another")
        print("2. Exit")
        return input("Enter 1 or 2: ").strip()

# ==========================================
# 7. MAIN CONTROLLER
# ==========================================
class Application:
    """Orchestrates the program flow."""
    
    def __init__(self):
        self.ui = UserInterface()
        self.current_template = None

    def _acquire_template(self):
        """Loop until valid template is loaded."""
        print("\nSTEP 1: Select your Template")
        while True:
            path = self.ui.open_file_dialog()
            if not path:
                print(">> No file selected. Exiting.")
                return False
            
            template = InspectionTemplate(path)
            
            if template.type == "UNKNOWN":
                print(f"\n>> [!] ERROR: Filename must contain {Config.VALID_INSP_TYPES}")
                input("   Press Enter to retry...")
                continue
            
            print(f">> Loaded '{template.filename}' ([{template.type}])")
            self.current_template = template
            return True

    def run(self):
        print("--- Universal Inspection Form Generator (OOP Version) ---")
        
        while True:
            # 1. Get Template
            if not self._acquire_template():
                break

            # 2. Get Request & Validate
            while True:
                req_type, year = self.ui.ask_type_and_year(self.current_template.type)
                
                if req_type != self.current_template.type:
                    choice = self.ui.ask_conflict_resolution(self.current_template.type, req_type)
                    if choice == '1': # Re-upload
                        if self._acquire_template():
                            continue # Check inputs again with new template
                        else:
                            return # Exit if they cancel re-upload
                    elif choice == '2': # Change request
                        continue
                
                # If we get here, everything is valid
                break

            # 3. Generate
            generator = FormGenerator(self.current_template)
            generator.generate_batch(req_type, year)

            # 4. Loop or Exit
            if self.ui.ask_next_step() != '1':
                print(">> Exiting. Goodbye!")
                break
            else:
                print("\n>> Restarting...")

# ==========================================
# ENTRY POINT
# ==========================================
if __name__ == "__main__":
    app = Application()
    app.run()
