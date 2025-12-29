# Inspection_Form_Generator
Automate yearly inspection form generation from Word templates. This tool intelligently extracts project details, detects types (SAFE, RGI, etc.), and creates monthly files with randomized Mon-Fri dates. Features smart validation, conflict resolution, and auto-saving to Downloads.
# Universal Inspection Form Generator

A robust Python automation tool designed to generate yearly batches of Inspection Forms based on a Word Document (`.docx`) template. It automatically randomizes dates (Mon-Fri), preserves document formatting, and intelligently extracts project details from the template itself.

## 🚀 Features

* **Smart Template Detection:** Automatically detects the inspection type (e.g., SAFE, RGI) based on the filename.
* **Auto-Extraction:** Scans the uploaded template to find "Location", "Project No", "Inspector", etc., eliminating manual data entry.
* **Intelligent Validation:** Prevents errors by checking if the requested inspection type matches the uploaded file type.
* **Randomized Dates:** Generates valid dates (Monday-Friday only) for every month of the year.
* **Batch Generation:** Creates 12 months' worth of forms in seconds.
    * *SAFE Type:* Generates 2 forms per month (bi-weekly logic).
    * *Other Types:* Generates 1 form per month.
* **Continuous Workflow:** After finishing a job, allows the user to immediately start a new batch without restarting the program.

## 🛠️ Prerequisites & Installation

### 1. Execute EXE
Download InspectionFormGenerator.exe and double-click it to execute it.
