A Python script that updates the "Supported" column in a target Excel file by matching feature descriptions from a source Excel file using fuzzy string matching. Ideal when exact string matching (e.g. VLOOKUP) isn’t sufficient.

🚀 Features

Keeps original structure: Preserves all columns, including serial numbers and section headings.

Fuzzy matching: Uses RapidFuzz to match similar but not identical feature descriptions.

Configurable: Column names are defined at the top of the script for easy customization.

No database needed: Works directly on Excel files.

📋 Prerequisites

Python 3.7+

The following Python packages:

pandas

openpyxl

rapidfuzz

Install dependencies:
