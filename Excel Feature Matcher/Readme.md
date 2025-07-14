**Excel Feature Matcher**

A Python script that updates the "Supported" column in a target Excel file by matching feature descriptions from a source Excel file using fuzzy string matching. Ideal when exact string matching (e.g. VLOOKUP) isn’t sufficient.

---

## 🚀 Features

- **Keeps original structure**: Preserves all columns, including serial numbers and section headings.
- **Fuzzy matching**: Uses RapidFuzz to match similar but not identical feature descriptions.
- **Configurable**: Column names are defined at the top of the script for easy customization.
- **No database needed**: Works directly on Excel files.

---

## 📋 Prerequisites

- Python 3.7+
- The following Python packages:
  - `pandas`
  - `openpyxl`
  - `rapidfuzz`

Install dependencies:

```bash
pip install pandas openpyxl rapidfuzz
```

---

## ⚙️ Configuration

At the top of the script, adjust the following variables to match your Excel file:

```python
import pandas as pd
from rapidfuzz import fuzz, process

# Configuration
SOURCE_FILE = "source.xlsx"
TARGET_FILE = "target.xlsx"
OUTPUT_FILE = "target_updated.xlsx"

# Column names in source (one feature key, two value columns)
SRC_FEATURE_COL     = "Feature_Key"
SRC_VALUE_COLS      = ["Provision_Value", "Additional_Info"]

# Column names in target
TGT_SERIAL_COL      = "Sr. No."
TGT_FEATURE_COL     = "Feature_Label"
TGT_OPTIONAL_COL    = "Optionality"
TGT_PROVISION_COLS  = ["Supported_Response", "Extra_Response"]

# Threshold for fuzzy matching (0–100)
MATCH_THRESHOLD     = 80

# Load Excel files
target_df = pd.read_excel(TARGET_FILE)
source_df = pd.read_excel(SOURCE_FILE)

```

- **Serial column** is preserved unmodified.
- Rows where the `Optionality` column is **not empty** are considered actual features to match; others (e.g., headings) are skipped.

---

## 📖 Usage

1. Place your `source.xlsx` and `target.xlsx` files in the same directory as the script (or update paths).
2. Run the script:
   ```bash
   python excel_feature_matcher.py
   ```
3. The updated file `target_updated.xlsx` will be generated with the `Supported_Response` column filled based on the closest fuzzy matches.

---

## 💡 How It Works

1. **Load Excel files** into pandas DataFrames.
2. **Forward-fill** the feature column in the target to propagate section headings if needed.
3. Build a **dictionary** from source feature keys to provision values.
4. For each target row marked as an actual feature:
   - Perform a **fuzzy match** against source keys.
   - If the match score ≥ `MATCH_THRESHOLD`, fill the `Supported_Response` with the corresponding source value.
   - Otherwise, leave the original cell unchanged.
5. **Save** the result to a new Excel file, preserving all other columns.

---

## 🛠️ Script (`excel_feature_matcher.py`)

```python
import pandas as pd
from rapidfuzz import fuzz, process

# Configuration
SOURCE_FILE = "source.xlsx"
TARGET_FILE = "target.xlsx"
OUTPUT_FILE = "target_updated.xlsx"

SRC_FEATURE_COL     = "Feature_Key"
SRC_VALUE_COL       = "Provision_Value"
TGT_SERIAL_COL      = "Sr. No."
TGT_FEATURE_COL     = "Feature_Label"
TGT_OPTIONAL_COL    = "Optionality"
TGT_PROVISION_COL   = "Supported_Response"

MATCH_THRESHOLD     = 80

# Load files
target_df = pd.read_excel(TARGET_FILE)
source_df = pd.read_excel(SOURCE_FILE)

# Clean headers
target_df.columns = [c.strip() for c in target_df.columns]
source_df.columns = [c.strip() for c in source_df.columns]

# Forward-fill section headings in target
if target_df[TGT_FEATURE_COL].isnull().any():
    target_df[TGT_FEATURE_COL] = target_df[TGT_FEATURE_COL].fillna(method='ffill')

# Build mapping from source
source_map = dict(zip(source_df[SRC_FEATURE_COL], source_df[SRC_VALUE_COL]))

# Matching function
def fuzzy_lookup(text):
    match = process.extractOne(text, source_map.keys(), scorer=fuzz.token_sort_ratio)
    if match and match[1] >= MATCH_THRESHOLD:
        return source_map[match[0]]
    return None

# Update supported column
updated = []
for _, row in target_df.iterrows():
    if pd.notna(row[TGT_OPTIONAL_COL]):  # actual feature row
        result = fuzzy_lookup(row[TGT_FEATURE_COL])
        updated.append(result or row[TGT_PROVISION_COL])
    else:
        updated.append(row[TGT_PROVISION_COL])  # keep headings blank or existing

target_df[TGT_PROVISION_COL] = updated

# Save output
target_df.to_excel(OUTPUT_FILE, index=False)
```
