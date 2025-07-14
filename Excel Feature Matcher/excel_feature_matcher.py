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
