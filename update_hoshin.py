import re
import json

with open("index_hoshin_matrice_XMatrix4_mise_a_jour.html", "r", encoding="utf-8") as f:
    content = f.read()

# Extract HOSHIN_DATABASE
db_match = re.search(r"const HOSHIN_DATABASE = (\[.*?\]);", content, re.DOTALL)
if db_match:
    db_str = db_match.group(1)
else:
    print("Could not find HOSHIN_DATABASE")
    exit(1)

# Extract HOSHIN_PROCESS_BLUEPRINT
pb_match = re.search(r"const HOSHIN_PROCESS_BLUEPRINT = (\[.*?\]);", content, re.DOTALL)
if pb_match:
    pb_str = pb_match.group(1)
else:
    print("Could not find HOSHIN_PROCESS_BLUEPRINT")
    exit(1)

import pyjson5

try:
    db = pyjson5.decode(db_str)
    pb = pyjson5.decode(pb_str)
    
    with open("data/hoshin_content.json", "w", encoding="utf-8") as f:
        json.dump({"leaders": db, "processBlueprint": pb}, f, indent=2, ensure_ascii=False)
    print("Successfully updated data/hoshin_content.json")
except Exception as e:
    print(f"Error parsing JSON: {e}")

