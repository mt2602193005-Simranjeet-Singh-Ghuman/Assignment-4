
import sys
from pathlib import Path

# Locate the project folder
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

import bookdata

# -------------------------------
# PART 1: Count unique ISBN numbers
# -------------------------------

combined = set()

for source in bookdata.SOURCES:

    records = bookdata.load_metadata(source)

    for isbn in records:
        combined.add(isbn)

print("Total unique ISBN:", len(combined))


# -------------------------------
# PART 2: Count blurb files
# -------------------------------

print("\nBlurb file counts:")

for source in bookdata.SOURCES:

    blurb_names = bookdata.list_files(source, "blurb")

    print(source["name"], ":", len(blurb_names))
