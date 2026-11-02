# Shared reading code for Assignment 4.
# Every question folder calls this file, so all 12 programs
# look at the same data and treat a blank value the same way.

import json
from pathlib import Path


# This file sits in the Assignment 4 folder.
ROOT = Path(__file__).resolve().parent

# The three websites asked for in the PDF. Audible, BookBub, and Kobo are not used.
# review_folder follows the PDF: Openlibrary uses "reviews", the other two use "review".
# file_tag is the website name that a correct file name should contain.
SOURCES = [
    {
        "name": "Goodreads",
        "folder": "goodreads",
        "review_folder": "review",
        "file_tag": "Goodreads",
    },
    {
        "name": "Amazon",
        "folder": "Amazon",
        "review_folder": "review",
        "file_tag": "Amazon",
    },
    {
        "name": "OpenLibrary",
        "folder": "Openlibrary",
        "review_folder": "reviews",
        "file_tag": "OpenLibrary",
    },
]

# Question 4 names these six fields. The right-hand name is the key inside the JSON.
ASKED_FIELDS = [
    ("title", "title"),
    ("authors", "authors"),
    ("publisher", "publisher"),
    ("language", "language"),
    ("date of publication", "publication_date"),
    ("genre", "genres"),
]

# The scraper wrote these words when the website had no value.
# I count every one of them as missing. A real title or a real number is not missing.
EMPTY_WORDS = {"", "n/a", "na", "none", "null", "-", "unknown"}

# Remember a JSON file after the first read, so one program does not parse it twice.
_CACHE = {}


def is_missing(value):
    # None means the key was there but the value was null.
    if value is None:
        return True

    # A text value is missing when it is blank or one of the placeholder words.
    if isinstance(value, str):
        return value.strip().lower() in EMPTY_WORDS

    # An empty list or an empty dict is also nothing.
    if isinstance(value, (list, dict)) and len(value) == 0:
        return True

    return False


def source_folder(source):
    # data/goodreads, data/Amazon, or data/Openlibrary
    return ROOT / "data" / source["folder"]


def metadata_file(source):
    # Each metadata folder has one JSON file. The name has a space in it.
    folder = source_folder(source) / "metadata"
    found = sorted(folder.glob("*.json"))
    if len(found) != 1:
        raise FileNotFoundError("Expected one JSON file in " + str(folder))
    return found[0]


def load_metadata(source):
    # The JSON is one big dictionary. The key is the ISBN-13. The value is the book.
    name = source["name"]
    if name in _CACHE:
        return _CACHE[name]

    path = metadata_file(source)
    with path.open(encoding="utf-8") as handle:
        records = json.load(handle)

    _CACHE[name] = records
    return records


def kind_folder(source, kind):
    # kind is "coverpage", "blurb", or "review".
    # The review folder name is not the same word on every website.
    if kind == "review":
        return source_folder(source) / source["review_folder"]
    return source_folder(source) / kind


def list_files(source, kind):
    # File names only. Hidden files such as .gitkeep are not covers, blurbs, or reviews.
    folder = kind_folder(source, kind)
    names = []
    for path in folder.iterdir():
        if path.is_file() and not path.name.startswith("."):
            names.append(path.name)
    names.sort()
    return names


def isbn_from_filename(name):
    # A saved file starts with the 13-digit ISBN, for example 9780002171045_b_Amazon_1.txt.
    # ISBN-13 book numbers start with 978 or 979.
    if len(name) < 13:
        return None
    prefix = name[:13]
    if not prefix.isdigit():
        return None
    if not (prefix.startswith("978") or prefix.startswith("979")):
        return None
    return prefix


def index_files(names):
    # Group file names by the ISBN written in front of them.
    # A name with no ISBN is kept in the broken list.
    by_isbn = {}
    broken = []
    for name in names:
        isbn = isbn_from_filename(name)
        if isbn is None:
            broken.append(name)
            continue
        if isbn not in by_isbn:
            by_isbn[isbn] = []
        by_isbn[isbn].append(name)
    return by_isbn, broken


class Report:
    # Prints each line and keeps it, so the same text can be saved for the write-up.
    def __init__(self, folder):
        self.folder = Path(folder)
        self.folder.mkdir(parents=True, exist_ok=True)
        self.lines = []

    def say(self, text=""):
        # The Windows console cannot print every character that appears in a genre name.
        # The saved file still keeps the original text.
        try:
            print(text)
        except UnicodeEncodeError:
            print(text.encode("cp1252", "replace").decode("cp1252"))
        self.lines.append(text)

    def save(self, filename="output.txt"):
        path = self.folder / filename
        path.write_text("\n".join(self.lines) + "\n", encoding="utf-8")
        return path

    def save_lines(self, filename, rows):
        # One item per line. Used for long ISBN lists that should not flood the screen.
        path = self.folder / filename
        if len(rows) == 0:
            path.write_text("", encoding="utf-8")
        else:
            path.write_text("\n".join(rows) + "\n", encoding="utf-8")
        return path
