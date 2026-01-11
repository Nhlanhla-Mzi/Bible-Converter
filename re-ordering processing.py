import pythonbible as bible
import re

# 1. Define your custom book order
CUSTOM_ORDER = [
    "Genesis", "Exodus", "Leviticus", "Numbers", "Deuteronomy",
    "Joshua", "Judges", "1 Samuel", "2 Samuel", "1 Kings", "2 Kings",
    "Isaiah", "Jeremiah", "Ezekiel", "Hosea", "Joel", "Amos",
    "Obadiah", "Jonah", "Micah", "Nahum", "Habakkuk", "Zephaniah",
    "Haggai", "Zechariah", "Malachi", "Psalms", "Proverbs", "Job",
    "Song of Solomon", "Ruth", "Lamentations", "Ecclesiastes",
    "Esther", "Daniel", "Ezra", "Nehemiah", "1 Chronicles", "2 Chronicles",
    "Matthew", "Mark", "Luke", "John", "Acts", "James", "1 Peter",
    "2 Peter", "1 John", "2 John", "3 John", "Jude", "Romans",
    "1 Corinthians", "2 Corinthians", "Galatians", "Ephesians",
    "Philippians", "Colossians", "1 Thessalonians", "2 Thessalonians",
    "Hebrews", "1 Timothy", "2 Timothy", "Titus", "Philemon", "Revelation"
]

# 2. Map common book name variations to a standard name.
BOOK_NAME_MAP = {
    "Gen": "Genesis",
    "Exo": "Exodus",
    "Lev": "Leviticus",
    "Num": "Numbers",
    "Deu": "Deuteronomy",
    "Jos": "Joshua",
    "Jdg": "Judges",
    "1Sam": "1 Samuel",
    "2Sam": "2 Samuel",
    "1Ki": "1 Kings",
    "2Ki": "2 Kings",
    "Isa": "Isaiah",
    "Jer": "Jeremiah",
    "Eze": "Ezekiel",
    "Hos": "Hosea",
    "Joe": "Joel",
    "Amo": "Amos",
    "Oba": "Obadiah",
    "Jon": "Jonah",
    "Mic": "Micah",
    "Nah": "Nahum",
    "Hab": "Habakkuk",
    "Zep": "Zephaniah",
    "Hag": "Haggai",
    "Zec": "Zechariah",
    "Mal": "Malachi",
    "Psa": "Psalms", "Psalm": "Psalms",
    "Pro": "Proverbs",
    "Job": "Job",
    "Son": "Song of Solomon",
    "Rut": "Ruth",
    "Lam": "Lamentations",
    "Ecc": "Ecclesiastes",
    "Est": "Esther",
    "Dan": "Daniel",
    "Ezr": "Ezra",
    "Neh": "Nehemiah",
    "1Ch": "1 Chronicles",
    "2Ch": "2 Chronicles",
    "Mat": "Matthew", "Matt": "Matthew",
    "Mar": "Mark",
    "Luk": "Luke",
    "Joh": "John",
    "Act": "Acts",
    "Jam": "James", "Jas": "James",
    "1Pe": "1 Peter", "1Pet": "1 Peter",
    "2Pe": "2 Peter", "2Pet": "2 Peter",
    "1Jo": "1 John", "1Joh": "1 John",
    "2Jo": "2 John", "2Joh": "2 John",
    "3Jo": "3 John", "3Joh": "3 John",
    "Jud": "Jude",
    "Rom": "Romans",
    "1Co": "1 Corinthians", "1Cor": "1 Corinthians",
    "2Co": "2 Corinthians", "2Cor": "2 Corinthians",
    "Gal": "Galatians",
    "Eph": "Ephesians",
    "Php": "Philippians", "Phil": "Philippians",
    "Col": "Colossians",
    "1Th": "1 Thessalonians", "1Thes": "1 Thessalonians",
    "2Th": "2 Thessalonians", "2Thes": "2 Thessalonians",
    "Heb": "Hebrews",
    "1Ti": "1 Timothy", "1Tim": "1 Timothy",
    "2Ti": "2 Timothy", "2Tim": "2 Timothy",
    "Tit": "Titus",
    "Phm": "Philemon",
    "Rev": "Revelation"
}

def standardize_book_name(raw_name):
    """Clean and map a raw book name to the standard form."""
    # Remove numbers or "The Book of" prefixes if present
    raw_name = raw_name.replace("The Book of", "").strip()
    # If the name starts with a number, separate it (e.g., "1Samuel" -> "1 Samuel")
    raw_name = re.sub(r'^(\d)([A-Za-z])', r'\1 \2', raw_name)

    # Check mapping, otherwise assume it's already standard
    return BOOK_NAME_MAP.get(raw_name, raw_name)

def parse_source_file(filepath):
    """Reads your source Bible file and returns a dictionary organized by book and chapter."""
    bible_dict = {}
    
    print(f"Opening file: {filepath}")  # Debug line

    with open(filepath, 'r', encoding='utf-8') as f:
        current_book = None
        current_chapter = None
        chapter_text = []

        for line in f:
            line = line.strip()
            if not line:
                continue

            # Check if the line looks like a book/chapter header
            match = re.match(r'^([\w\d\s]+)\s+(\d+)', line)
            if match:
                # Save the previous chapter's text if we have one
                if current_book and current_chapter is not None:
                    bible_dict.setdefault(current_book, {})[current_chapter] = '\n'.join(chapter_text)
                    chapter_text = []

                raw_book = match.group(1).strip()
                current_book = standardize_book_name(raw_book)
                current_chapter = int(match.group(2))
            else:
                # It's a verse line, add to the current chapter text
                if current_book and current_chapter is not None:
                    chapter_text.append(line)

        # Don't forget the last chapter
        if current_book and current_chapter is not None:
            bible_dict.setdefault(current_book, {})[current_chapter] = '\n'.join(chapter_text)

    return bible_dict

def reorder_and_output(bible_dict, output_filepath):
    """Takes the parsed dictionary, reorders by CUSTOM_ORDER, and writes to file."""
    with open(output_filepath, 'w', encoding='utf-8') as out_f:
        for book_name in CUSTOM_ORDER:
            if book_name in bible_dict:
                out_f.write(f"\n\n--- {book_name} ---\n\n")
                # Get chapters for this book, sort by chapter number
                chapters = bible_dict[book_name]
                for chap_num in sorted(chapters.keys()):
                    out_f.write(f"{book_name} {chap_num}\n")
                    out_f.write(chapters[chap_num])
                    out_f.write("\n\n")

# --- Main Execution ---
if __name__ == "__main__":
    # CHANGED: Use the actual file in your directory
    input_file = "kjv_formatted.txt"  # Changed from "path/to/your/kjv_source.txt"
    
    # CHANGED: Output to overwrite the original file (or use a different name)
    output_file = "kjv_formatted.txt"  # This will overwrite the original
    
    print("Parsing source file...")
    parsed_bible = parse_source_file(input_file)
    print(f"Found {len(parsed_bible)} books.")
    
    # Debug: Show which books were found
    print(f"Books found: {list(parsed_bible.keys())[:5]}...")  # Show first 5

    print("Reordering and writing output...")
    reorder_and_output(parsed_bible, output_file)

    print(f"Done! Output written to {output_file}")