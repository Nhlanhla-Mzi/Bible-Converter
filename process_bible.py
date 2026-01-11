import re

# Load your KJV text file
with open('kjv_source.txt', 'r', encoding='utf-8') as file:
    bible_text = file.read()

# 1. Change "Spirit" to lowercase "spirit" (for later Word formatting)
bible_text = re.sub(r'\b(Spirit)\b', 'spirit', bible_text)
# 2. Change "Holy Spirit/Ghost" to lowercase
bible_text = re.sub(r'Holy\s+(Ghost|Spirit)', 'holy spirit', bible_text, flags=re.IGNORECASE)
# 3. Change words in [brackets] to _italics_ format
bible_text = re.sub(r'\[(.*?)\]', r'_\1_', bible_text)

# Save the processed text to a new file
with open('kjv_formatted.txt', 'w', encoding='utf-8') as file:
    file.write(bible_text)

print("Processing complete! Check 'kjv_formatted.txt' in your file explorer.")
