"""
main.py
----------------
Tool that converts XML based NER tags in Excel file in column NER with
sentence_id into Conll format.

Author: Arun
Date: March 9, 2026
"""

import pandas as pd
import os
import re

def convert_to_conll(input_excel):
    """
    Convert Input excel with NER and sentence_id to Conll format
    """

    file_root, _ = os.path.splitext(input_excel)

    # Append the new extension
    output_conll = f"{file_root}.conll"


    # Get the base name for the error file
    base_name = os.path.splitext(os.path.basename(input_excel))[0]
    error_file_name = f"{base_name}_error.txt"

    # Load the Excel file
    df = pd.read_excel(input_excel)

    # Regex to capture any tag name and its TYPE attribute
    # Group 1: Tag Name, Group 2: TYPE value, Group 3: Content
    tag_pattern = re.compile(r'<(\w+)\s+TYPE\s*=\s*"([^"]+)">\s*(.*?)\s*</\1>')

    with open(output_conll, 'w', encoding='utf-8') as conll_file, \
         open(error_file_name, 'w', encoding='utf-8') as error_file:

        for _, row in df.iterrows():
            sentence_id = row['sentence_id']
            # Using 'NER' as the column name—change to 'POS' if your column is named that
            ner_text = str(row.get('NER', row.get('POS', ''))) 

            try:
                last_idx = 0
                tokens_found = False

                # We iterate through the string to find tagged entities and the "O" text between them
                for match in tag_pattern.finditer(ner_text):
                    tokens_found = True
                    
                    # 1. Process "O" text before the tag
                    pre_text = ner_text[last_idx:match.start()].strip()
                    if pre_text:
                        for word in pre_text.split():
                            conll_file.write(f"{sentence_id}\t{word}\tO\n")

                    # 2. Process the Tagged Entity
                    label = match.group(2)
                    content = match.group(3).strip()
                    words = content.split()

                    for i, word in enumerate(words):
                        prefix = "B-" if i == 0 else "I-"
                        conll_file.write(f"{sentence_id}\t{word}\t{prefix}{label}\n")
                    
                    last_idx = match.end()

                # 3. Process remaining "O" text after the last tag
                post_text = ner_text[last_idx:].strip()
                if post_text:
                    tokens_found = True
                    for word in post_text.split():
                        conll_file.write(f"{sentence_id}\t{word}\tO\n")

                if not tokens_found:
                    raise ValueError("No text or tags found in this row.")

                # Add a blank line after each sentence to separate them in CoNLL
                conll_file.write("\n")

            except Exception as e:
                error_file.write(f"Error in Sentence ID: {sentence_id}\n")
                error_file.write(f"Content: {ner_text}\n")
                error_file.write(f"Message: {str(e)}\n\n")
                print(f"Error in Sentence ID: {sentence_id}, logged to error file.")

    print(f"Conversion complete. CoNLL file saved to: {output_conll}")

# Example usage:
input_excel = 'data/ner/input.xlsx'

convert_to_conll(input_excel)