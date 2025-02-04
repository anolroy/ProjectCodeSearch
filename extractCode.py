import os
import re
import json

def read_file_with_encoding(file_path):
    """
    Read a file with the appropriate encoding, handling cases where UTF-8 fails.
    """
    encodings = ['utf-8', 'latin1', 'cp1252']  # Common encodings for VB6 .frm files
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as file:
                return file.read()
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Unable to decode file: {file_path}")

def extract_sub_procedures(file_path):
    """
    Extract sub-procedures from a VB6 .frm file.
    """
    sub_procedures = {}
    content = read_file_with_encoding(file_path)
    # Regex to find sub procedures
    matches = re.findall(r"(Private|Public)\s+Sub\s+(\w+)[^\n]*\n(.*?)End Sub", content, re.DOTALL)
    for match in matches:
        sub_type, sub_name, sub_body = match
        full_procedure = f"{sub_type} Sub {sub_name}\n{sub_body.strip()}\nEnd Sub"
        sub_procedures[sub_name] = full_procedure
    return sub_procedures

def process_vb6_forms(directory):
    """
    Process all .frm files in the given directory.
    """
    data = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.frm'):
                file_path = os.path.join(root, file)
                try:
                    sub_procedures = extract_sub_procedures(file_path)
                    sorted_procedures = {k: sub_procedures[k] for k in sorted(sub_procedures)}
                    data.append({
                        "name": file,
                        **sorted_procedures
                    })
                except ValueError as e:
                    print(e)
    return data

def save_to_json(data, output_file):
    """
    Save the extracted data to a JSON file.
    """
    with open(output_file, 'w', encoding='utf-8') as json_file:
        json.dump({"data": data}, json_file, indent=4)

if __name__ == "__main__":
    # Input directory containing VB6 .frm files
    input_directory = "forms"
    # Output JSON file
    output_file = "output.json"
    
    # Process VB6 forms and extract sub-procedures
    extracted_data = process_vb6_forms(input_directory)
    
    # Save the extracted data to a JSON file
    save_to_json(extracted_data, output_file)
    print(f"Extracted data saved to {output_file}")
