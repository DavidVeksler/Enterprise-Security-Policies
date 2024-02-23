import zipfile
import re
import os
import tempfile
from xml.etree import ElementTree as ET

def replace_text_in_xml(xml_content, target, replacement):
    """Replace text in XML content."""
    # This uses a simple string replace. For more complex replacements (e.g., considering word boundaries),
    # you might need to use regular expressions.
    return xml_content.replace(target, replacement)

def process_docx(file_path, target, replacement):
    # Temporary directory to extract the docx files
    temp_dir = tempfile.mkdtemp()
    new_zip_path = file_path

    # Extract the docx file's contents
    with zipfile.ZipFile(file_path, 'r') as docx_file:
        docx_file.extractall(temp_dir)

    # Path to the document.xml file which contains the main text
    document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')

    # Read, replace, and overwrite the document.xml content
    with open(document_xml_path, 'r', encoding='utf-8') as xml_file:
        xml_content = xml_file.read()

    new_xml_content = replace_text_in_xml(xml_content, target, replacement)

    with open(document_xml_path, 'w', encoding='utf-8') as xml_file:
        xml_file.write(new_xml_content)

    # Create a new docx file from the modified contents
    with zipfile.ZipFile(new_zip_path, 'w', zipfile.ZIP_DEFLATED) as new_docx:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                # Create the correct path for the file to be stored in the zip
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_dir)
                new_docx.write(file_path, arcname)

    # Cleanup the temporary directory
    os.system(f'rm -rf {temp_dir}')

    return new_zip_path

# Example usage
directory_path = "/Users/davidv/Projects/Enterprise-Security-Policies"
target_word = "[Company Name]"
replacement_word = "[Your Company Name :-)]"

for filename in os.listdir(directory_path):
    if filename.endswith(".docx"):
        file_path = os.path.join(directory_path, filename)
        new_file_path = process_docx(file_path, target_word, replacement_word)
        print(f"Processed {filename}, saved as {new_file_path}")
