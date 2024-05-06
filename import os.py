import os
from docx import Document
import re
from docx.opc.exceptions import PackageNotFoundError

def search_and_record(directory):
    responses = []
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".docx"):
                filepath = os.path.join(root, file)
                try:
                    document = Document(filepath)
                    print("Processing file:", filepath)
                    for paragraph in document.paragraphs:
                        match = re.search(r"Total AHI \(events/hr\)\s*=\s*", paragraph.text)
                        if match:
                            print("Found match in paragraph:", paragraph.text)
                            response = paragraph.text[match.end():].strip()[:5]  # Take the first 5 characters after "="
                            responses.append((file, response))
                except PackageNotFoundError as e:
                    print(f"Error: {e}. Skipping file: {filepath}")
    
    return responses

def save_responses(responses, output_file):
    print("Saving responses to:", output_file)
    with open(output_file, "w") as file:
        for filename, response in responses:
            print(f"Writing: {filename}: AHI = {response}")
            file.write(f"{filename}: AHI = {response}\n")

directory = r"R:\0 FINALISED STUDIES"
output_file = r"C:\Users\jspa8961\Documents\Activity Reports\responses.txt"

responses = search_and_record(directory)
print("Responses:", responses)
save_responses(responses, output_file)
