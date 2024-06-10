import os
import openai
import PyPDF2
import docx
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging

# Configure logging
logging.basicConfig(filename='cv_analysis.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Set your OpenAI API key securely (avoid storing directly in code)
openai.api_key = ("This is the secret magic")

def is_valid_file(file_path):
  """Checks if the file exists and has content."""
  if not os.path.isfile(file_path):
    logging.error(f"File not found: {file_path}")
    return False
  if os.path.getsize(file_path) == 0:
    logging.warning(f"Empty file: {file_path}")
    return False
  return True

def extract_text_from_pdf(pdf_path):
    text = ''
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text += page.extract_text()
    return text


def extract_text_from_docx(docx_path):
    doc = docx.Document(docx_path)
    text = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
    return text


def extract_text_from_file(file_path):
    if not is_valid_file(file_path):
        return None
    if file_path.lower().endswith('.pdf'):
        return extract_text_from_pdf(file_path)
    elif file_path.lower().endswith('.docx'):
        return extract_text_from_docx(file_path)
    else:
        raise ValueError('Unsupported file format')


def analyze_cv(cv_text):
    prompt =[
            {"role": "user",
             "content": f"Extract the following details from the CV: Name, Date of Birth, Education, Skills.\n\n{cv_text}"
             },
        ]
    response = openai.chat.completions.create(
        model="gpt-4-1106-preview",
        messages=prompt
    )
    return response.choices[0].message.content.strip()


def process_cv(file_path):
    try:
        cv_text = extract_text_from_file(file_path)
        if cv_text is None:
            return {'File Name': os.path.basename(file_path), 'Extracted Info': 'Error'}
        extracted_info = analyze_cv(cv_text)
        return {'File Name': os.path.basename(file_path), 'Extracted Info': extracted_info}
    except Exception as e:
        logging.error(f"Failed to process {file_path}: {e}")
        return {'File Name': os.path.basename(file_path), 'Extracted Info': 'Error'}


def process_cvs_in_folder(folder_path):
    data = []
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = []
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if file_path.lower().endswith(('.pdf', '.docx')):
                futures.append(executor.submit(process_cv, file_path))

        for future in as_completed(futures):
            data.append(future.result())

    df = pd.DataFrame(data)
    return df


# Define the folder containing CVs (ensure the folder exists)
folder_path = '/Users/andreaslukas/Desktop/CV'
if not os.path.exists(folder_path):
  logging.error(f"Folder not found: {folder_path}")
  exit(1)

# Process CVs and get the result as a DataFrame
cv_data_df = process_cvs_in_folder(folder_path)

# Display or save the DataFrame (you can modify for different outputs)
print(cv_data_df)
cv_data_df.to_csv('cv_extracted_data.csv', index=False)
