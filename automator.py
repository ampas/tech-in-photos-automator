import requests
import openpyxl
import json

def create_topic(api_key, api_username, topic_title, raw_template, data):
    base_url = 'https://forums.techinphotos.com/'
    category = '6'  # Always post to category 6 'photo-identification'

    headers = {
        'Api-Key': api_key,
        'Api-Username': api_username,
    }

    raw = raw_template.format(**data)  # format the raw string with data from the Excel row

    payload = {
        'title': topic_title,
        'raw': raw,
        'category': category
    }

    response = requests.post(f'{base_url}/posts', headers=headers, json=payload)

    if response.status_code == 200:
        print(f'Successfully created topic "{topic_title}"')
    else:
        print(f'Failed to create topic "{topic_title}"')

def main():
    # Read json configuration file with store api key details
    with open('config.json') as f:
        config = json.load(f)

    api_key = config['API_KEY']
    api_username = config['API_USERNAME']

    # Load the raw template from the .md file
    with open('template.md', 'r') as template_file:
        raw_template = template_file.read()

    wb = openpyxl.load_workbook('test.xlsx')
    sheet = wb.active

    for row in sheet.iter_rows(min_row=2, values_only=True):  # skip the first row if it's a header
        project_id, image, id_num, topic_title, library_desc, search_terms, subject_headings = row  # adjust the unpacking

        # split the cell's content by line breaks, remove leading hyphens, and
        # format as bullet points
        if topic_title is None:
            print("Skipping a row because the topic title is None.")
            continue  # Skip the rest of this loop iteration and go to the next row
        if subject_headings is None:
            subject_headings = ''  # Use an empty string if the cell is empty
        subject_headings = '\n'.join([f'* {line.lstrip("— ").strip()}' for line in subject_headings.split('\n') if line.strip()])

        data = {
            'project_id': project_id,
            'image': image,
            'id_num': id_num,
            'library_desc': library_desc,
            'topic_title': topic_title,
            'search_terms': search_terms,
            'subject_headings': subject_headings,
        }
        create_topic(api_key, api_username, topic_title, raw_template, data)

if __name__ == '__main__':
    main()
