import requests
import pandas as pd
import json
import argparse
import re

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

def validate_data(df):
    # Check for duplicate topic titles
    duplicate_titles = df[df.duplicated(subset=['topic_title'], keep=False)]['topic_title']
    if not duplicate_titles.empty:
        print("Duplicate topic titles and their matching row numbers:")
        duplicate_dict = {}
        for title in duplicate_titles.unique():
            matching_rows = df[df['topic_title'] == title]
            row_numbers = matching_rows.index + 2
            duplicate_dict[title] = row_numbers
        for title, rows in duplicate_dict.items():
            if len(rows) > 1:
                print(f"- {title}: Rows {', '.join(map(str, rows[:-1]))}, and {rows[-1]} match")
        return False

    # Check for missing required columns
    required_columns = ['project_id', 'image', 'id_num', 'library_desc', 'topic_title', 'search_terms']
    missing_data = df[df[required_columns].isnull().any(axis=1)]

    # Replace "https://www.dropbox.com" with "https://dl.dropboxusercontent.com" in string columns
    string_columns = ['project_id', 'image', 'id_num', 'library_desc', 'topic_title', 'search_terms', 'subject_headings']
    for col in string_columns:
        df[col] = df[col].apply(lambda x: re.sub(r'https://www.dropbox.com', 'https://dl.dropboxusercontent.com', str(x)))

    if not missing_data.empty:
        print("Missing required data in the following rows:")
        for index, row in missing_data.iterrows():
            missing_columns = [col for col in required_columns if pd.isna(row[col])]
            print(f"Row {index + 2}: Missing data in columns: {', '.join(missing_columns)}")
        return False

    return True

def main():
    parser = argparse.ArgumentParser(description='Process an Excel file.')
    parser.add_argument('excel_file', help='Path to the Excel file')

    # Parse arguments
    args = parser.parse_args()

    # Read json configuration file with store api key details
    with open('config.json') as f:
        config = json.load(f)

    api_key = config['API_KEY']
    api_username = config['API_USERNAME']

    # Load the raw template from the .md file
    with open('template.md', 'r') as template_file:
        raw_template = template_file.read()

    # Load the Excel file and specify dtype=str to read all data as strings
    df = pd.read_excel(args.excel_file, dtype=str)

    # Validate the data before processing
    if not validate_data(df):
        return

    for index, row in df.iterrows():
        data = {
            'project_id': row['project_id'],
            'image': row['image'],
            'id_num': row['id_num'],
            'library_desc': row['library_desc'],
            'topic_title': row['topic_title'],
            'search_terms': row['search_terms'],
            'subject_headings': row['subject_headings'],
        }

        create_topic(api_key, api_username, row['topic_title'], raw_template, data)

if __name__ == '__main__':
    main()
