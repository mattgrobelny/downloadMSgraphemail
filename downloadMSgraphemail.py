
import os
import requests
from ms_graph import generate_access_token

# from https://learndataanalysis.org/source-code-download-outlook-email-attachments-using-microsoft-graph-api-in-python/


def download_email_attachments(message_id, headers, save_folder=os.getcwd()):
    try:
        response = requests.get(
            GRAPH_API_ENDPOINT + '/me/messages/{0}/attachments'.format(message_id),
            headers=headers
        )

        attachment_items = response.json()['value']
        for attachment in attachment_items:
            file_name = attachment['name']
            attachment_id = attachment['id']
            attachment_content = requests.get(
                GRAPH_API_ENDPOINT + '/me/messages/{0}/attachments/{1}/$value'.format(message_id, attachment_id),  headers=headers
            )
            print('Saving file {0}...'.format(file_name))
            with open(os.path.join(save_folder, file_name), 'wb') as _f:
                _f.write(attachment_content.content)
        return True
    except Exception as e:
        print(e)
        return False

# Step 1. Get the access token
APP_ID = '<app id>'
SCOPES = ['Mail.ReadWrite']
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'

access_token = generate_access_token(app_id=APP_ID, scopes=SCOPES)
headers = {
    'Authorization': 'Bearer ' + access_token['access_token']
}

# Step 2. Retrieve emails
params = {
    'top': 3, # max is 1000 messages per request
    'select': 'subject,hasAttachments',
    'filter': 'hasAttachments eq true',
    'count': 'true'
}

response = requests.get(GRAPH_API_ENDPOINT + '/me/mailFolders/inbox/messages', headers=headers, params=params)
if response.status_code != 200:
    raise Exception(response.json())

response_json = response.json()
response_json.keys()

response_json['@odata.count']

emails = response_json['value']
for email in emails:
    if email['hasAttachments']:
        email_id = email['id']
        download_email_attachments(email_id, headers, r'C:\Users\Jie\Desktop\Lesson\Attachments')
