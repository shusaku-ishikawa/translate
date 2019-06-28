from __future__ import print_function
from time import sleep 
import os, sys, datetime
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import gspread
from gspread.exceptions import *
from oauth2client.service_account import ServiceAccountCredentials
import shutil
import base64   
import pickle
import os.path
import io
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient.http import MediaIoBaseDownload
from google.cloud import translate_v3beta1 as translate
import csv

scope = ['https://www.googleapis.com/auth/drive']

def connect_driveapi(picklefile, client_id_json):

    ''' oauth処理 '''
    creds = None
    if os.path.exists(picklefile):
        with open(picklefile, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                client_id_json, scope)
            creds = flow.run_local_server()
        with open(picklefile, 'wb') as token:
            pickle.dump(creds, token)

    return build('drive', 'v3', credentials=creds)

def get_drive_files(service):
    results = service.files().list(
        pageSize=1,
        q="'1Gj7lRZ47TRK4EafzKIoe12mPBt0PJQOS' in parents",
        spaces='drive',
        fields="nextPageToken, files(id, name)"
    ).execute()
    items = results.get('files', [])
    file_idname_list = []
    if not items:
        print('No files found.')
    else:
        for item in items:
            file_idname_list.append((item['id'], item['name']))
    return file_idname_list

def download_drive_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    file_path = file_id + '.tsv'
    fh = io.FileIO(file_path, mode='w')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
    return file_path

def get_gspread_book(secret_key, book_name):
    

    credentials = ServiceAccountCredentials.from_json_keyfile_name(secret_key, scope)
    gc = gspread.authorize(credentials)
    book = gc.open(book_name)
    return book

def paste_csv_to_gspread(book, sheet_name, csv_file, cell):
    try:
        ''' sheetを取得し、値を初期化 '''
        sheet = book.worksheet(sheet_name)
        sheet.clear()
    except WorksheetNotFound:
        print('タブ: ' + sheet_name + ' が見つかりませんでした。新規作成します。')
        sheet = book.add_worksheet(title=sheet_name, rows=1000, cols=30)

    (firstRow, firstColumn) = gspread.utils.a1_to_rowcol(cell)

    with open(csv_file, 'r', encoding='shift_jis') as f:
        content = f.read()
    body = {
        'requests': [{
            'pasteData': {
                "coordinate": {
                    "sheetId": sheet.id,
                    "rowIndex": firstRow-1,
                    "columnIndex": firstColumn-1,
                },
                "data": content,
                "type": 'PASTE_NORMAL',
                "delimiter": ',',
            }
        }]
    }
    return book.batch_update(body)


def translate_to_jp(project_id, service, engword):
    location = 'global'

    parent = service.location_path(project_id, location)

    response = service.translate_text(
        parent=parent,
        contents=[engword],
        mime_type='text/plain',  # mime types: text/plain, text/html
        source_language_code='en',
        target_language_code='ja')
    
    result = response.translations
    if len(result) > 0:
        return result[0].translated_text
    print('翻訳に失敗しました : {} '.format(engword))
    return ''

''' 以下メイン処理 '''
if __name__ == '__main__':

    project_id = 'lancers01'

    ''' 認証情報 '''
    client_id_json = 'secret_key/clientsecret.json'
    picklefile = 'secret_key/token.pickle'
    
    service_account = 'secret_key/serviceaccount.json'
    
    ''' Sheet Name '''
    #sheetname = 'tranlate'

    '''  '''
    #book = get_gspread_book(service_account, sheetname)


    drive_service = connect_driveapi(picklefile, client_id_json)
    translate_service = translate.TranslationServiceClient()
    #translate_to_jp(project_id, translate_service, 'hello')

    file_idname_list = get_drive_files(drive_service)
    
    for fileinf in file_idname_list:
        fileid = fileinf[0]
        filename = fileinf[1]

        file_path = download_drive_file(drive_service, fileid)
        print('processing {}'.format(filename))
        with open(file_path) as tsvfile:
            reader = csv.reader(tsvfile, delimiter='\t')
            header = next(reader)
            for row in reader:
                engword = row[0]
                print(engword)
                print(translate_to_jp(project_id, translate_service, engword))

        os.remove(file_path)



