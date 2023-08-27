from __future__ import print_function

import os.path
from collections import defaultdict
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


def delete_sheet(service, spreadsheet_id, sheet_title):
    try:
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', [])
        sheet_id_to_delete = None

        for sheet in sheets:
            if sheet['properties']['title'] == sheet_title:
                sheet_id_to_delete = sheet['properties']['sheetId']
                break

        if sheet_id_to_delete:
            request = service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "requests": [
                        {"deleteSheet": {"sheetId": sheet_id_to_delete}}
                    ]
                }
            )
            response = request.execute()
            print(f"Sheet '{sheet_title}' deleted.")

    except HttpError as err:
        print(f"An error occurred while deleting sheet: {err}")


def create_new_sheet(service, spreadsheet_id, new_sheet_title):
    try:
        
        delete_sheet(service, spreadsheet_id, new_sheet_title)
        
        batch_update_spreadsheet_request_body = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": new_sheet_title
                        }
                    }
                }
            ]
        }

        request = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=batch_update_spreadsheet_request_body
        )
        response = request.execute()

        # Get the new sheet ID
        new_sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']

        # Update the headers in the new sheet
        update_header_request = {
            "updateCells": {
                "rows": [
                    {
                        "values": [
                            {"userEnteredValue": {"stringValue": "Name"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                            {"userEnteredValue": {"stringValue": "ID"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                            {"userEnteredValue": {"stringValue": "Occurrence"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                            {"userEnteredValue": {"stringValue": "Card Type"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                            {"userEnteredValue": {"stringValue": "Property"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                            {"userEnteredValue": {"stringValue": "Family"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                            {"userEnteredValue": {"stringValue": "Type"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                            {"userEnteredValue": {"stringValue": "Rarity"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                            {"userEnteredValue": {"stringValue": "Price Per Card"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                            {"userEnteredValue": {"stringValue": "Total Value"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                            {"userEnteredValue": {"stringValue": "Link"}, "userEnteredFormat": {"textFormat": {"bold": True}}},
                        ]
                    }
                ],
                "fields": "userEnteredValue,userEnteredFormat.textFormat.bold",  # Specify bold property here
                "start": {
                    "sheetId": new_sheet_id,
                    "rowIndex": 0,
                    "columnIndex": 0
                }
            }
        }


        request = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [update_header_request]}
        )
        response = request.execute()

        print(f"New sheet '{new_sheet_title}' created with headers.")

    except HttpError as err:
        print(f"An error occurred while creating sheet: {err}")