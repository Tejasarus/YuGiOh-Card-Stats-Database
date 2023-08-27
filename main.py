from __future__ import print_function

import os.path
from collections import defaultdict
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from gsheets_tools import create_new_sheet, delete_sheet
import requests
import prices
from json import dumps
import yugioh
import re
api = prices.YGOPricesAPI()

#variables for global stats
max_price = 0
max_price_name = ""
max_price_card_name = ""

min_price = 999
min_price_name = ""
min_price_card_name = ""

card_total = 0
card_price_total = 0

print(f"Accessing Spreadsheet...")

#sort out credentials
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1ZouKj9lKDJWmz6_YWdCPsmd_hZ5Em3jGwITDD5oOyQo' #replace this with your spreadsheer ID (in http link after /d/)
creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

try:
    service = build('sheets', 'v4', credentials=creds)
    sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = sheet_metadata.get('sheets', []) #get all of the metadata for the spreadsheet

    # Iterate through sheets starting from the second sheet (index 1) to ignore the instructions sheet
    for sheet in sheets[1:]:   
        sheet_title = sheet['properties']['title']
        if "result" in sheet_title.lower(): #skip over sheet if its the result sheet
            print(f"Skipping sheet: {sheet_title} (contains 'result' in title)")
            continue

        print(f"Reading data from sheet {sheet_title}")

        range_name = f"{sheet_title}!B2:B" #second column, has the tags
        result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
        values = result.get('values', []) #put it in a dictionary
        
        #make new sheet for results and put data in
        if not values:
            print(f"No data found in sheet: {sheet_title}")
        else:
            print(f"Metadata read! Reading ID values and obtaining card data...")

            #get all of the data from each row
            sheet_data = {} # Outer dictionary, will store the data for each persons card collection
            # Inner dictionary for card data
            id_occurrences = defaultdict(int)
            id_name = defaultdict(str)
            id_card_type = defaultdict(str)
            id_property = defaultdict(str)
            id_family = defaultdict(str)
            id_type = defaultdict(str)
            id_rarity = defaultdict(str)
            id_price = defaultdict(float)
            id_total_value = defaultdict(float)
            id_link = defaultdict(str)
            total_price_of_collection = 0
            for row in values:
                if row:
                    id_value = row[0]
                    id_occurrences[id_value] += 1
                    url = "https://yugiohprices.com/api/price_for_print_tag/" + id_value
                    response = requests.get(url)
                    data = response.json()  # Parse JSON data
                    try: 
                        id_name[id_value] = data['data']['name']
                        id_card_type[id_value] = data['data']['card_type']
                        id_property[id_value] = data['data']['property']
                        id_family[id_value] = data['data']['family']
                        id_type[id_value] = data['data']['type']
                        id_rarity[id_value] = data['data']['price_data']['rarity']
                        id_price[id_value] = yugioh.get_card(card_name = data['data']['name']).tcgplayer_price
                        id_link[id_value] = "https://yugiohprices.com/card_price?name=" + re.sub(r"\s", "%20", data['data']['name'])
                        print(f"Card tag: {id_value} data obtained")
                    except: #if bad id, place error in its place
                        id_name[id_value] = "Error: ID not found"
                        id_card_type[id_value] = "Error: ID not found"
                        id_property[id_value] = "Error: ID not found"
                        id_family[id_value] = "Error: ID not found"
                        id_type[id_value] = "Error: ID not found"
                        id_rarity[id_value] = "Error: ID not found"
                        id_price[id_value] = 0
                        id_total_value[id_value] = 0
                        id_link[id_value] = "Error: ID not found"
                        print(f"Error reading from tag: {id_value}")
            
            # Calculate the total value for each occurrence
            

            #put them all together in the big dictionary
            for id_value, occurrences in id_occurrences.items():
                total_value = float(id_price[id_value]) * id_occurrences[id_value]
                id_total_value[id_value] = total_value
                total_price_of_collection = total_price_of_collection + total_value
                sheet_data[id_value] = {
                    'name': id_name[id_value],
                    'occurrences': occurrences,
                    'card type':id_card_type[id_value],
                    'property': id_property[id_value],
                    'family': id_family[id_value],
                    'type': id_type[id_value],
                    'rarity': id_rarity[id_value],
                    'price':id_price[id_value],
                    'total value':id_total_value[id_value],
                    'link':id_link[id_value]
                }

            #Write data to spreadsheet
            print(f"Writing data for: {sheet_title}")

            #Make a new sheet to hold results
            new_sheet_name = sheet_title + " results"
            create_new_sheet(service, SPREADSHEET_ID, new_sheet_name) 
            
            unique_ids = list(id_occurrences.keys())
            occurrences_list = [id_occurrences[id_value] for id_value in unique_ids]

            range_name = f"{new_sheet_name}!A2:A"
            range_id = f"{new_sheet_name}!B2:B" 
            range_occurrence = f"{new_sheet_name}!C2:C"
            range_card_type= f"{new_sheet_name}!D2:D"
            range_property= f"{new_sheet_name}!E2:E"
            range_family= f"{new_sheet_name}!F2:F"
            range_type= f"{new_sheet_name}!G2:G"
            range_rarity= f"{new_sheet_name}!H2:H"
            range_price= f"{new_sheet_name}!I2:I"
            range_total_value= f"{new_sheet_name}!J2:J"
            range_link= f"{new_sheet_name}!K2:K"
            value_input_option = "RAW"
            value_range_body = {"values": [[id_value] for id_value in unique_ids]}

            request_id = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_id,
            valueInputOption=value_input_option,
            body=value_range_body
            )
            request_occurrence = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_occurrence,
                valueInputOption=value_input_option,
                body={"values": [[occurrence] for occurrence in occurrences_list]}
            )
            request_name = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name,  # Use the correct range for the name column
                valueInputOption=value_input_option,
                body={"values": [[sheet_data[id_value]['name']] for id_value in unique_ids]}
            )
            request_card_type = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_card_type,  # Use the correct range for the name column
                valueInputOption=value_input_option,
                body={"values": [[sheet_data[id_value]['card type']] for id_value in unique_ids]}
            )
            request_property = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_property,  # Use the correct range for the name column
                valueInputOption=value_input_option,
                body={"values": [[sheet_data[id_value]['property']] for id_value in unique_ids]}
            )
            request_family = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_family,  # Use the correct range for the name column
                valueInputOption=value_input_option,
                body={"values": [[sheet_data[id_value]['family']] for id_value in unique_ids]}
            )
            request_type = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_type,  # Use the correct range for the name column
                valueInputOption=value_input_option,
                body={"values": [[sheet_data[id_value]['type']] for id_value in unique_ids]}
            )
            request_rarity = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_rarity,  # Use the correct range for the name column
                valueInputOption=value_input_option,
                body={"values": [[sheet_data[id_value]['rarity']] for id_value in unique_ids]}
            )
            request_price = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_price,  # Use the correct range for the name column
                valueInputOption=value_input_option,
                body={"values": [[sheet_data[id_value]['price']] for id_value in unique_ids]}
            )
            request_total_value = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_total_value,  # Use the correct range for the name column
                valueInputOption=value_input_option,
                body={"values": [[sheet_data[id_value]['total value']] for id_value in unique_ids]}
            )
            request_link = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_link,  # Use the correct range for the name column
                valueInputOption=value_input_option,
                body={"values": [[sheet_data[id_value]['link']] for id_value in unique_ids]}
            )
            response_id = request_id.execute()
            response_occurrence = request_occurrence.execute()
            response_name = request_name.execute()
            response_card_type = request_card_type.execute()
            response_property = request_property.execute()
            responese_family = request_family.execute()
            response_type = request_type.execute()
            response_rarity = request_rarity.execute()
            response_price = request_price.execute()
            response_total_value = request_total_value.execute()
            response_link = request_link.execute()
            

except HttpError as err:
    print(err)