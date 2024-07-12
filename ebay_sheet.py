import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
# SAMPLE_SPREADSHEET_ID = "1Qjyla8tBUtHJFEYWgwNmdb0O_LxTeon9Lj_LA5DzHwQ"
SAMPLE_RANGE_NAME = "商品管理表!A8:D"
def search_sku(sku, values):
    try:
        for item in values:
            if item[0] == sku:
                return item
    except:
        return None
def main(sku, product_sheet_id):
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        print(creds)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId=product_sheet_id, range=SAMPLE_RANGE_NAME)
            .execute()
        )
        values = result.get("values", [])

        if not values:
            print("No data found.")
            return 0
        
        ##### return values
        items = values
        item = search_sku(sku, items)
        try:
            print(f"Item 2: {item[2]}")
            print(f"Item 3: {item[3]}")
            product_name = item[3]
            person_charge = item[2]
        except TypeError:
            print("SKU not found")
            product_name = ""
            person_charge = ""
        except IndexError:
            print("Unexpected item format")
            product_name = ""
            person_charge = ""
        return product_name, person_charge
        #### return values end

    except HttpError as err:
        print(err)
if __name__ == "__main__":
    main()