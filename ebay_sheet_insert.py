import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
# SAMPLE_SPREADSHEET_ID = "1wuq9MBXS1iWcwTSUWRSn4mRMApUB23h_jZ0VkoHB680"

def get_credentials():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds
def duplicate_sheet(spreadsheet_id, source_sheet_name, new_sheet_name):
    creds = get_credentials()
    sheets = build("sheets", "v4", credentials=creds)
    spreadsheet = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    source_sheet_id = next((sheet['properties']['sheetId'] for sheet in spreadsheet['sheets'] if sheet['properties']['title'] == source_sheet_name), None)
    if source_sheet_id is None:
        print(f"Error: Sheet '{source_sheet_name}' not found in the spreadsheet.")
        return
    request_body = {
        "requests": [
            {
                "duplicateSheet": {
                    "sourceSheetId": source_sheet_id,
                    "newSheetName": new_sheet_name
                }
            }
        ]
    }
    response = sheets.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=request_body
    ).execute()
    print(f"Duplicated sheet: '{source_sheet_name}' to '{new_sheet_name}'")

def search_sku(buyer_name, sku, values):
    print(f"Buyer name : {buyer_name}")
    print(f"SKU : {sku}")
    try:
        for item in values:
            print(f"buyer_name in for : {item[5]}")
            print(f"sku in for : {item[8]}")
            if item[5] == buyer_name and item[8] == sku:
                return item
    except:
        return None
    
def insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id):
    create_sheet_w = int(date_sold.split("月")[0])
    sheet_name = f"販売管理{create_sheet_w}月"
    range_name = f"{sheet_name}!A6:AO"
    
    creds = get_credentials()
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId=order_sheet_id, range=range_name)
            .execute()
        )
        values = result.get("values", [])
        last_row = len(values) + 6
        profit = f"=AM{last_row}*($F$2-($F$2*AG{last_row}))-P{last_row}-N{last_row}-AJ{last_row}*$F$2-AK{last_row}"
        profit_including_refund = f"=AH{last_row}+P{last_row}*0.1+(AE{last_row}*Z{last_row}*($F$2-($F$2*AG{last_row}))*0.1)"
        remarks = f"=Z{last_row}+AC{last_row}-AN{last_row}"
        tax_fee = f"=AO{last_row}*1.1"
        total_fee = f"=((Z{last_row}+AA{last_row}+AC{last_row})*AE{last_row} + 0.4 + (Z{last_row}+AA{last_row}+AC{last_row})*AF{last_row} + (Z{last_row}+AA{last_row}+AC{last_row})*AD{last_row} + ((Z{last_row}+AA{last_row}+AC{last_row})*AD{last_row}*0.1))"

        values = [[states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, tax_fee, total_fee]]
        body = {"values": values}
        result = (sheet.values().append(
            spreadsheetId=order_sheet_id, 
            range=range_name, 
            valueInputOption="USER_ENTERED", 
            body=body
        ).execute())
        print("data insert")
        return result
    except HttpError as err:
        print(err)  
def update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id):
    create_sheet_w = int(date_sold.split("月")[0])
    sheet_name = f"販売管理{create_sheet_w}月"
    range_name = f"{sheet_name}!A6:AO"

    creds = get_credentials()
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()

        sheetNameresult = sheet.get(spreadsheetId=order_sheet_id).execute()
        sheet_names = [sheet['properties']['title'] for sheet in sheetNameresult.get('sheets', [])]
        if sheet_name not in sheet_names:
            duplicate_sheet(order_sheet_id, "販売管理", sheet_name)

        result = (
            sheet.values()
            .get(spreadsheetId=order_sheet_id, range=range_name)
            .execute()
        )
        values = result.get("values", [])
        item = search_sku(buyer_name, sku, values)
        if item:
            print(item)
            print("Ok matching data")
            profit = f"=AM{values.index(item)+6}*($F$2-($F$2*AG{values.index(item)+6}))-P{values.index(item)+6}-N{values.index(item)+6}-AJ{values.index(item)+6}*$F$2-AK{values.index(item)+6}"
            profit_including_refund = f"=AH{values.index(item)+6}+P{values.index(item)+6}*0.1+(AE{values.index(item)+6}*Z{values.index(item)+6}*($F$2-($F$2*AG{values.index(item)+6}))*0.1)"
            remarks = f"=Z{values.index(item)+6}+AC{values.index(item)+6}-AN{values.index(item)+6}"
            tax_fee = f"=AO{values.index(item)+6}*1.1"
            total_fee = f"=((Z{values.index(item)+6}+AA{values.index(item)+6}+AC{values.index(item)+6})*AE{values.index(item)+6} + 0.4 + (Z{values.index(item)+6}+AA{values.index(item)+6}+AC{values.index(item)+6})*AF{values.index(item)+6} + (Z{values.index(item)+6}+AA{values.index(item)+6}+AC{values.index(item)+6})*AD{values.index(item)+6} + ((Z{values.index(item)+6}+AA{values.index(item)+6}+AC{values.index(item)+6})*AD{values.index(item)+6}*0.1))"

            if tracking_info == "Refunded":
                states = "返品"
            elif tracking_info == "Partially refunded":
                states = "一部返金"
            elif tracking_info == "Canceled":
                states = "キャンセル"
            elif tracking_info == "Awaiting payment":
                states = "未確認"
                inspection = "決済待機中"
            elif "Ship by" in tracking_info:
                states = "未確認"
            else:
                states = "発送済"
            create_sheet_w = int(item[24].split("月")[0])
            sheet_name = f"販売管理{create_sheet_w}月"

            updated_row = [item[0], item[1], item[2], item[3], item[4], buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item[12], item[13], item[14], item[15], item[16], item[17], item[18], item[19], item[20], item[21], item[22], item[23], date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, item[33], item[34], item[35], item[36], tracking_info, remarks, tax_fee, total_fee]
            result = (sheet.values().update(
                spreadsheetId=order_sheet_id, 
                range=f"{sheet_name}!A{values.index(item)+6}:AO", 
                valueInputOption="USER_ENTERED", 
                body={"values": [updated_row]}
            ).execute())
            print("data update")
        else:
            print("No matching data")
            if tracking_info == "Refunded":
                states = "返品"
            elif tracking_info == "Partially refunded":
                states = "一部返金"
            elif tracking_info == "Canceled":
                states = "キャンセル"
            elif tracking_info == "Awaiting payment":
                states = "未確認"
                inspection = "決済待機中"
            elif "Ship by" in tracking_info:
                states = "未確認"
            else:
                states = "発送済"
            
            insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)

    except HttpError as err:
        print(err)
if __name__ == "__main__":
    update_or_insert_data()