# https://learn.microsoft.com/en-us/graph/api/listitem-create?view=graph-rest-1.0&tabs=python

import pandas as pd
import numpy
import asyncio
import datetime
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.list_item import ListItem
from msgraph.generated.models.field_value_set import FieldValueSet

tenantID = "xxxxx"
clientID = "yyyyy"
clientSecret = "zzzzzz"

site_id = "site.sharepoint.com,xxxxx,yyyyy"

list_id = "aaaaa"

excel_skip_columns = ["Column1","Company country", "Project value Value", "Item Type", "Path"]
excel_date_columns = ["Input date","Signature date", "Contract Date", "Offer Date"]

excel_field_lookup = {
    "Responsibility":  "field_1",
    "Company":"field_2",
    "Product":"field_3",
    "Status":"field_4",
    "Priority":"field_5",
    "Input date":"field_11",
    "Signature date":"field_12",
    "Contract Date":"field_13",
    "Offer Date":"field_14",
 
}

excel_file = "file.xlsx"
sheet_name = "query (1)"

excel = pd.read_excel(excel_file, sheet_name = sheet_name)
credentials = ClientSecretCredential(tenantID,clientID, clientSecret)
scopes = ['https://graph.microsoft.com/.default']
graph_client = GraphServiceClient(credentials=credentials, scopes=scopes)



async def get_site():
    result = await graph_client.sites.by_site_id(site_id=site_id).lists.by_list_id(list_id=list_id).items.get
    print(result)



#asyncio.run(get_site())

async def insert_line(n):

    value_dict ={}

    #Only the dropped needs to be imported
    if excel["Status"][n] != "5. Dropped":
        #print("Not line of dropped")
        return

    #Skip empty lines
    if pd.isna(excel["Responsibility"][n]):
        #print("skipping line " + str(n))
        return
    
    value_dict["Title"] = "Dropped " + str(n)

    # Project value Value field exists in Excel bt not in SP
    for col in excel.columns:
        if not col in excel_skip_columns:
            sp_field = excel_field_lookup[col]
            if not pd.isna(excel[col][n]):
                if isinstance(excel[col][n], numpy.float64 ):
                    if excel[col][n].item().is_integer():
                        value_dict[sp_field] = int(excel[col][n].item())
                    else:
                        value_dict[sp_field] = excel[col][n].item()
                else:
                    value = excel[col][n]
                    if col in excel_date_columns:
                        #value = value.strftime("%Y-%m-%dT%H:%M:%SZ")
                        value = value.strftime("%Y-%m-%dT22:00:00Z")
                        value_dict[sp_field] = value
                    else:
                        value_dict[sp_field] = excel[col][n]
            
    print(value_dict)

    request_body = ListItem(
        fields = FieldValueSet(
            additional_data = value_dict))


    result = await graph_client.sites.by_site_id(site_id).lists.by_list_id(list_id).items.post(request_body)

    print(result)

async def insert_line_runner():
    for row in range(1, len(excel)):
        await insert_line(row)

asyncio.run(insert_line_runner())


# request_body = ListItem(
# 	fields = FieldValueSet(
# 		additional_data = {
# 				"title" : "Widget",
# 				"color" : "Purple",
# 				"weight" : 32,
# 		}
# 	),
# )

# result = await graph_client.sites.by_site_id('site-id').lists.by_list_id('list-id').items.post(request_body)

