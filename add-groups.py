import asyncio
import pandas as pd
import math
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.users_request_builder import UsersRequestBuilder
from msgraph.generated.models.reference_create import ReferenceCreate
from msgraph.generated.models.group import Group

tenantID = "xxxxxxx"
clientID = "yyyyyyy"
clientSecret = "zzzzzzz"

credentials = ClientSecretCredential(tenantID,clientID, clientSecret)
scopes = ['https://graph.microsoft.com/.default']
graph_client = GraphServiceClient(credentials=credentials, scopes=scopes)

excel_file = "users_togroup.xlsx"
sheet_name = "users_togroup"

excel = pd.read_excel(excel_file, sheet_name = sheet_name)

excel = excel.sample(frac=1).reset_index(drop=True)

async def get_users():
    query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
		select = ["userPrincipalName","displayName","id","mail"]
    )

    request_configuration = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
    query_parameters = query_params,
    )   

    result = await graph_client.users.get(request_configuration = request_configuration)
    print(result)

#asyncio.run(get_users())



async def assign_users():
   user_counter = 0
   print((len(excel.index)+1)/15.0)
   num_groups = math.ceil((len(excel.index) + 1)/15.0)
   done_user_counter = 0
   for group in range(0,num_groups):
        request_body = Group(
            description = "New Group",
            display_name = "New Group "+str(group),
            group_types = [
            ],
            mail_enabled = False,
            mail_nickname = "NewGroup"+str(group),
            security_enabled = True,
            additional_data = {
                "owners@odata_bind" : [
                    "https://graph.microsoft.com/v1.0/directoryObjects/xxxxxxx", 
                ],
            }
        )

        print("===============")
        print(request_body.display_name)

        result = await graph_client.groups.post(request_body)

        group_id = result.id

        group_users = []
        for ind in range(user_counter, user_counter+15):
            if ind < len(excel.index):
                request_body = ReferenceCreate(
                	odata_id = "https://graph.microsoft.com/v1.0/directoryObjects/"+str(excel["Object Id"][ind]),
                )
                print(excel["Display name"][ind])
                done_user_counter=done_user_counter+1
                result = await graph_client.groups.by_group_id(group_id).members.ref.post(request_body)
        user_counter = user_counter + 15
    
        print(done_user_counter)
asyncio.run(assign_users())

