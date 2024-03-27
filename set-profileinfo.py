import pandas as pd
import asyncio
from kiota_abstractions.api_error import APIError
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.user import User
from msgraph.generated.models.authorization_info import AuthorizationInfo
from msgraph.generated.models.reference_update import ReferenceUpdate
import io
import math

tenantID = "xxxx"
clientID = "yyyyy"
clientSecret = "zzzzz"

credentials = ClientSecretCredential(tenantID,clientID, clientSecret)
scopes = ['https://graph.microsoft.com/.default']
client = GraphServiceClient(credentials=credentials, scopes=scopes)

# GET /users/{id | userPrincipalName}
async def get_user(UPN):
    try:
        user = await client.users.by_user_id(UPN).get()
        print(user.display_name)
    except APIError as e:
        print(f'Error: {e.error.message}')

#asyncio.run(get_user('me@ma.com'))




async def get_users():            
    users = await client.users.get()
    file = io.open("userid.csv", mode="w", encoding="utf-8")

    file.write("UserID\tDisplayName\tMail\tJobTitle\n")
    
    if users and users.value:
        for user in users.value:
            print(user.id, "\t", user.display_name, "\t", user.mail, "\t", user.job_title)
            file.write(user.id+ "\t" + user.display_name + "\t")
            if user.mail is None:
                file.write("None\n")
            else:
                file.write(user.mail + "\n")
    while users is not None and users.odata_next_link is not None:
        users = await client.users.with_url(users.odata_next_link).get()
        for user in users.value:
            print(user.id, "\t", user.display_name, "\t", user.mail, "\t", user.job_title)
            file.write(user.id+ "\t" + user.display_name + "\t")
            if user.mail is None:
                file.write("None\n")
            else:
                file.write(user.mail + "\n")
    file.close() 
#asyncio.run(get_users())

async def set_employeeid(UPN, EmployeeID):
    request_body = User(
        employee_id=EmployeeID
    )

    result = await client.users.by_user_id(UPN).patch(request_body)
    return result

async def set_employeeidmain():

    excel_file = "Report-2024-03-13-16-08-27.xlsx"
    sheet_name = "Report"

    excel = pd.read_excel(excel_file, sheet_name = sheet_name)
    tasks =[]

    for ind in excel.index:
        UserID=str(excel["UserID"][ind])
        EmployeeID = str(excel["EmployeeID"][ind])
        if UserID != "---" and EmployeeID != "---":
            print(UserID + "  " + EmployeeID)
            await set_employeeid(UserID, EmployeeID)

        # if dfusers.loc[dfusers[1]]
    # for ind in excel.index:
    # print(excel["Team Member: Name"])

asyncio.run(set_employeeidmain())


async def set_user_jobtitle(UPN, country, JobTitle, mobilePhone, department ):   
    if mobilePhone=="":
        request_body = User(
            country=country,
            job_title=JobTitle,
            department=department
        )
    else:
        request_body = User(
            country=country,
            job_title=JobTitle,
            mobile_phone=mobilePhone,
            department=department
        )

    result = await client.users.by_user_id(UPN).patch(request_body)
    return result

async def set_maininfo():
    excel_file = "Report-2024-03-13-16-08-27.xlsx"
    sheet_name = "Report"

    excel = pd.read_excel(excel_file, sheet_name = sheet_name)
    tasks =[]

    for ind in excel.index:
        UserID=str(excel["UserID"][ind])
        country=str(excel["Home Address Country"][ind])
        JobTitle = str(excel["Job Title"][ind])
        if not isinstance(excel["Mobile"][ind], str):
            mobile=""
        else:
            mobile = excel["Mobile"][ind]
        # if mobile=="nan":
        #     mobile =""
        # if mobile is None:
        #     mobile =""
        Department = str(excel["Function"][ind])
        if UserID != "---":
            print(UserID,"  ",country, "  ", JobTitle, "  ", mobile , "  ", Department)
            await set_user_jobtitle(UserID, country, JobTitle, mobile, Department)

        # if dfusers.loc[dfusers[1]]
    # for ind in excel.index:
    # print(excel["Team Member: Name"])

#asyncio.run(set_maininfo())

async def set_user_manager(UPN, ManagerID):   
    request_body = ReferenceUpdate(
	    odata_id = "https://graph.microsoft.com/v1.0/users/{0}".format(ManagerID)
    )

    print(request_body.odata_id)
    await client.users.by_user_id(UPN).manager.ref.put(request_body)

async def set_manager():
    excel_file = "Report-2024-03-13-16-08-27.xlsx"
    sheet_name = "Report"

    excel = pd.read_excel(excel_file, sheet_name = sheet_name)
    tasks =[]

    for ind in excel.index:
        UserID=str(excel["UserID"][ind])
        ManagerID=str(excel["ManagerId"][ind])
        if ManagerID != "---" and UserID != "---":
            print(UserID,"  ",ManagerID)
            await set_user_manager(UserID, ManagerID)


#asyncio.run(set_manager())