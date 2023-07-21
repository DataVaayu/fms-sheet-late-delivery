# ___________________CredentialsAndConnectingToGoogleSheets_____________________________________________________

from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pandas as pd

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1QUd5cRH05AzlFInej1TtMGesHHS_q3Z0YloGDGi_4G0'
SAMPLE_RANGE_NAME = 'FMS!A:BY'



"""Shows basic usage of the Sheets API.
Prints values from a sample spreadsheet.
"""
creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

try:
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
        
    # creating the dataframe by reading each row
    df = pd.DataFrame.from_records(values)

except HttpError as err:
    print(err)

    
# __________________________________________________RestructuringColumns_________________________________________________________
    
# making a working copy of the dataframe
df_for_processing = df

# first level column names
col_index_row = [i for i in df_for_processing.loc[1] if i not in [None,""]]

# printing the topmost row, I.E, the first level of columns
# print(col_index_row)

# dropping the first row and few other rows
# as they have no relevant information

df.drop(index=[0,2,3,4],axis=0,inplace=True)

# Creating the top level of columns

first_level_cols=["Timestamp",
                  "Create Order",
                  "Create Order",
                  "Create Order",
                  "Create Order",
                  "Create Order",
                  "Create Order",
                  "Create Order",
                  "Create Order",
                  "Create Order",
                  "Create Order",
                  "Create Order",
                  "Create Order",
                  "Stock Availability in Pretture",
                  "Stock Availability in Pretture",
                  "Stock Availability in Pretture",
                  "Stock Availability in Pretture",
                  "Stock Availability in PRO(production)",
                  "Stock Availability in PRO(production)",
                  "Stock Availability in PRO(production)",
                  "Stock Availability in PRO(production)",
                  "Raise PRO(production) if not available anywhere",
                  "Raise PRO(production) if not available anywhere",
                  "Raise PRO(production) if not available anywhere",
                  "Raise PRO(production) if not available anywhere",
                  "Raise PO",
                  "Rec. from Production",
                  "Rec. from Production",
                  "Rec. from Production",
                  "Rec. from Production",
                  "Alteration needed?",
                  "Alteration needed?",
                  "Alteration needed?",
                  "Alteration needed?",
                  "Alteration requisition form \n\n( if Alteration required )",
                  "Rec. to Alteration",
                  "Rec. to Alteration",
                  "Rec. to Alteration",
                  "Rec. to Alteration",
                  "Received pcs in Warehouse 2",
                  "Received pcs in Warehouse 2",
                  "Received pcs in Warehouse 2",
                  "Received pcs in Warehouse 2",
                  "QC Approve ",
                  "QC Approve ",
                  "QC Approve ",
                  "QC Approve ",
                  "Send to Production",
                  "Send to Production",
                  "Send to Production",
                  "Send to Production",
                  "GR in Tally",
                  "GR in Tally",
                  "GR in Tally",
                  "GR in Tally",
                  "Update as Ready for dispatch",
                  "Update as Ready for dispatch",
                  "Update as Ready for dispatch",
                  "Update as Ready for dispatch",
                  "Generate Tax Invoice Pretture",
                  "Generate Tax Invoice Pretture",
                  "Generate Tax Invoice Pretture",
                  "Generate Tax Invoice Pretture",
                  "Dispatch",
                  "Dispatch",
                  "Dispatch",
                  "Dispatch",
                  "Generate Tax Invoice",
                  "Generate Tax Invoice",
                  "Generate Tax Invoice",
                  "Generate Tax Invoice",
                  "Connect with Client",
                  "Connect with Client",
                  "Connect with Client",
                  "Connect with Client",
                  "Status",
                  "Remarks"]

# creating the second level of columns

second_level_cols=df_for_processing.loc[5]

# combining the two levels of columns using zip

index_tuple_cols = list(zip(first_level_cols,second_level_cols))
df.columns = index_tuple_cols

# save the dataframe into a dataframe

df_for_processing.to_csv("D:\All Data\desktop\Master Sheet Reports\LateDelivery From FMS - Dash Report\src\O2D FMS - Blame Department.csv")

# creating one level of columns from a multi level column index

list_cols = []
for i in df_for_processing.columns:
    join_cols = "-".join(j for j in i)
    list_cols.append(join_cols)
    
df_for_processing.columns=list_cols

# dropping the rows that are not necessary

df_for_processing.drop(index=[1,5],axis=0,inplace=True)




# __________________________________________CreatingCalculatedColumns__________________________________

# create column for order is delayed or not
df_for_processing["Order Delayed Status"]=""

for index,row in df_for_processing.iterrows():
    if((df_for_processing.loc[index,"Received pcs in Warehouse 2-Time Delay"]!="")&("-" not in df_for_processing.loc[index,"Received pcs in Warehouse 2-Time Delay"])):
      df_for_processing.loc[index,"Order Delayed Status"]="Order Delayed"
    else:
        df_for_processing.loc[index,"Order Delayed Status"]="Order on time"

# creating a new column called Column Stuck In

df_for_processing["Column Stuck In"]=""

# iterate through the dataframe and fill the Department Stuck In column

for index, row in df_for_processing.iterrows():
    if (df_for_processing.loc[index,'Received pcs in Warehouse 2-Time Delay']=="") | ("-" not in df_for_processing.loc[index,'Received pcs in Warehouse 2-Time Delay']):
        
        # find the list of columns that have the Time Delay word in it
        time_delay_col = [i for i in df.columns if "Delay" in i]
        
        # Remove columns that can never be in a delay
        non_delay_columns = ["Stock Availability in Pretture-Time Delay",
                             "Stock Availability in PRO(production)-Time Delay",
                             "Raise PRO(production) if not available anywhere-Time Delay",
                             "Alteration needed?-Time Delay"]
        
        # modifying the time delay column after removing the non delay column
        time_delay_col=[i for i in time_delay_col if i not in non_delay_columns]
        
        for i in time_delay_col:
            if df_for_processing.loc[index,i]!="":
                df_for_processing.loc[index,"Column Stuck In"]=i
                break
                
# creating the Blame Department from Column Stuck In column

df_for_processing["Blame Department"]=df_for_processing["Column Stuck In"].apply(lambda x: x.split("-")[0])

df_for_processing.fillna("no information",inplace=True)

#saving the column to csv
df_for_processing.to_csv(r"D:\All Data\desktop\Master Sheet Reports\O2D FMS - Blame Department.csv")

# ___________________________________DashAPP_____________________________________________________________

# Creating the Dash app

from dash import Dash, dcc, Input, Output, callback, html,dash_table
import dash_cool_components
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import dash_bootstrap_components as dbc


app = Dash(__name__)
server = app.server

app.layout=html.Div([
    
    html.H1("Late Delivery Report"),
    html.H3("Select dates for Date of Delivery Client"),
    dcc.Dropdown(options=df_for_processing['Create Order-DOD ( Client )'].unique(),value=["10/07/2023","11/07/2023"],id="input-1",multi=True),
        
    html.H3(id="fill-value"),
    html.H3(id="fill-value2"),
    
    html.H3("Select parameters for visualization"),
    dcc.Dropdown(options=['Order Delayed Status','Blame Department','Create Order-Barcode no.',"Create Order-Pretture no.",
                         "Create Order-Department"],value=['Order Delayed Status'],id="input-2",multi=True),
    dcc.Graph(id="graph-1"),
    
])

@callback(

    Output("fill-value","children"),
    Output("fill-value2","children"),
    Output("graph-1","figure"),
    Input("input-1","value"),
    Input("input-2","value"),
)

def update_graph(value1,value2):
    # calculate the number of orders for input dates
    
    df_for_processing_orders=df_for_processing[df_for_processing['Create Order-DOD ( Client )'].isin(value1)]
    number_of_orders = df_for_processing_orders.shape[0]
    
    # calculate the number of orders delayed
    df_for_processing_orders_delayed = df_for_processing_orders[df_for_processing_orders["Order Delayed Status"]=="Order Delayed"]
    number_of_delayed_orders=df_for_processing_orders_delayed.shape[0]
    
    dist_orders_delayed = f"Total number of orders is: {number_of_orders}, number of orders delayed is: {number_of_delayed_orders}"
    
    percentage_order_delayed=(number_of_delayed_orders/number_of_orders)*100
    dist_orders_delayed2 = "Percentage of orders delayed: {0:.2f}%".format(percentage_order_delayed)
    
    # making the sunburst plot for category delay and filling empty values 
    sunburst = df_for_processing_orders.replace("","no data")    
    fig = px.sunburst(sunburst,path=value2,height=700,width=1000,color_discrete_sequence=["brown","purple"])
    fig.update_traces(textinfo="label+value+percent parent")
    
    return dist_orders_delayed,dist_orders_delayed2,fig

if __name__=="__main__":
    app.run(debug=True)
    