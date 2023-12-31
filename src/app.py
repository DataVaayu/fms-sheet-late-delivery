# ___________________CredentialsAndConnectingToGoogleSheets_____________________________________________________

from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pandas as pd
from datetime import datetime as dt
from datetime import timedelta

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

# df_for_processing.to_csv("D:\All Data\desktop\Master Sheet Reports\LateDelivery From FMS - Dash Report\src\O2D FMS - Blame Department.csv")

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



# find the list of columns that have the Time Delay word in it
time_delay_col = [i for i in df_for_processing.columns if "Delay" in i]
        
# Remove columns that can never be in a delay
# First creating a list of non-dely columns
non_delay_columns = ["Stock Availability in Pretture-Time Delay",
                             "Stock Availability in PRO(production)-Time Delay",
                             "Raise PRO(production) if not available anywhere-Time Delay",
                             "Alteration needed?-Time Delay"]
        
# modifying the time delay column after removing the non delay column
time_delay_col=[i for i in time_delay_col if i not in non_delay_columns]

# iterate through the dataframe and fill the Department Stuck In column

for index, row in df_for_processing.iterrows():
    if (df_for_processing.loc[index,'Received pcs in Warehouse 2-Time Delay']=="") | ("-" not in df_for_processing.loc[index,'Received pcs in Warehouse 2-Time Delay']):
        
              
        for i in time_delay_col:
            if df_for_processing.loc[index,i]!="":
                df_for_processing.loc[index,"Column Stuck In"]=i
                break
                
# creating the Blame Department from Column Stuck In column

df_for_processing["Blame Department"]=df_for_processing["Column Stuck In"].apply(lambda x: x.split("-")[0])

df_for_processing.fillna("no information",inplace=True)


# preprocessing the data of delivery clients by string and creating a new date column
# splitting the old column dates
df_for_processing["Create Order-DOD ( Client )"]=df_for_processing["Create Order-DOD ( Client )"].astype(str)
#print(df_for_processing["Create Order-DOD ( Client )"])

for index,row in df_for_processing.iterrows():
    try:
        df_for_processing.loc[index,"Date_Day"]=row["Create Order-DOD ( Client )"].split("/")[0]
        df_for_processing.loc[index,"Date_Month"]=row["Create Order-DOD ( Client )"].split("/")[1]
        df_for_processing.loc[index,"Date_Year"]=row["Create Order-DOD ( Client )"].split("/")[2]
    except:
        try:
            df_for_processing.loc[index,"Date_Day"]=row["Create Order-DOD ( Client )"].split("-")[0]
            df_for_processing.loc[index,"Date_Month"]=row["Create Order-DOD ( Client )"].split("-")[1]
            df_for_processing.loc[index,"Date_Year"]=row["Create Order-DOD ( Client )"].split("-")[2]
        except:
            print(row["Create Order-DOD ( Client )"])

# creating a new dod client date column
for index,row in df_for_processing.iterrows():
    df_for_processing.loc[index,"Create Order-DOD ( Client )_2"]="-".join([str(row["Date_Year"]),str(row["Date_Month"]),str(row["Date_Day"])])

# ___________________________________DashAPP_____________________________________________________________

# Creating the Dash app

from dash import Dash, dcc, Input, Output, callback, html,dash_table
import pandas as pd
import plotly.express as px
import dash_bootstrap_components as dbc


app = Dash(__name__)
server = app.server

app.layout=html.Div([
    
    html.H1("LATE DELIVERY REPORT",style={"text-align":"center","background-color":"rgb(207,239,249)"}),
    html.Br(),    

    #html.H3("Select dates for Date of Delivery Client"),
    #dcc.Dropdown(options=df_for_processing['Create Order-DOD ( Client )'].unique(),value=["10/07/2023","11/07/2023"],id="input-1",multi=True),

    # Range for Date of Delivery Client

    html.H3("Select Date Range for Date of Delivery of Client"),
    #Creating a datepicker range in the layout
    dcc.DatePickerRange(
        id="date-picker-range",
        calendar_orientation="horizontal",
        day_size=39,
        end_date_placeholder_text="Return",
        first_day_of_week=1,
        reopen_calendar_on_clear=True,
        clearable=True,
        number_of_months_shown=1,
        min_date_allowed=dt(2023, 6, 19),
        initial_visible_month=dt(2023, 7, 1),
        display_format='MMM Do, YY',
        month_format='MMMM, YYYY',
        minimum_nights=1,
        updatemode='singledate',
        start_date='2023-07-31',
        end_date="2023-08-05",
        
    ),

    html.Br(),

    # Listing the dates that are in the dates list selected from the DatePickerRange
    #html.H3(id="dates-between-start-end-dates"),

    html.H3(id="fill-value"),
    html.H3(id="fill-value2"),
    
    html.H3("Select parameters for visualization"),
    dcc.Dropdown(options=['Order Delayed Status','Blame Department','Create Order-Barcode no.',"Create Order-Pretture no.",
                         "Create Order-Department"],value=['Order Delayed Status'],id="input-2",multi=True),
    dcc.Graph(id="graph-1"),

    dash_table.DataTable([{"data":i,"id":i} for i in df_for_processing[["Create Order-Pretture no.","Create Order-Barcode no.","Create Order-Customer Name","Create Order-Department","Create Order-DOD ( Client )","Create Order-Design no.","Blame Department"]].columns],page_size=10,id="late-table")
    
])

@callback(

    Output("fill-value","children"),
    Output("fill-value2","children"),
    Output("graph-1","figure"),
    Output("late-table","data"),
    #Output("dates-between-start-end-dates","children"),
    #Input("input-1","value"),
    Input("input-2","value"),
    Input("date-picker-range","start_date"),
    Input("date-picker-range","end_date"),
)

def update_graph(value2,start_date, end_date):

    # create the date range with all dates from the DatePickerRange start and end dates
    #print(start_date,end_date)
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    dates_list=[]

    while start_date<=end_date:
        dates_list.append(start_date.strftime("%Y-%m-%d"))
        start_date +=timedelta(days=1)
    #print(dates_list)


    # calculate the number of orders for input dates
    #print(f"column dates: {df_for_processing['Create Order-DOD ( Client )'][df_for_processing['Create Order-DOD ( Client )'].isin(dates_list)]}")
    df_for_processing_orders=df_for_processing[df_for_processing['Create Order-DOD ( Client )_2'].isin(dates_list)]
    
    # Filtering out the dataframe from the start and end dates
    #df_for_processing_orders=df_for_processing.loc[start_date:end_date]

    # Calculation of number of orders

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

    # the dataframe for late table visualization
    df_for_processing2=df_for_processing_orders_delayed[["Create Order-Pretture no.","Create Order-Barcode no.","Create Order-Customer Name","Create Order-Department","Create Order-DOD ( Client )","Create Order-Design no.","Blame Department"]]
    
    return dist_orders_delayed,dist_orders_delayed2,fig,df_for_processing2.to_dict("records")

if __name__=="__main__":
    app.run(debug=True,port=8033)
    