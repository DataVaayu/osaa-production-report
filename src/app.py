import pandas as pd
from dash import Dash, html, dcc, Input, Output, callback
import plotly.express as px

# Reading the excel file
osaa_production_report = pd.ExcelFile(r"Osaa Production Report '23.xlsx")

# listing the sheets by their names
sheet_list = [i for i in osaa_production_report.sheet_names]

print(sheet_list)

print(len(sheet_list))

# creating an empty list to store the dataframes
total_data = []

for i in range(len(sheet_list)):
    data_month = pd.read_excel(osaa_production_report,sheet_name=sheet_list[i])

    stock_data = data_month.iloc[0:,0:4]
    stock_data.columns=stock_data.iloc[0]
    stock_data=stock_data.drop(stock_data.index[0],axis=0)
    stock_data["INVENTORY TYPE"]="STOCK"
    stock_data=stock_data.dropna(subset="DESIGN CODE")
    print(stock_data.shape)


    client_data = data_month.iloc[0:,6:10]
    client_data.columns=client_data.iloc[0]
    client_data=client_data.drop(client_data.index[0],axis=0)
    client_data["INVENTORY TYPE"]="CLIENT"
    client_data=client_data.dropna(subset="DESIGN CODE")
    print(client_data.shape)


    monthly_dataframe = pd.concat([stock_data,client_data],axis=0)
    monthly_dataframe = monthly_dataframe.reset_index()
    monthly_dataframe.drop("index",axis=1,inplace=True)
    print(monthly_dataframe.shape)
    total_data.append(monthly_dataframe)
    

production_order_data_osaa = pd.concat([i for i in total_data],axis=0)

print(production_order_data_osaa)
osaa_production_report.close()

production_order_data_osaa["ORDER DATE"]=pd.to_datetime(production_order_data_osaa["ORDER DATE"])
production_order_data_osaa["Month Name"]=production_order_data_osaa["ORDER DATE"].dt.month_name()
production_order_data_osaa["Week Number"]=production_order_data_osaa["ORDER DATE"].dt.day.apply(lambda x: (x-1)//7 + 1)
production_order_data_osaa.fillna("no data",inplace=True)


# ........................................................................................................................


# creating the app


app = Dash(__name__)
server=app.server

app.layout = html.Div([
    
    html.H1("Osaa Production Report"),
    html.H3("Choose Parameters"),
    dcc.Dropdown(options = production_order_data_osaa.columns,value=["Month Name"],id="dropdown-1",multi=True),
    dcc.Graph(id="graph-1")
])

@callback(
        Output("graph-1","figure"),
        Input("dropdown-1","value")
)

def update_graph(value1):
    columns=value1
    fig = px.sunburst(production_order_data_osaa,path=columns,width=1000,height=700)
    fig.update_traces(textinfo="label+value")
    
    return fig

if __name__ == "__main__":
    app.run(debug=True)