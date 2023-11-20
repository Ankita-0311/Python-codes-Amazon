import base64
import datetime
import io

import dash
from dash.dependencies import Input, Output, State
import dash_core_components as dcc
import dash_html_components as html
import dash_table
import dash_design_kit as ddk
import dash_core_components as dcc



import pandas as pd




app = dash.Dash(__name__)

app.layout = html.Div([
    dcc.Upload(
        id='upload-data',
        children=html.Div([
            'Drag and Drop or ',
            html.A('Select Files')
        ]),
        style={
            'width': '100%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
        },
        # Allow multiple files to be uploaded
        multiple=False
    ),
    html.Div(id='output-data-upload'),
])

html.Br(),


html.Div([
    dcc.RadioItems(
        id='Radiobutton',
        options=[
            {'label': 'Weekly Report', 'value': 1},
            {'label': 'Monthly Report', 'value': 2},
            {'label': 'Quaterly Report','value': 3}
    ],
    value=1
)  ])


def parse_contents(contents, filename, date):
    content_type, content_string = contents.split(',')

    decoded = base64.b64decode(content_string)
    try:
        if 'csv' in filename:
            # Assume that the user uploaded a CSV file
            df = pd.read_csv(
                io.StringIO(decoded.decode('utf-8')))
        elif 'xls' in filename:
            # Assume that the user uploaded an excel file
            df = pd.read_excel(io.BytesIO(decoded))
            
        df1= pd.read_excel(df, sheet_name=0)
        df1= pd.DataFrame(df1)
        
        
        df2= pd.read_excel(df, sheet_name=1)
        df2= pd.DataFrame(df2)
        
        df3= pd.read_excel(df, sheet_name=2 )
        df3= pd.DataFrame(df3)
        
        df4= pd.read_excel(df, sheet_name=3, index_col=None, usecols= "A:J")
        df4= pd.DataFrame(df4)
        
        df5= pd.read_excel(df, sheet_name=3, index_col=None, usecols= "M:U")
        df5= pd.DataFrame(df5)
        
        df5= df5.rename(columns={df5.columns[0]:'Name',
                                 df5.columns[1]:'Login',
                                 df5.columns[3]:'Tagged Volume',
                                 df5.columns[4]:'Target Volume',
                                 df5.columns[5]:'Actual Productivity',
                                 df5.columns[6]:'Total QA Sample',
                                 df5.columns[7]:'Total Errors',
                                 df5.columns[8]:'QA %'})
        
        df6= pd.read_excel(df, sheet_name=3, index_col=None, usecols= "X:AF")
        df6= pd.DataFrame(df6)
        
        df6= df6.rename(columns={df6.columns[0]:'Name',
                                 df6.columns[1]:'Login',
                                 df6.columns[3]:'Tagged Volume',
                                 df6.columns[4]:'Target Volume',
                                 df6.columns[5]:'Actual Productivity',
                                 df6.columns[6]:'Total QA Sample',
                                 df6.columns[7]:'Total Errors',
                                 df6.columns[8]:'QA %'})
        
        
        df7= pd.read_excel(df, sheet_name=4, index_col=None, usecols= "A:E")
        df7= pd.DataFrame(df7)
        
        df8= pd.read_excel(df, sheet_name=4, index_col=None, usecols= "G:K")
        df8= pd.DataFrame(df8)
        
        df8= df8.rename(columns={df8.columns[2]:'Name',
                                 df8.columns[1]:'Login',
                                 df8.columns[3]:'Actual Productivity',
                                 df8.columns[4]:'QA %'})
        
        df9= pd.read_excel(df, sheet_name=4, index_col=None, usecols= "M:Q")
        df9= pd.DataFrame(df9)
        
        df9= df9.rename(columns={df9.columns[2]:'Name',
                                 df9.columns[1]:'Login',
                                 df9.columns[3]:'Actual Productivity',
                                 df9.columns[4]:'QA %'})
        
        
        df10= pd.read_excel(df, sheet_name=5)
        df10= pd.DataFrame(df10)
        
    except Exception as e:
        print(e)
        return html.Div([
            'There was an error processing this file.'
        ])

    return html.Div([
        html.H5(filename),
        html.H6(datetime.datetime.fromtimestamp(date)),

        dash_table.DataTable(
            data=df.to_dict('records'),
            columns=[{'name': i, 'id': i} for i in df.columns]
        ),

        html.Hr(),  # horizontal line

        # For debugging, display the raw contents provided by the web browser
        html.Div('Raw Content'),
        html.Pre(contents[0:200] + '...', style={
            'whiteSpace': 'pre-wrap',
            'wordBreak': 'break-all'
        })
    ])


@app.callback(Output('output-data-upload', 'children'),
              Input('upload-data', 'contents'),
              State('upload-data', 'filename'),
              State('upload-data', 'last_modified'))
def update_output(list_of_contents, list_of_names, list_of_dates):
    if list_of_contents is not None:
        children = [
            parse_contents(c, n, d) for c, n, d in
            zip(list_of_contents, list_of_names, list_of_dates)]
        return children



if __name__ == '__main__':
    app.run_server(debug=True)