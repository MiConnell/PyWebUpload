# based on plotly's dash upload component
# which I modified to read data and insert/update
# https://dash.plot.ly/dash-core-components/upload

# library imports
import base64
import datetime
import io

import dash
import dash_auth
from dash.dependencies import Input, Output
import dash_core_components as dcc
import dash_html_components as html
import dash_table_experiments as dt

import pandas as pd
# import pyodbc

# define database connection
# conn = pyodbc.connect(
#     r'DRIVER=SQL Server;'
#     r'SERVER= ;'
#     r'DATABASE= ;'
#     r'Trusted_Connection=yes;'
# )

# # call cursor
# cursor = conn.cursor()

# username and password for login to site
VALID_USERNAME_PASSWORD_PAIRS = [
    ['hello', 'world']
]

# company logo
img = ''

# call app
app = dash.Dash()
# app title
app.title = "Web Upload"
# app authentication using above username/password pairs
auth = dash_auth.BasicAuth(
    app,
    VALID_USERNAME_PASSWORD_PAIRS
)

# app config
app.scripts.config.serve_locally = True

# main layout
app.layout = html.Div([
    # upload div
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
        multiple=True
    ),
    # output div
    html.Div(id='output-data-upload'),
    html.Div(dt.DataTable(rows=[{}]), style={'display': 'none'}),
    html.Div([
        html.Div([
            html.Img(src=img, height='60', width='190')
        ], className="twelve columns"),
        html.Div([
            html.H5('Upload')
        ])
    ], className="row gs-header")
])

# define path for error log
out = 'Drive/Error Log.txt'
now = datetime.datetime.now()
nowFormat = now.strftime("%m/%d/%Y %X")

# main function to read spreadsheet


def parse_contents(contents, filename, date):
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)

    if 'xls' in filename:
        # Assume that the user uploaded an excel file
        uploadSheet = pd.read_excel(io.BytesIO(decoded))
        for a, b in zip(
                uploadSheet['Column_1'],      # a
                uploadSheet['Column_2']):     # b
            # query to verify that the spreadsheet contains valid data
            ProductionQuery = pd.io.sql.read_sql(
                f"""select * from table where column = '{str(a)}'""", conn)
            # if data not recognized, write to the error log and return a message saying so
            if ProductionQuery.empty:
                with open(out, 'a+') as f:
                    f.write(
                        f'Production Error: Column_1 {str(a)} Not Recognized At Time {str(nowFormat)} \n')
                return html.Div([
                    html.H5(
                        f'Production Error: Column_1 {str(a)} Not Recognized'),
                    html.Hr(),  # horizontal line

                ])
            try:
                # update query to run based on data on the spreadsheet
                updateQuery = f"""UPDATE Table SET Column = '{str(b)}'
                                  WHERE Column_1 = '{str(a)}'
                                     """

                cursor.execute(updateQuery)
                conn.commit()
            # error handling
            except Exception as e:
                # if error, write to error log and return a message saying so
                with open(out, 'a+') as f:
                    f.write(
                        f'Production Error: {str(e)} At Time {str(nowFormat)} \n')
                return html.Div([
                    html.H5(f'Production Error: {e}'),
                    html.Hr(),  # horizontal line

                ])
    else:
        # if file is not in .xlsx format
        with open(out, 'a+') as f:
            f.write(
                f'Production Error: File must be .xlsx format at time {nowFormat} \n')
        return html.Div([
            html.H5(f'Production Error: File must be .xlsx format'),
            html.Hr(),  # horizontal line

        ])
    # if successful, return a success message and the date/time
    return html.Div([
        html.H5('Upload Complete'),
        html.H6(datetime.datetime.fromtimestamp(date)),
        html.Hr(),  # horizontal line
    ])

# main callback


@app.callback(Output('output-data-upload', 'children'),
              [Input('upload-data', 'contents'),
               Input('upload-data', 'filename'),
               Input('upload-data', 'last_modified')])
# function to update output
def update_output(list_of_contents, list_of_names, list_of_dates):
    try:
        if list_of_contents is not None:
            children = [
                parse_contents(c, n, d) for c, n, d in
                zip(list_of_contents, list_of_names, list_of_dates)]
            return children
    except Exception as e:
        print(e)
        with open(out, 'a+') as f:
            f.write(f'Production Error: {str(e)} At Time {str(nowFormat)} \n')
        return html.Div([
            html.H5(f'Production Error: {e}'),
            html.Hr(),  # horizontal line
        ])


# styling
app.css.append_css({
    'external_url': 'https://codepen.io/chriddyp/pen/bWLwgP.css'
})

# run the app
if __name__ == '__main__':
    app.run_server(debug=True, host='0.0.0.0', port=8052
                   )
