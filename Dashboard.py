import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import pyodbc
from datetime import datetime, date, timedelta

con = pyodbc.connect(Trusted_Connection='no',
                     driver='{SQL Server}',
                     server='192.168.15.32',
                     database='Alpha_Live',
                     UID='pladis_dba',
                     PWD='BigFlats')
cursor = con.cursor()

shift_length_hr = 12

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

line_speed_df = pd.read_sql('SELECT [Timestamp], [Line Speed (Rows/Min)] FROM [dbo].[KGM 080]', con,
                            index_col='Timestamp')
sig1_profile_df = pd.read_sql('SELECT [Timestamp],[Workstation_Location],[Pieces Received at 3D Profile Station],'
                              '[Rejected Count for Too Long], [Rejected Count for Too Short], [Rejected Count for Too '
                              'High],[Rejected Count for Too Wide] FROM [Alpha_Live].[dbo].[HCM 274]', con,
                              index_col='Timestamp')
sig2_profile_df = pd.read_sql('SELECT [Timestamp],[Workstation_Location],[Pieces Received at 3D Profile Station],'
                              '[Rejected Count for Too Long],[Rejected Count for Too Short], '
                              '[Rejected Count for Too High],[Rejected Count for Too Wide] '
                              'FROM [Alpha_Live].[dbo].[HCM 273]', con, index_col='Timestamp')
sig1_profile_transpose = sig1_profile_df.loc[:,
                         [x for x in sig1_profile_df.columns if x not in ['Workstation_Location',
                                                                          'Pieces Received at 3D Profile Station']]].T
sig2_profile_transpose = sig2_profile_df.loc[:,
                         [x for x in sig2_profile_df.columns if x not in ['Workstation_Location',
                                                                          'Pieces Received at 3D Profile Station']]].T

sig1_wrapper_df = pd.read_sql('SELECT [Timestamp],[Workstation_Location],[Packages into Wrapper], '
                              '[Double Packages Rejected], [Packages Rejected Waste],[Packages Rejected Too Short], '
                              '[Packages Rejected Too Long],[Packages Rejected Empty], '
                              '[Packages Rejected Out of Registration],[Packages Rejected During Splice], '
                              '[Packages Rejected Supplier Splice] FROM [Alpha_Live].[dbo].[HCM 274]', con, index_col='Timestamp')
sig2_wrapper_df = pd.read_sql('SELECT [Timestamp],[Workstation_Location],[Packages into Wrapper], '
                              '[Double Packages Rejected], [Packages Rejected Waste],[Packages Rejected Too Short], '
                              '[Packages Rejected Too Long],[Packages Rejected Empty], '
                              '[Packages Rejected Out of Registration],[Packages Rejected During Splice], '
                              '[Packages Rejected Supplier Splice] FROM [Alpha_Live].[dbo].[HCM 273]', con, index_col='Timestamp')
sig1_wrapper_transpose = sig1_wrapper_df.loc[:, [x for x in sig1_wrapper_df.columns if x not in
                                                 ['Workstation_Location', 'Packages into Wrapper']]].T
sig2_wrapper_transpose = sig2_wrapper_df.loc[:, [x for x in sig2_wrapper_df.columns if x not in
                                                 ['Workstation_Location', 'Packages into Wrapper']]].T

turtle_size_df = pd.read_sql('SELECT [Date],[SKU],[Width (mm)], [Length (mm)], [Height (mm)] FROM '
                             '[QA_Warehouse].[dbo].[QA_3D_Profile]', con).dropna()

line_speed_fig = px.line(line_speed_df, x=line_speed_df.index, y='Line Speed (Rows/Min)')
line_speed_fig.update_xaxes(rangeslider_visible=True)
line_speed_fig.update_layout(transition_duration=500)
line_speed_fig.update_layout(margin={'l': 40, 'b': 40, 't': 10, 'r': 0}, hovermode='closest')

app.layout = html.Div([html.H1(children='Big Flats Alpha Line'),
                       html.Div(['Select Date: ',
                                 dcc.DatePickerSingle(id='input-date', min_date_allowed=line_speed_df.index.min(),
                                                      max_date_allowed=line_speed_df.index.max(),
                                                      initial_visible_month=date(2021, 3, 1),
                                                      date=date(2021, 3, 6)),
                                 dcc.RadioItems(
                                     id='input-shift',
                                     options=[{'label': i, 'value': i} for i in ['Day', 'Night']],
                                     value='Day',
                                     labelStyle={'display': 'inline-block'}
                                 )
                                 ]),
                       html.Div([dcc.Graph(figure=line_speed_fig)]),
                       html.Div([dcc.Graph(id='sig1-profile-waterfall')],
                                style={'width': '32%', 'display': 'inline-block'}),
                       html.Div([dcc.Graph(id='sig2-profile-waterfall')],
                                style={'width': '32%', 'display': 'inline-block', 'padding': '0px 10px'}),
                       html.Div([dcc.Graph(id='turtle-size-histogram')],
                                style={'width': '32%', 'display': 'inline-block', 'padding': '0px 10px'}),
                       html.Div([dcc.Graph(id='sig1-wrapper-waterfall')],
                                style={'width': '49%', 'display': 'inline-block'}),
                       html.Div([dcc.Graph(id='sig2-wrapper-waterfall')],
                                style={'width': '49%', 'display': 'inline-block', 'padding': '0px 10px'}),

                       ])


@app.callback(
    Output(component_id='sig1-profile-waterfall', component_property='figure'),
    Output(component_id='sig2-profile-waterfall', component_property='figure'),
    Output(component_id='turtle-size-histogram', component_property='figure'),
    Output(component_id='sig1-wrapper-waterfall', component_property='figure'),
    Output(component_id='sig2-wrapper-waterfall', component_property='figure'),
    Input(component_id='input-date', component_property='date'),
    Input(component_id='input-shift', component_property='value')
)
def update_output_div(date_value, shift_value):
    if date_value == datetime.strftime(datetime.today(), '%m/%d/%Y'):
        date_time_value = sig1_profile_df['Timestamp'].max()
    else:
        if shift_value == 'Day':
            date_time_value = date_value + ' 05:45'
        else:
            date_time_value = datetime.strftime(datetime.strptime(date_value, '%Y-%m-%d') + timedelta(days=1),
                                                '%Y-%m-%d') + ' 05:45'
    partial_yaxis = max(
        sig1_profile_transpose[date_time_value].to_list() + sig2_profile_transpose[date_time_value].to_list())

    sig1_profile_waterfall = go.Figure(
        go.Waterfall(orientation='v', measure=['relative', 'relative', 'relative', 'relative'],
                     x=['Too Long', 'Too Short', 'Too High', 'Too Wide'],
                     y=sig1_profile_transpose[date_time_value].to_list(),
                     connector={'line': {'color': 'rgb(63, 63, 63)'}},
                     increasing={'marker': {'color': 'rgb(240, 136, 146)'}}))
    sig1_profile_waterfall.update_layout(title='Sig 1 Turtle Blow Offs')
    sig1_profile_waterfall.update_yaxes(range=[0, partial_yaxis])
    sig1_profile_waterfall.update_layout(height=225, margin={'l': 2, 'b': 2, 'r': 2, 't': 40})

    sig2_profile_waterfall = go.Figure(
        go.Waterfall(orientation='v', measure=['relative', 'relative', 'relative', 'relative'],
                     x=['Too Long', 'Too Short', 'Too High', 'Too Wide'],
                     y=sig2_profile_transpose[date_time_value].to_list(),
                     connector={'line': {'color': 'rgb(63, 63, 63)'}},
                     increasing={'marker': {'color': 'rgb(240, 136, 146)'}}))
    sig2_profile_waterfall.update_layout(title='Sig 2 Turtle Blow Offs')
    sig2_profile_waterfall.update_yaxes(range=[0, partial_yaxis])
    sig2_profile_waterfall.update_layout(height=225, margin={'l': 2, 'b': 2, 'r': 2, 't': 40})

    turtle_size_histogram_filtered_df = turtle_size_df[turtle_size_df['Date'] > date_value]
    turtle_size_histogram = px.histogram(turtle_size_histogram_filtered_df, x='Length (mm)',
                                         histnorm='probability density')
    turtle_size_histogram.update_layout(title='Turtle Length')
    turtle_size_histogram.update_layout(height=225, margin={'l': 2, 'b': 2, 'r': 2, 't': 40})

    wrapper_yaxis = max(sig1_wrapper_transpose[date_time_value].to_list() + sig2_wrapper_transpose[date_time_value].to_list())

    sig1_wrapper_waterfall = go.Figure(
        go.Waterfall(orientation='v', measure=['relative', 'relative', 'relative', 'relative','relative', 'relative', 'relative', 'relative'],
                     x=['Double Packages', 'Waste', 'Too Short', 'Too Long', 'Empty', 'Out of Registration', 'During Splice', 'Supplier Splice'],
                     y=sig1_wrapper_transpose[date_time_value].to_list(),
                     connector={'line': {'color': 'rgb(63, 63, 63)'}},
                     increasing={'marker': {'color': 'rgb(240, 136, 146)'}}))
    sig1_wrapper_waterfall.update_layout(title='Sig 1 Wrapper Defects')
    sig1_wrapper_waterfall.update_yaxes(range=[0, wrapper_yaxis])
    sig1_wrapper_waterfall.update_layout(height=225, margin={'l': 2, 'b': 2, 'r': 2, 't': 40})

    sig2_wrapper_waterfall = go.Figure(
        go.Waterfall(orientation='v', measure=['relative', 'relative', 'relative', 'relative','relative', 'relative', 'relative', 'relative'],
                     x=['Double Packages', 'Waste', 'Too Short', 'Too Long', 'Empty', 'Out of Registration', 'During Splice', 'Supplier Splice'],
                     y=sig2_wrapper_transpose[date_time_value].to_list(),
                     connector={'line': {'color': 'rgb(63, 63, 63)'}},
                     increasing={'marker': {'color': 'rgb(240, 136, 146)'}}))
    sig2_wrapper_waterfall.update_layout(title='Sig 2 Wrapper Defects')
    sig2_wrapper_waterfall.update_yaxes(range=[0, wrapper_yaxis])
    sig2_wrapper_waterfall.update_layout(height=225, margin={'l': 2, 'b': 2, 'r': 2, 't': 40})

    return sig1_profile_waterfall, sig2_profile_waterfall, turtle_size_histogram, sig1_wrapper_waterfall, sig2_wrapper_waterfall


if __name__ == '__main__':
    app.run_server(debug=True)
