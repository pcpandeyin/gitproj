import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import pandas as pd
import plotly.express as px
import datetime

# Read the Excel file
def load_data():
    df = pd.read_excel('c://pcp/sales_data.xlsx', sheet_name='Sales')
    df['Date'] = pd.to_datetime(df['Date'])
    return df

# Create Dash app
app = dash.Dash(__name__)
app.title = "Sales Dashboard"

app.layout = html.Div([
    html.H1("📊 Sales Dashboard", style={'textAlign': 'center'}),

    dcc.Interval(
        id='interval-component',
        interval=10*1000,  # Refresh every 10 seconds
        n_intervals=0
    ),

    html.Div([
        html.H3("Total Sales Summary"),
        html.Div(id='summary-div'),
    ], style={'marginBottom': 30}),

    html.Div([
        dcc.Graph(id='sales-graph')
    ]),

    html.Div([
        html.H3("Tabular Data"),
        dash_table.DataTable(
            id='sales-table',
            columns=[{'name': i, 'id': i} for i in ['Date', 'Region', 'Sales']],
            style_table={'overflowX': 'auto'},
            style_cell={'textAlign': 'left'},
        )
    ])
])

@app.callback(
    [Output('summary-div', 'children'),
     Output('sales-graph', 'figure'),
     Output('sales-table', 'data')],
    Input('interval-component', 'n_intervals')
)
def update_dashboard(n):
    df = load_data()

    # Summary
    total_sales = df['Sales'].sum()
    avg_sales = df['Sales'].mean()
    last_updated = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    summary = html.Ul([
        html.Li(f"Total Sales: ${total_sales}"),
        html.Li(f"Average Sales: ${avg_sales:.2f}"),
        html.Li(f"Last Updated: {last_updated}")
    ])

    # Graph
    fig = px.line(df, x='Date', y='Sales', color='Region', markers=True, title="Sales Over Time")

    return summary, fig, df.to_dict('records')

if __name__ == '__main__':
    app.run(debug=True)
