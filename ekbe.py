import pandas as pd
import plotly.graph_objects as go
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import webbrowser
import os

# Load Excel data
file_path = 'EKBE Table Data1.XLSX'
df = pd.read_excel(file_path)

# Select required columns
data = df[['Material Document', 'Movement Type', 'Material Doc.Item', 'Purchasing Document', 'PO History Category']]

# Count occurrences
count_103 = data[data['Movement Type'] == 103].shape[0]
count_105 = data[data['Movement Type'] == 105].shape[0]
count_done = data[data['PO History Category'] == 'D'].shape[0]  # Assuming 'D' indicates done

# Output files
file_103 = "Materials_without_105.xlsx"
file_105 = "105_Movement_VRe_Not_Done.xlsx"
file_done = "Filtered_Purchasing_Documents.xlsx"

# Save Excel files (ensure they exist)
materials_not_105 = data[(data['Movement Type'] == 103) & ~(data['Material Document'].isin(data[data['Movement Type'] == 105]['Material Document']))]
materials_not_105.to_excel(file_103, index=False)

movement_105_vre_not_done = data[(data['Movement Type'] == 105) & (data['PO History Category'] != 'T')]
movement_105_vre_not_done.to_excel(file_105, index=False)

filtered_data = data[(data['PO History Category'] == 'T') & (data['PO History Category'] != 'Q')].drop(columns=['Movement Type'])
filtered_data.to_excel(file_done, index=False)

# Define pie chart data
labels = ['Movement Type 103', 'Movement Type 105', 'PO History Done']
sizes = [count_103, count_105, count_done]
file_paths = [file_103, file_105, file_done]  # File paths mapped to segments

# Dash App
app = dash.Dash(__name__)

app.layout = html.Div([
    html.H1("Interactive Pie Chart for Movement Types"),
    dcc.Graph(
        id='pie-chart',
        figure=go.Figure(data=[go.Pie(
            labels=labels, 
            values=sizes,
            hoverinfo="label+percent",
            textinfo="percent",
        )])
    ),
    html.Div(id='output-div')
])

# Callback to detect clicks and open file
@app.callback(
    Output('output-div', 'children'),
    Input('pie-chart', 'clickData')
)
def display_click_data(clickData):
    if clickData:
        label_clicked = clickData['points'][0]['label']
        
        # Get corresponding file
        file_map = dict(zip(labels, file_paths))
        file_to_open = file_map.get(label_clicked)

        if file_to_open and os.path.exists(file_to_open):
            webbrowser.open(file_to_open)  # Opens the file
            return f"Opening file: {file_to_open}"
        else:
            return "File not found!"

    return "Click on a segment to open the corresponding file."

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
