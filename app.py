from flask import Flask, request, send_file, render_template, jsonify
import os
import json
import requests
from openai import AzureOpenAI
import pandas as pd
import ast
from datetime import datetime
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from io import BytesIO

app = Flask(__name__)

# Azure OpenAI client configuration
endpoint = os.environ.get('AZURE_ENDPOINT')
model_name = "gpt-4o"
deployment = "gpt-4o"

api_key = os.environ.get('AZURE_KEY')
api_version = "2024-12-01-preview"

client = AzureOpenAI(
    api_version=api_version,
    azure_endpoint=endpoint,
    api_key=api_key,
)

def get_response(user_input):
    messages = [
        {
            "role": "system",
            "content": "You are a status summarization assistant that will only respond with a string for a python dictionary and never with anything else. The string dictionary will then become a pandas dataframe with Four columns: Workstream (Workstream name), Status (either: Done, On Time, At Risk, Late, Cancelled), Achievements (One short sentence summarizing the weekly achievements. If available, please include dates of when things were completed), Next Steps(One short sentence summarizing the next steps. If available, please include dates of when things are expected to get done), and Expected End Date (When The initiative should be complete). Don't say anything else except the string dictionary. DO NOT GIVE ME IN A PYTHON CODE FORMAT, GIVE ME A STRING"
        },
        {"role": "user", "content": f"Here are my notes: {user_input}"}
    ]

    response = client.chat.completions.create(
        stream=False,
        messages=messages,
        max_tokens=4096,
        temperature=0,
        top_p=1.0,
        model=deployment,
    )

    response_content = response.choices[0].message.content
    print(response_content)  # Debugging line

    first_bracket = response_content.find('{')
    last_bracket = response_content.rfind('}')

    if first_bracket != -1 and last_bracket != -1:
        response_content = response_content[first_bracket:last_bracket + 1]

    response_dict = ast.literal_eval(response_content)
    return response_dict

def createDf(data):
    df = pd.DataFrame(data)
    return df

def populate_powerpoint_template(df):
    presentation = Presentation("template.pptx")  # Ensure you have a template.pptx file

    # Get the first slide
    slide = presentation.slides[0]

    # Define the position and size of the table
    left = Inches(0.53)
    top = Inches(2.25)
    # Define column widths
    col_widths = [Inches(1.14), Inches(0.8), Inches(2.64), Inches(2.64), Inches(1.79)]
    rows, cols = df.shape

    # Add a table to the slide
    table = slide.shapes.add_table(rows + 1, cols, left, top, sum(col_widths), Inches(0.5 + rows * 0.5)).table

    # Set column widths
    for i, width in enumerate(col_widths):
        table.columns[i].width = width

    # Set header row height
    table.rows[0].height = Inches(0.3)  # Set header row height to 0.3 inches
    
    # Set header row with background color and font size
    for col_idx, col_name in enumerate(df.columns):
        cell = table.cell(0, col_idx)
        cell.text = str(col_name)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(255, 250, 204)  # #FFFACC
        cell.text_frame.paragraphs[0].font.size = Pt(12)

    # Populate the table with DataFrame data and format other rows
    for row_idx in range(rows):
        for col_idx in range(cols):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = str(df.iat[row_idx, col_idx])
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(250, 250, 250)  #White background
            # Conditional formatting for the second column (status)
            if col_idx == 1:  # Check if it's the second column
                status = str(df.iat[row_idx, col_idx])
                if status == "On Time":
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(210, 244, 220)  # Light Green
                elif status == "At Risk":
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(255, 240, 204)  # Light Yellow
                elif status == "Late":
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(255, 217, 215)  # Light Red
                elif status == "Cancelled":
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(191, 191, 191)  # Light Gray
                elif status == "Done":
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(208, 232, 250)  # Light Blue
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(250, 250, 250)  # Default to white background
            cell.text_frame.paragraphs[0].font.size = Pt(10)
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)  # Black font

    # Save the presentation to a BytesIO object
    ppt_stream = BytesIO()
    presentation.save(ppt_stream)
    ppt_stream.seek(0)  # Move to the beginning of the stream
    return ppt_stream


def set_shape_format(shape):
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Arial'
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(0, 0, 0)

def set_shape_fill_color(shape, status):
    color_mapping = {
        "On Time": RGBColor(0, 176, 80),
        "At Risk": RGBColor(255, 192, 0),
        "Late": RGBColor(192, 0, 0),
        "Done": RGBColor(0, 112, 192),
        "Cancelled": RGBColor(125, 125, 125)
    }
    fill_color = color_mapping.get(status, RGBColor(255, 255, 255))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        notes = request.form['notes']
        result = get_response(notes)
        df = createDf(result)
        
        # Include column names in the response
        response_data = {
            'columns': df.columns.tolist(),  # Get the column names
            'data': df.to_dict(orient='records')  # Get the data
        }
        
        # Save the PowerPoint file to a BytesIO object
        ppt_stream = populate_powerpoint_template(df)

        # Store the stream in a global variable or use a session to access it later
        global ppt_file_stream
        ppt_file_stream = ppt_stream

        # Return the DataFrame as JSON
        return jsonify(response_data)

    return render_template('index.html')

@app.route('/download', methods=['GET'])
def download():
    # Send the PowerPoint file for download
    return send_file(ppt_file_stream, as_attachment=True, download_name=f"Status_Report_{datetime.now().strftime('%Y-%m-%d')}.pptx")


if __name__ == '__main__':
    app.run(debug=True)
