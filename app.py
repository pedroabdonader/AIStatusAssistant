from flask import Flask, request, send_file, render_template, jsonify
import os
import json
import requests
from openai import AzureOpenAI
import pandas as pd
import ast
from datetime import datetime
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from io import BytesIO
import pytz
 
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
            "content": "You are a status summarization assistant. I want the following information per Workstream: Workstream (Workstream name), Status (either: Done, On Time, At Risk, Late, Cancelled), Achievements (One short sentence summarizing the weekly achievements. If available, please include dates of when things were completed), Next Steps(One short sentence summarizing the next steps. If available, please include dates of when things are expected to get done), and Expected End Date (If available, when each workstream should be complete)."
            "Additionally, I want the following with information from all workstreams: Title (Should include most key information across workstreams and include the date), Description (A summary description of all status update across all workstreams), Key Decisions (Key decisions made during the status update period), and Issues/Risks (Issues and risks identified in the update)."
        },
        {"role": "user", "content": f"Here are my notes as of today {datetime.now(tz=pytz.timezone('America/New_York')).strftime('%Y-%m-%d')}: {user_input}"}
    ]

    response_format = {
        "type" : "json_schema",
        "json_schema" : {
    "name": "weekly_status_update",
    "schema": {
        "type": "object",
        "properties": {
        "title": {
            "type": "string",
            "description": "The title of the status update."
        },
        "description": {
            "type": "string",
            "description": "A summary description of the status update across all workstreams."
        },
        "updates": {
            "type": "array",
            "description": "List of workstreams and their statuses.",
            "items": {
            "type": "object",
            "properties": {
                "Workstream": {
                "type": "string",
                "description": "Name of the workstream."
                },
                "Status": {
                "type": "string",
                "description": "Current status of the workstream."
                },
                "Achievements": {
                "type": "array",
                "description": "Achievements made in the workstream.",
                "items": {"type": "string"}
                },
                "Next Steps": {
                "type": "array",
                "description": "Next steps planned for the workstream.",
                "items": {"type": "string"}
                },
                "Planned End Date": {
                "type": "string",
                "description": "When the workstream is planned to end."
                }
            },
            "required": [
                "Workstream",
                "Status",
                "Achievements",
                "Next Steps",
                "Planned End Date"
            ],
            "additionalProperties": False
            }
        },
        "key_decisions": {
            "type": "array",
            "description": "Key decisions made during the status update period.",
            "items": {
            "type": "string"
            }
        },
        "issues_risks": {
            "type": "array",
            "description": "Issues and risks identified in the update.",
            "items": {
            "type": "string"
            }
        }
        },
        "required": [
        "title",
        "description",
        "updates",
        "key_decisions",
        "issues_risks"
        ],
        "additionalProperties": False
    },
    "strict": True
    }
    }



    response = client.chat.completions.create(
        stream=False,
        messages=messages,
        response_format=response_format,
        max_tokens=4096,
        temperature=0,
        top_p=1.0,
        model=deployment,
    )

    response_dict = ast.literal_eval(response.choices[0].message.content)

    print(response_dict)  # Debugging line
    print(type(response_dict))  # Debugging line
    return response_dict

def createDf(data):
    df = pd.DataFrame(data)
    return df

def populate_powerpoint_template(df, title, description, key_decisions, issues_risks):
    presentation = Presentation("template.pptx")  # Ensure you have a template.pptx file

    #Get the cover slide
    slide = presentation.slides[0]

    for shape in slide.shapes:
        if shape.name == "Cover Title":
            shape.text = title
            for paragraph in shape.text_frame.paragraphs:
                paragraph.font.size = Pt(32)
                paragraph.font.color.rgb = RGBColor(255, 255, 255)


    # Get the Status slide
    slide = presentation.slides[1]

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

            if isinstance(df.iat[row_idx, col_idx],list):   # Check if the cell contains a list, if it does, join the list into a string in different lines (This is for the Achievements and Next Steps columns)
                cell.text = str("\n".join(df.iat[row_idx, col_idx]))

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
            # Set font size and color for all paragraphs in the cell
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(11)
                paragraph.font.color.rgb = RGBColor(0, 0, 0)  # Black font


    # Populate other shapes with the title, description, key decisions, and issues/risks
    # Format Key Decisions Box with bullets
    for shape in slide.shapes:
        if shape.name == "Description":
            shape.text = description
            for paragraph in shape.text_frame.paragraphs:
                paragraph.font.size = Pt(12)
                paragraph.font.color.rgb = RGBColor(0, 0, 0)    
        elif shape.name == "Title":
            shape.text = title
            for paragraph in shape.text_frame.paragraphs:
                paragraph.font.size = Pt(28)
                paragraph.font.color.rgb = RGBColor(0, 0, 0)
        elif shape.name == "Key Decisions Box":
            shape.text = "\n".join(key_decisions)
            for paragraph in shape.text_frame.paragraphs:
                paragraph.font.size = Pt(12)
                paragraph.font.color.rgb = RGBColor(0, 0, 0)
        elif shape.name == "Issues/Risks Box":
            shape.text = "\n".join(issues_risks)
            for paragraph in shape.text_frame.paragraphs:
                paragraph.font.size = Pt(12)
                paragraph.font.color.rgb = RGBColor(0, 0, 0)



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
        df = createDf(result['updates'])
        
        # Include column names in the response
        response_data = {
            'columns': df.columns.tolist(),  # Get the column names
            'data': df.to_dict(orient='records')  # Get the data
        }
        
        # Save the PowerPoint file to a BytesIO object
        ppt_stream = populate_powerpoint_template(df, result['title'], result['description'], result['key_decisions'], result['issues_risks'])

        # Store the stream in a global variable or use a session to access it later
        global ppt_file_stream
        ppt_file_stream = ppt_stream

        # Return the DataFrame as JSON
        return jsonify(response_data)

    return render_template('index.html')

@app.route('/download', methods=['GET'])
def download():
    # Send the PowerPoint file for download
    return send_file(ppt_file_stream, as_attachment=True, download_name=f"Status_Report_{datetime.now(tz=pytz.timezone('America/New_York')).strftime('%Y-%m-%d')}.pptx")


if __name__ == '__main__':
    app.run(debug=True)
