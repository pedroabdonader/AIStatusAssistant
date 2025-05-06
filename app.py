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
            "content": "You are a status summarization assistant that will only respond with a string for a python dictionary and never with anything else. The string dictionary will then become a pandas dataframe with Four columns: Workstream name, Status (either: Done, On Time, At Risk, Late, Cancelled), Current Week achievements (One short sentence summarizing the weekly achievements. If available, please include dates of when things were completed), and next steps(One short sentence summarizing the next steps. If available, please include dates of when things are expected to get done). Don't say anything else except the string dictionary. DO NOT GIVE ME IN A PYTHON CODE FORMAT, GIVE ME A STRING"
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
    slide = presentation.slides[0]  # Access slide 2 (index 1)

    for i in range(len(df)):
        current_shape_name = f"Workstream {i + 1} Current"
        future_shape_name = f"Workstream {i + 1} Future"
        status_shape_name = f"Workstream {i + 1} Status"

        for shape in slide.shapes:
            if shape.name == current_shape_name:
                shape.text = df.iat[i, 2]  # Current Week achievements
                set_shape_format(shape)
            elif shape.name == future_shape_name:
                shape.text = df.iat[i, 3]  # Next steps
                set_shape_format(shape)
            elif shape.name == status_shape_name:
                set_shape_fill_color(shape, df.iat[i, 1])  # Status

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
        print (df)  # Debugging line
        # Save the PowerPoint file to a BytesIO object
        ppt_stream = populate_powerpoint_template(df)

        # Store the stream in a global variable or use a session to access it later
        global ppt_file_stream
        ppt_file_stream = ppt_stream

        # Return the DataFrame as JSON
        return jsonify(df.to_dict(orient='records'))

    return render_template('index.html')

@app.route('/download', methods=['GET'])
def download():
    # Send the PowerPoint file for download
    return send_file(ppt_file_stream, as_attachment=True, download_name=f"Status_Report_{datetime.now().strftime('%Y-%m-%d')}.pptx")


if __name__ == '__main__':
    app.run(debug=True)
