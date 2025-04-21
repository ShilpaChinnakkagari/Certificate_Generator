from flask import Flask, render_template, request, send_file
from PIL import Image, ImageDraw, ImageFont
import os
import pandas as pd

app = Flask(__name__)

# Create directories if not exist
if not os.path.exists('generated'):
    os.makedirs('generated')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' in request.files:  # If file is uploaded
            file = request.files['file']
            if file.filename.endswith('.xlsx'):
                # Read the Excel file using pandas
                df = pd.read_excel(file)
                
                # Strip any extra spaces from column names
                df.columns = df.columns.str.strip()

                # Print column names for debugging
                print(df.columns)  # Check column names in the uploaded file

                # Loop through each row in the Excel sheet
                for index, row in df.iterrows():
                    name = row['Name']  # Ensure column names match the actual ones
                    department = row['Department']
                    event = row['Event']

                    # Generate certificate for each participant
                    image = Image.open("static/certificate_template.jpg")
                    draw = ImageDraw.Draw(image)

                    font = ImageFont.truetype("arial.ttf", 50)
                    draw.text((480, 1080), name, fill="black", font=font)
                    draw.text((520, 1150), department, fill="black", font=font)
                    draw.text((810, 1200), event, fill="black", font=font)

                    # Save the certificate
                    output_path = f"generated/{name}_certificate.jpg"
                    image.save(output_path)

                # Optionally, return the last generated certificate or a success message
                return send_file(f"generated/{name}_certificate.jpg", as_attachment=True)

        elif 'name' in request.form:  # If form data is submitted
            name = request.form['name']
            department = request.form['department']
            event = request.form['event']

            # Generate a single certificate
            image = Image.open("static/certificate_template.jpg")
            draw = ImageDraw.Draw(image)

            font = ImageFont.truetype("arial.ttf", 50)
            draw.text((480, 1080), name, fill="black", font=font)
            draw.text((560, 1162), department, fill="black", font=font)
            draw.text((947, 1210), event, fill="black", font=font)

            output_path = f"generated/{name}_certificate.jpg"
            image.save(output_path)

            return send_file(output_path, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
