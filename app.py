from flask import Flask, request, render_template, send_file
import pandas as pd
import io
from automation import process_excel

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    df = pd.read_excel(file)

    output = process_excel(df)
    return send_file(output, as_attachment=True, download_name="updated_file.xlsx")

if __name__ == '__main__':
    app.run()
