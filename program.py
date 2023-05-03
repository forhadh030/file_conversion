from flask import Flask, render_template, request, make_response
import openpyxl
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/', methods=['POST'])
def convert_file():
    file = request.files['file']
    if not file:
        return 'No file uploaded.'

    wb = openpyxl.load_workbook(filename=io.BytesIO(file.read()))
    ws = wb.active

    text = ''
    for row in ws.iter_rows(values_only=True):
        text += '\t'.join(str(cell) for cell in row) + '\n'

    # Get the original filename and extension
    filename = file.filename
    extension = filename.split('.')[-1]

    # Create a response with the converted text
    response = make_response(text)

    # Set the filename and extension for the response
    response.headers['Content-Disposition'] = f'attachment; filename={filename.replace(extension, "txt")}'
    response.mimetype = 'text/plain'
    return response

if __name__ == '__main__':
    app.run(debug=True)