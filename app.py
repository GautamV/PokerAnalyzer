import os 

from flask import Flask, render_template, send_from_directory, request, redirect
from werkzeug.utils import secure_filename
from analyzer import Analyzer

app = Flask(__name__)

UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/uploads/'
DOWNLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/downloads/'

if not os.path.exists(UPLOAD_FOLDER):
	os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(DOWNLOAD_FOLDER):
	os.makedirs(DOWNLOAD_FOLDER)

ALLOWED_EXTENSIONS = {'csv'}
def allowed_file(filename):
	return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def process_file(input_filename, output_filename, player_name): 
	Analyzer(input_filename, player_name).save_sheet(output_filename)

@app.route('/', methods=['GET', 'POST'])
def main():
	if request.method == 'POST': 
		if 'file' not in request.files: 
			return redirect(request.url)
		file = request.files['file']
		if file is None or not allowed_file(file.filename): 
			return redirect(request.url)
		filename = secure_filename(file.filename)
		file.save(os.path.join(UPLOAD_FOLDER, filename))
		output_filename = '{}-({}).xlsx'.format(filename.split('.')[0], request.form['name'])
		process_file(os.path.join(UPLOAD_FOLDER, filename), os.path.join(DOWNLOAD_FOLDER, output_filename), request.form['name'])
		return send_from_directory(DOWNLOAD_FOLDER, output_filename, as_attachment=True)
	elif request.method == 'GET':
		return render_template('app.html')