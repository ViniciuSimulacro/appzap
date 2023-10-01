import os
from flask import Flask, flash, request, redirect, url_for, render_template
from werkzeug.utils import secure_filename
import openpyxl

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xls','xlsx','png', 'jpg', 'jpeg', 'gif'}

mensagem_armazenada = ''

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file_contact():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return redirect(url_for('upload_file_img', filename=filename))
    return render_template('download.html')

@app.route('/download_img', methods=['GET', 'POST'])
def upload_file_img():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return redirect(url_for('results', name=filename))
    return render_template('download_img.html')

@app.route('/message', methods=['GET', 'POST'])
def write_msg():
    if request.method == 'POST':
        global mensagem_armazenada
        mensagem = request.form['message']
        mensagem_armazenada = mensagem
        return redirect(url_for('results'))
    return render_template('message.html')

@app.route('/results', methods=['GET','POST'])
def results():
    image = request.args.get('image')
    book = openpyxl.load_workbook('uploads/contatos.xlsx')
    contacts = book['Contatos']
    send_list = []
    for rows in contacts.iter_rows(min_row=2):
        send_list.append([rows[0].value,rows[1].value])
    return render_template('results.html', send_list=send_list, image=image)

@app.route('/send_msg', methods=['GET', 'POST'])
def send_msg():
    return 'Vamos enviar as msgs'

if __name__ == '__main__':
    app.run(debug=True)
