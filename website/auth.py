from flask import *
import img2pdf
import ppt2pdf
from pdf2docx import parse
import os
from typing import Tuple
import comtypes.client
import tempfile
import comtypes
from .models import User
from werkzeug.security import generate_password_hash, check_password_hash
from . import db   ##means from __init__.py import db
from flask_login import login_user, login_required, logout_user, current_user

UPLOADER_FOLDER = ''

auth = Blueprint('auth', __name__)

app = Flask(__name__)

app.config['UPLOADER_FOLDER']=UPLOADER_FOLDER
@auth.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')

        user = User.query.filter_by(email=email).first()
        if user:
            if check_password_hash(user.password, password):
                flash('Logged in successfully!', category='success')
                login_user(user, remember=True)
                return redirect(url_for('views.home'))
            else:
                flash('Incorrect password, try again.', category='error')
        else:
            flash('Username does not exist.', category='error')

    return render_template("login.html", user=current_user)


@auth.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('auth.login'))


@auth.route('/sign-up', methods=['GET', 'POST'])
def sign_up():
    if request.method == 'POST':
        email = request.form.get('email')
        first_name = request.form.get('fullName')
        password1 = request.form.get('password1')
        password2 = request.form.get('password2')

        user = User.query.filter_by(email=email).first()
        if user:
            flash('Email already exists.', category='error')
        elif len(email) < 4:
            flash('Email must be greater than 3 characters.', category='error')
        elif len(first_name) < 2:
            flash('First name must be greater than 1 character.', category='error')
        elif password1 != password2:
            flash('Passwords don\'t match.', category='error')
        elif len(password1) < 7:
            flash('Password must be at least 7 characters.', category='error')
        else:
            new_user = User(email=email, first_name=first_name, password=generate_password_hash(
                password1, method='sha256'))
            db.session.add(new_user)
            db.session.commit()
            login_user(new_user, remember=True)
            flash('Account created!', category='success')
            return redirect(url_for('views.home'))

    return render_template("sign_up.html", user=current_user)

@auth.route('/privacy', methods=['GET', 'POST'])
@login_required
def privacy():
    return render_template("privacy_policy.html", user=current_user)

@auth.route('/about-us', methods=['GET', 'POST'])
@login_required
def aboutus():
    return render_template("aboutus.html", user=current_user)

@auth.route('/contact-us', methods=['GET', 'POST'])
@login_required
def contact():
    return render_template("contactus.html", user=current_user)

@auth.route('/file-convertor', methods=['GET', 'POST'])
@login_required
def file_convertor():
    return render_template("file-convert.html", user=current_user)

# Image Covertor
@auth.route('/img-convertor', methods=['GET', 'POST'])
@login_required
def image_convertor():
    return render_template("img-convert.html", user=current_user)

@auth.route('/img-converted',methods = ['GET', 'POST'])
@login_required
def convert():
    global f1
    fi = request.files['img']
    f1 = fi.filename
    fi.save(os.path.join(app.config['UPLOADER_FOLDER'],fi.filename))
    i2pconverter(f1)
    return render_template('img-converted.html', user=current_user)

@auth.route('/img-download')
def download():
    filename = f1.split('.')[0]+'converted.pdf'
    #image = send_file(filename, as_attachment=False)
    #image.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
    return redirect(url_for('auth.file_convertor'))

def i2pconverter(file):
    pdfname = file.split('.')[0]+'converted'+'.pdf'
    with open(pdfname,'wb') as f:
        f.write(img2pdf.convert(file))

# Word Convertor
@auth.route('/word-convertor', methods=['GET', 'POST'])
@login_required
def word_convertor():
    return render_template("word-convert.html", user=current_user)

@auth.route('/word-converted', methods=['GET', 'POST'])
@login_required
def word_converted():
    if request.method=="POST":
        def convert_pdf2docx(input_file:str,output_file:str,pages:Tuple=None):
           if pages:
               pages = [int(i) for i in list(pages) if i.isnumeric()]

           result = parse(pdf_file=input_file,docx_with_path=output_file, pages=pages)
           summary = {
               "File": input_file, "Pages": str(pages), "Output File": output_file
            }

           print("\n".join("{}:{}".format(i, j) for i, j in summary.items()))
           return result
        file=request.files['filename']
        if file.filename!='':
           file.save(os.path.join(app.config['UPLOADER_FOLDER'],file.filename))
           input_file=file.filename
           output_file=r"hello.docx"
           convert_pdf2docx(input_file,output_file)
           doc=input_file.split(".")[0]+".docx"
           print(doc)
           lis=doc.replace(" ","=")
    return render_template("word-converted.html", user=current_user)

@auth.route('/word-download', methods=['GET', 'POST'])
def worddownload():
    if request.method == "POST":
        lis = request.form.get('filename', None)
        lis = lis.replace("=", " ")
        return send_file(lis, as_attachment=True)
    return redirect(url_for('auth.file_convertor'))

#PPT TO PDF Convertor
@auth.route('/ppt-convertor', methods=['GET', 'POST'])
@login_required
def ppt_convertor():
    return render_template("ppt-convert.html", user=current_user)

@auth.route('/ppt-converted', methods=['GET', 'POST'])
@login_required
def ppt_converted():
    ppt_file = request.files.get('ppt_file')
    if not ppt_file:
        return {'error': 'No file was provided.'}, 400

    # Save the uploaded file to a temporary file on the server.
    temp_file_path = os.path.join(tempfile.gettempdir(), ppt_file.filename)
    ppt_file.save(temp_file_path)

    # Create a temporary PDF file to write the converted PDF output to.
    temp_pdf_file = 'temp.pdf'

    # Initialize the COM library.
    comtypes.CoInitialize()

    try:
        # Use comtypes to convert the PowerPoint file to PDF.
        powerpoint = comtypes.client.CreateObject('Powerpoint.Application')
        powerpoint.Visible = 1
        ppt = powerpoint.Presentations.Open(temp_file_path)
        ppt.SaveAs(temp_pdf_file, 32)  # 32 is the value of the constant ppSaveAsPDF
        ppt.Close()
        powerpoint.Quit()
    finally:
        # Uninitialize the COM library.
        comtypes.CoUninitialize()

    # Send the PDF file as a response to the client.
    return send_file(temp_pdf_file, as_attachment=True, attachment_filename='converted.pdf')
    return redirect(url_for('auth.file_convertor'))