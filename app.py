from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from tools import upload_file_to_s3,read_object
from docsplit import lambda_split_word

app = Flask(__name__)

@app.route('/upload')
def upload_file():
   return render_template('upload.html')
	
      
@app.route("/documents/merge", methods=["POST"])
def demo():
    if "file" not in request.files:
        return "No user_file key in request.files"

    file = request.files["file"]
    #flask.request.files.getlist("file[]") Multiple files
    #for file in files:
    #file = request.files["parameter"] add split parameter

    if file.filename == "":
        return "Please select a file"

    if file:
        file.filename = secure_filename(file.filename)
        output = upload_file_to_s3(file, "appbucket0912")
        file_object=read_object(file, "appbucket0912")
        lambda_split_word(file_object)
        return str(output)

    
   
@app.route('/documents/split', methods = ['GET', 'POST'])
def split_file():
   if "file" not in request.files:
        return "No user_file key in request.files"

   file = request.files["file"]
    #flask.request.files.getlist("file[]") Multiple files
    #file = request.files["parameter"] add split parameter

   if file.filename == "":
        return "Please select a file"

   if file:
        file.filename = secure_filename(file.filename)
        output = upload_file_to_s3(file, "appbucket0912")
        file_object=read_object(file, "appbucket0912")
        return str(output)
		
if __name__ == '__main__':
   app.run(debug = True)