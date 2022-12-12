from flask import Flask, render_template, request,flash,redirect,send_from_directory,send_file
from werkzeug.utils import secure_filename
from tools import upload_file_to_s3,read_object
from docmerge import DocMerge
import splitdoc
import pathlib
import os

current_path=pathlib.Path(__file__).parent.resolve()

app = Flask(__name__)
app.secret_key = "1cd3a5de-0f6e-4110-9704-60c5d29797a1"
app.config["MERGE_OUTPUT"]=os.path.join(current_path,"tmp","merge","output")


@app.route('/split')
def split_ui():
   return render_template('upload.html')

@app.route('/merge')
def merge_ui():
   return render_template('uploadmerge.html')

      
@app.route("/documents/merge", methods=["POST"])
def merge_file():
    if request.method == 'POST':
        if 'files[]' not in request.files:
            flash('No file part')
            return redirect(request.url)
        files = request.files.getlist('files[]')
        mType = request.form.get('mergeType')
        output_files=[]
        for file in files:
            if file :
                #save_path=os.getcwd()+ "/tmp/input/"+secure_filename(file.filename)
                filename = secure_filename(file.filename)
                save_path=os.getcwd()+ "/tmp/merge/input/"+secure_filename(file.filename)
                output_files.append(save_path)
                file.save(save_path)
        object=DocMerge(output_files,file.filename[:-6])
        object.cleanup()
        if mType.lower()=="section":      
            filename=object.merge()
        else:
            filename=object.mergeByPage()

        return send_file(filename,
                mimetype = 'docx',
                as_attachment = True)
        #return send_from_directory(app.config["MERGE_OUTPUT"], filename=filename, as_attachment=True)

    
   
@app.route('/documents/split', methods = ['GET', 'POST'])
def split_file():
   if "file" not in request.files:
        return "No user_file key in request.files"

   file = request.files["file"]
   if file.filename == "":
        return "Please select a file"

   if file:
        file.filename = secure_filename(file.filename)
        save_path=os.getcwd()+ "/tmp/input/"+secure_filename(file.filename)
        file.save(save_path)
        split_object=splitdoc.SplitDocument(save_path,file.filename[:-4])
        split_object.cleanup()
        split_object.create_split_sequence()
        split_object.split_documents()
        return split_object.download_all(file.filename[:-4])
		
if __name__ == '__main__':
   app.run(debug = True)