from flask import Flask, request, send_file
from werkzeug.utils import secure_filename
import os, io, uuid
import comtypes.client

app = Flask(__name__)

@app.route("/")
def hello_world():
  return "<p>Hello, World!</p>"

@app.route("/convert", methods = ["POST"])
def convert():
  f = request.files['file']
  filename = str(uuid.uuid4()) + secure_filename(f.filename)
  f.save(filename)
  _, extension = os.path.splitext(f.filename)
  data = None
  match extension:
    case '.doc' | '.docx':
      data = convert_doc_to_pdf(filename)
      os.remove(filename)
    case '.ppt' | '.pptx':
      data = convert_ppt_to_pdf(filename)
      os.remove(filename)
    case '.xls' | '.xlsx':
      data = convert_xls_to_pdf(filename)
      os.remove(filename)
    case _:
      pass
  if data is not None:
    return send_file(data, mimetype='application/pdf', download_name=f.filename+'.pdf')
  return ''

def convert_xls_to_pdf(filename):
  comtypes.CoInitialize()
  excel = comtypes.client.CreateObject('Excel.Application')
  excel.Visible = False
  sheet = excel.Workbooks.Open(os.path.abspath(filename))
  sheet.ExportAsFixedFormat(0, os.path.abspath(filename+".pdf"), 1, 0)
  sheet.Close()
  excel.Quit()
  comtypes.CoUninitialize()
  return cache_delete_file(filename+".pdf")

def convert_ppt_to_pdf(filename, formatType = 32):
  comtypes.CoInitialize()
  powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
  powerpoint.Visible = True
  deck = powerpoint.Presentations.Open(os.path.abspath(filename))
  deck.SaveAs(os.path.abspath(filename+".pdf"), formatType) # formatType = 32 for ppt to pdf
  deck.Close()
  powerpoint.Quit()
  comtypes.CoUninitialize()
  return cache_delete_file(filename+".pdf")

def convert_doc_to_pdf(filename, formatType = 17):
  comtypes.CoInitialize()
  word = comtypes.client.CreateObject('Word.Application')
  word.Visible = False
  doc = word.Documents.Open(os.path.abspath(filename))
  doc.SaveAs(os.path.abspath(filename+".pdf"), formatType)
  doc.Close()
  word.Quit()
  comtypes.CoUninitialize()
  return cache_delete_file(filename+".pdf")

def cache_delete_file(filename):
  cached_file = io.BytesIO()
  with open(os.path.abspath(filename), 'rb') as fo:
    cached_file.write(fo.read())
  cached_file.seek(0)
  os.remove(os.path.abspath(filename))
  return cached_file