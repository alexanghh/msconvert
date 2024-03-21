import aiofiles as aiof
import comtypes.client
import io
import os
import uuid
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse

app = FastAPI(
    title="MSConvert",
    description="Converts Microsoft Office file to PDF file using **Microsoft Office software**. Supported file formats include **doc, docx, ppt, pptx, xls, xlsx**.",
    summary="Microsoft Office to PDF File Converter.",
    version="0.0.1",
)

@app.get("/")
async def hello_world():
    return {"Hello": "World"}

@app.post("/convert/",
          tags=["Conversion"],
          response_class=StreamingResponse,
          responses={
              200: {
                  "content": {"application/pdf": {
                      "example": "(no example available)"
                  }},
                  "description": "Return the PDF file.",
              }
          },)
async def convert(office_file: UploadFile = File(...)):
    """
    Converts Microsoft Office file to PDF file
    """
    filename = f'{uuid.uuid4()}{office_file.filename}'
    async with aiof.open(filename, "wb") as out:
        await out.write(await office_file.read())
        await out.flush()
    _, extension = os.path.splitext(office_file.filename)
    try:
        data = None
        match extension:
            case '.doc' | '.docx':
                data = await convert_doc_to_pdf(filename)
            case '.ppt' | '.pptx':
                data = await convert_ppt_to_pdf(filename)
            case '.xls' | '.xlsx':
                data = await convert_xls_to_pdf(filename)
            case _:
                raise HTTPException(status_code=400, detail="File format not supported")
        if data is not None:
            return StreamingResponse(data,
                                     media_type='application/pdf',
                                     headers={
                                         'Content-Disposition': 'attachment; filename="{}"'.format(
                                             office_file.filename + ".pdf")})
    finally:
        os.remove(filename)
    raise HTTPException(status_code=400, detail="File not converted")


async def convert_xls_to_pdf(filename):
    try:
        comtypes.CoInitialize()
        excel = comtypes.client.CreateObject('Excel.Application')
        excel.Visible = False
        sheet = excel.Workbooks.Open(os.path.abspath(filename))
        sheet.ExportAsFixedFormat(0, os.path.abspath(filename + ".pdf"), 1, 0)
        sheet.Close()
        excel.Quit()
    finally:
        comtypes.CoUninitialize()
    return await cache_delete_file(filename + ".pdf")


async def convert_ppt_to_pdf(filename, format_type=32):
    try:
        comtypes.CoInitialize()
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = True
        deck = powerpoint.Presentations.Open(os.path.abspath(filename))
        deck.SaveAs(os.path.abspath(filename + ".pdf"), format_type)  # formatType = 32 for ppt to pdf
        deck.Close()
        powerpoint.Quit()
    finally:
        comtypes.CoUninitialize()
    return await cache_delete_file(filename + ".pdf")


async def convert_doc_to_pdf(filename, format_type=17):
    try:
        comtypes.CoInitialize()
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(filename))
        doc.SaveAs(os.path.abspath(filename + ".pdf"), format_type)
        doc.Close()
        word.Quit()
    finally:
        comtypes.CoUninitialize()
    return await cache_delete_file(filename + ".pdf")


async def cache_delete_file(filename):
    cached_file = io.BytesIO()
    async with aiof.open(os.path.abspath(filename), 'rb') as fo:
        cached_file.write(await fo.read())
    cached_file.seek(0)
    os.remove(os.path.abspath(filename))
    return cached_file
