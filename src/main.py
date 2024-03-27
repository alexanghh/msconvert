import aiofiles as aiof
import comtypes.client
from dotenv import load_dotenv
import io
import os
import uuid
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.openapi.docs import (
    get_redoc_html,
    get_swagger_ui_html,
    get_swagger_ui_oauth2_redirect_html,
)
from fastapi.staticfiles import StaticFiles
from fastapi.responses import StreamingResponse
import uvicorn

app = FastAPI(
    title="MSConvert",
    description="Converts Microsoft Office file to PDF file using **Microsoft Office software**. Supported file formats include **doc, docx, ppt, pptx, xls, xlsx**.",
    summary="Microsoft Office to PDF File Converter.",
    version="0.0.1",
    docs_url=None, # disable to allow overriding of cdn
    redoc_url=None, # disable
)

app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", include_in_schema=False)
@app.get("/docs", include_in_schema=False)
async def custom_swagger_ui_html():
    return get_swagger_ui_html(
        openapi_url=app.openapi_url,
        title=app.title + " - Swagger UI",
        oauth2_redirect_url=app.swagger_ui_oauth2_redirect_url,
        swagger_js_url="/static/swagger-ui-bundle.js",
        swagger_css_url="/static/swagger-ui.css",
    )


@app.get(app.swagger_ui_oauth2_redirect_url, include_in_schema=False)
async def swagger_ui_redirect():
    return get_swagger_ui_oauth2_redirect_html()


@app.get("/redoc", include_in_schema=False)
async def redoc_html():
    return get_redoc_html(
        openapi_url=app.openapi_url,
        title=app.title + " - ReDoc",
        redoc_js_url="/static/redoc.standalone.js",
    )


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


@app.get("/health", tags=["Health"])
@app.get("/healthz", tags=["Health"])
def get_health():
    return "OK"


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=5000, log_level="info")
