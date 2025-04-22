from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from openpyxl import load_workbook
from fastapi.responses import FileResponse

app = FastAPI()

app.add_middleware(
    app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://ferdinandport.vercel.app/"],  # sesuaikan
    allow_credentials=True,
    allow_methods=["POST, GET, PUT, DELETE, OPTIONS"],
    allow_headers=["Content-Type", "Authorization"],
)
)

class ContactForm(BaseModel):
    name: str
    email: str
    message: str

@app.post("/submit")
async def submit_form(data: ContactForm):
    try:
        wb = load_workbook("contacts.xlsx")
        ws = wb.active
        ws.append([data.name, data.email, data.message])
        wb.save("contacts.xlsx")
        return {"message": "Pesan berhasil dikirim ✅"}
    except Exception as e:
        return {"message": f"Gagal menyimpan pesan: {e}"}

@app.get("/download")
def download_excel():
    return FileResponse("contacts.xlsx", media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename="contacts.xlsx")
