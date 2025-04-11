from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from openpyxl import load_workbook
import datetime

app = FastAPI()

# Biarkan React bisa akses API ini
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Ganti ini kalau kamu host React di domain tertentu
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Model data dari form
class ContactForm(BaseModel):
    name: str
    email: str
    message: str

@app.post("/submit")
async def submit_form(data: ContactForm):
    try:
        wb = load_workbook("contacts.xlsx")
        ws = wb.active

        # Tambahkan data ke baris berikutnya
        ws.append([data.name, data.email, data.message])
        wb.save("contacts.xlsx")

        return {"message": "Pesan berhasil dikirim ✅"}
    except Exception as e:
        return {"message": f"Gagal menyimpan pesan: {e}"}
#uvicorn main:app --reload
