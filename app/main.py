from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import StreamingResponse
from app.services.ppt_generator import generate_ppt
import re
from datetime import datetime

app = FastAPI()


def sanitize_filename(value: str) -> str:
    if not value:
        return "document"
    # elimina caracteres inválidos en nombres de archivo
    value = re.sub(r'[\\/*?:"<>|]', "-", value)
    return value.strip()


@app.post("/generate-document")
async def generate_document(
    template_url: str = Form(...),
    document_type: str = Form(...),
    client_name: str = Form(...),
    current_date: str = Form(...),
    due_date: str = Form(...),
    addressed_to: str = Form(...),
    service_type: str = Form(...),
    users: str = Form(...),
    price: str = Form(...),
    period: str = Form(...),
    logo_url: str = Form(None)  # ← ahora soporta logo opcional
):
    try:

        replacements = {
            "{{client_name}}": client_name,
            "{{current_date}}": current_date,
            "{{due_date}}": due_date,
            "{{addressed_to}}": addressed_to,
            "{{service_type}}": service_type,
            "{{users}}": users,
            "{{price}}": price,
            "{{period}}": period,
        }

        output = generate_ppt(
            template_url=template_url,
            replacements=replacements,
            logo_url=logo_url
        )

        safe_client = sanitize_filename(client_name)
        safe_date = sanitize_filename(current_date)
        safe_doc_type = sanitize_filename(document_type)

        file_name = f"{safe_doc_type}-{safe_client}-{safe_date}.pptx"

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f"attachment; filename={file_name}"
            }
        )

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating document: {str(e)}"
        )