from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import Response, HTMLResponse
from fastapi.staticfiles import StaticFiles
import tempfile, os

from parser import parse_relatorio
from gerador import gerar_dashboard_pdf

app = FastAPI(title="Concrelongo Dashboard")

# ── Página principal ──
@app.get("/", response_class=HTMLResponse)
async def index():
    with open("templates/index.html", encoding="utf-8") as f:
        return f.read()

# ── Endpoint de geração ──
@app.post("/gerar")
async def gerar(pdf: UploadFile = File(...)):
    if not pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Envie um arquivo .pdf")

    # Salva temporariamente
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(await pdf.read())
        tmp_path = tmp.name

    try:
        dados   = parse_relatorio(tmp_path)
        if not dados["remessas"]:
            raise HTTPException(422, "Nenhuma remessa encontrada no PDF. Verifique o arquivo.")
        pdf_out = gerar_dashboard_pdf(dados)
    finally:
        os.unlink(tmp_path)

    periodo = f'{dados["periodo_inicio"].replace("/","")}_a_{dados["periodo_fim"].replace("/","")}'
    filename = f"dashboard_concrelongo_{periodo}.pdf"

    return Response(
        content=pdf_out,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )
