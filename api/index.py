from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse
from docxtpl import DocxTemplate
from pathlib import Path
import tempfile
import zipfile
import uuid

app = FastAPI()

API_DIR = Path(__file__).resolve().parent
BASE_DIR = API_DIR.parent


def safe_name(value):
    if not value:
        return "DOSSIER"
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in str(value))


@app.get("/")
def root():
    return {"status": "ok", "message": "GO SIYAHA filler API"}


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/fill")
async def fill(request: Request):
    try:
        payload = await request.json()
        data = payload.get("data", payload)

        dossier = data.get("dossier", {})
        entreprise = data.get("entreprise", {})
        dirigeant = data.get("dirigeant", {})
        projet = data.get("projet", {})

        possible_paths = [
            API_DIR / "declaration_factures.docx",
            BASE_DIR / "declaration_factures.docx",
            BASE_DIR / "templates" / "declaration_factures.docx",
            API_DIR / "templates" / "declaration_factures.docx",
        ]

        template_path = None
        for path in possible_paths:
            if path.exists():
                template_path = path
                break

        if template_path is None:
            return JSONResponse(
                status_code=404,
                content={
                    "error": "declaration_factures.docx introuvable",
                    "checked_paths": [str(p) for p in possible_paths]
                }
            )

        context = {
            "dossier": dossier,
            "entreprise": entreprise,
            "dirigeant": dirigeant,
            "projet": projet
        }

        identifiant = safe_name(dossier.get("identifiant", "DOSSIER"))
        societe = safe_name(entreprise.get("raison_sociale", "SOCIETE"))

        job_id = str(uuid.uuid4())
        tmp_dir = Path(tempfile.gettempdir()) / f"go_siyaha_{job_id}"
        tmp_dir.mkdir(parents=True, exist_ok=True)

        output_docx = tmp_dir / f"{identifiant}_{societe}_declaration_factures.docx"
        output_zip = tmp_dir / f"{identifiant}_{societe}_GO_SIYAHA.zip"

        doc = DocxTemplate(str(template_path))
        doc.render(context)
        doc.save(str(output_docx))

        with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(output_docx, arcname=output_docx.name)

        return FileResponse(
            path=str(output_zip),
            media_type="application/zip",
            filename=output_zip.name
        )

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
