from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from pathlib import Path
import tempfile
import zipfile
import uuid
import json
import re

app = FastAPI()

API_DIR = Path(__file__).resolve().parent
BASE_DIR = API_DIR.parent

TEMPLATE_DIR = BASE_DIR / "templates"
MAPPING_DIR = BASE_DIR / "mappings"


def safe_name(value):
    if not value:
        return "DOSSIER"
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in str(value))


def deep_get(data, path, default=""):
    current = data
    for part in str(path).split("."):
        if isinstance(current, dict):
            current = current.get(part, default)
        else:
            return default
    return current if current is not None else default


def to_number(value):
    if value is None or value == "":
        return 0

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip()
    text = text.replace(" ", "").replace("\u00a0", "")
    text = text.replace(",", ".")

    if text.endswith("%"):
        try:
            return float(text.replace("%", "")) / 100
        except Exception:
            return 0

    text = re.sub(r"[^0-9.\-]", "", text)

    try:
        return float(text)
    except Exception:
        return 0


def normalize_value(value, value_type="text"):
    if value_type in ["number", "integer", "float"]:
        return to_number(value)

    if value_type == "boolean":
        return bool(value)

    return value if value is not None else ""


def normalize_investissements(data):
    """
    Accepte deux formats :
    1) investissements = { "terrain": [...], "constructions": [...] }
    2) investissements = [ {"categorie": "terrain", ...}, ... ]
    """
    investissements = data.get("investissements", {})

    if isinstance(investissements, dict):
        return investissements

    result = {
        "terrain": [],
        "constructions": [],
        "amenagement_agencement": [],
        "materiel_equipement": [],
        "frais_preliminaires": [],
        "divers_imprevus": []
    }

    if isinstance(investissements, list):
        for item in investissements:
            if not isinstance(item, dict):
                continue

            cat = (
                item.get("categorie")
                or item.get("category")
                or item.get("type")
                or ""
            ).lower()

            if "terrain" in cat:
                result["terrain"].append(item)
            elif "construction" in cat:
                result["constructions"].append(item)
            elif "amenagement" in cat or "aménagement" in cat or "agencement" in cat:
                result["amenagement_agencement"].append(item)
            elif "materiel" in cat or "matériel" in cat or "equipement" in cat or "équipement" in cat:
                result["materiel_equipement"].append(item)
            elif "frais" in cat or "preliminaire" in cat or "préliminaire" in cat:
                result["frais_preliminaires"].append(item)
            elif "divers" in cat or "imprevu" in cat or "imprévu" in cat:
                result["divers_imprevus"].append(item)
            else:
                result["materiel_equipement"].append(item)

    return result


def write_cell(ws, cell_ref, value, value_type="text"):
    cell = ws[cell_ref]

    # Ne jamais écraser les formules
    if isinstance(cell.value, str) and cell.value.startswith("="):
        return

    cell.value = normalize_value(value, value_type)


def apply_simple_mappings(wb, context, mappings):
    for item in mappings:
        sheet_name = item.get("sheet")
        cell_ref = item.get("cell")
        field = item.get("field")
        value_type = item.get("type", "text")

        if not sheet_name or not cell_ref or not field:
            continue

        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        value = deep_get(context, field, "")
        write_cell(ws, cell_ref, value, value_type)


def apply_table_mappings(wb, context, table_mappings):
    investissements = normalize_investissements(context)

    for table in table_mappings:
        sheet_name = table.get("sheet")
        source_array = table.get("source_array", "")
        start_row = int(table.get("start_row", 0))
        end_row = int(table.get("end_row", 0))
        columns = table.get("columns", [])

        if sheet_name not in wb.sheetnames:
            continue

        if not start_row or not end_row:
            continue

        ws = wb[sheet_name]

        # source_array exemple : investissements.terrain
        source_key = source_array.split(".")[-1]
        rows = investissements.get(source_key, [])

        if not isinstance(rows, list):
            continue

        max_rows = end_row - start_row + 1

        for i, row_data in enumerate(rows[:max_rows]):
            excel_row = start_row + i

            for col in columns:
                column_letter = col.get("column")
                field = col.get("field")
                value_type = col.get("type", "text")
                default_value = col.get("default", "")

                if not column_letter or not field:
                    continue

                cell_ref = f"{column_letter}{excel_row}"

                value = row_data.get(field, default_value)

                # TVA par défaut par catégorie
                if field == "taux_tva" and (value is None or value == ""):
                    default_by_category = col.get("default_by_category", {})
                    category = table.get("category", "default")
                    value = default_by_category.get(category, default_by_category.get("default", 0))

                write_cell(ws, cell_ref, value, value_type)


def render_docx(template_name, output_path, context):
    template_path = TEMPLATE_DIR / template_name

    if not template_path.exists():
        raise FileNotFoundError(f"Template Word introuvable : {template_name}")

    doc = DocxTemplate(str(template_path))
    doc.render(context)
    doc.save(str(output_path))


def render_bp_excel(output_path, context):
    template_path = TEMPLATE_DIR / "BP_template.xlsx"
    mapping_path = MAPPING_DIR / "mapping_bp_istitmar.json"

    if not template_path.exists():
        raise FileNotFoundError("Template Excel introuvable : templates/BP_template.xlsx")

    if not mapping_path.exists():
        raise FileNotFoundError("Mapping BP introuvable : mappings/mapping_bp_istitmar.json")

    with open(mapping_path, "r", encoding="utf-8") as f:
        mapping = json.load(f)

    wb = load_workbook(template_path)

    # Mapping structuré
    apply_simple_mappings(wb, context, mapping.get("scalar_mappings", []))
    apply_simple_mappings(wb, context, mapping.get("financement_pi_mappings", []))
    apply_table_mappings(wb, context, mapping.get("table_mappings", []))
    apply_simple_mappings(wb, context, mapping.get("cpc_mappings", []))
    apply_simple_mappings(wb, context, mapping.get("bilan_mappings", []))
    apply_simple_mappings(wb, context, mapping.get("impacts_mappings", []))

    # Forcer Excel à recalculer à l'ouverture
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

    wb.save(output_path)


@app.get("/")
def root():
    return {
        "status": "ok",
        "message": "GO SIYAHA filler API",
        "templates_dir": str(TEMPLATE_DIR),
        "mappings_dir": str(MAPPING_DIR)
    }


@app.get("/health")
def health():
    return {
        "status": "ok",
        "templates": {
            "declaration_factures.docx": (TEMPLATE_DIR / "declaration_factures.docx").exists(),
            "DAP_template.docx": (TEMPLATE_DIR / "DAP_template.docx").exists(),
            "BP_template.xlsx": (TEMPLATE_DIR / "BP_template.xlsx").exists(),
            "demande_honneur.docx": (TEMPLATE_DIR / "demande_honneur.docx").exists(),
            "engagement_capacite.docx": (TEMPLATE_DIR / "engagement_capacite.docx").exists()
        },
        "mappings": {
            "mapping_bp_istitmar.json": (MAPPING_DIR / "mapping_bp_istitmar.json").exists()
        }
    }


@app.post("/fill")
async def fill(request: Request):
    try:
        payload = await request.json()
        data = payload.get("data", payload)

        selected_template = (
            payload.get("selected_template")
            or data.get("selected_template")
            or "dossier_complet"
        )

        context = {
            **data,
            "selected_template": selected_template,
            "dossier": data.get("dossier", {}),
            "entreprise": data.get("entreprise", {}),
            "dirigeant": data.get("dirigeant", {}),
            "projet": data.get("projet", {}),
            "investissements": data.get("investissements", {}),
            "emplois": data.get("emplois", {}),
            "banque": data.get("banque", {}),
            "financement_pi": data.get("financement_pi", {}),
            "cpc_historique": data.get("cpc_historique", {}),
            "cpc_previsionnel": data.get("cpc_previsionnel", {}),
            "bilan_historique": data.get("bilan_historique", {}),
            "bilan_previsionnel": data.get("bilan_previsionnel", {}),
            "impacts_historique": data.get("impacts_historique", {}),
            "impacts_previsionnels": data.get("impacts_previsionnels", {}),
            "hypotheses_financieres": data.get("hypotheses_financieres", {}),
            "financement_expert": data.get("financement_expert", {})
        }

        dossier = context["dossier"]
        entreprise = context["entreprise"]

        identifiant = safe_name(dossier.get("identifiant", "DOSSIER"))
        societe = safe_name(entreprise.get("raison_sociale", "SOCIETE"))

        job_id = str(uuid.uuid4())
        tmp_dir = Path(tempfile.gettempdir()) / f"go_siyaha_{job_id}"
        tmp_dir.mkdir(parents=True, exist_ok=True)

        output_zip = tmp_dir / f"{identifiant}_{societe}_GO_SIYAHA.zip"
        generated_files = []

        def add_docx(template_name, output_name):
            output_path = tmp_dir / output_name
            render_docx(template_name, output_path, context)
            generated_files.append(output_path)

        def add_bp():
            output_path = tmp_dir / f"{identifiant}_{societe}_BP.xlsx"
            render_bp_excel(output_path, context)
            generated_files.append(output_path)

        if selected_template == "bp_excel":
            add_bp()

        elif selected_template == "dap_word":
            add_docx("DAP_template.docx", f"{identifiant}_{societe}_DAP.docx")

        elif selected_template == "demande_honneur":
            if (TEMPLATE_DIR / "demande_honneur.docx").exists():
                add_docx("demande_honneur.docx", f"{identifiant}_{societe}_demande_honneur.docx")
            else:
                add_docx("declaration_factures.docx", f"{identifiant}_{societe}_declaration_factures.docx")

        elif selected_template == "engagement_capacite":
            add_docx("engagement_capacite.docx", f"{identifiant}_{societe}_engagement_capacite.docx")

        elif selected_template == "dossier_complet":
            add_docx("declaration_factures.docx", f"{identifiant}_{societe}_declaration_factures.docx")

            if (TEMPLATE_DIR / "DAP_template.docx").exists():
                add_docx("DAP_template.docx", f"{identifiant}_{societe}_DAP.docx")

            if (TEMPLATE_DIR / "BP_template.xlsx").exists():
                add_bp()

            if (TEMPLATE_DIR / "demande_honneur.docx").exists():
                add_docx("demande_honneur.docx", f"{identifiant}_{societe}_demande_honneur.docx")

            if (TEMPLATE_DIR / "engagement_capacite.docx").exists():
                add_docx("engagement_capacite.docx", f"{identifiant}_{societe}_engagement_capacite.docx")

        else:
            return JSONResponse(
                status_code=400,
                content={
                    "error": f"selected_template inconnu : {selected_template}",
                    "allowed": [
                        "bp_excel",
                        "dap_word",
                        "demande_honneur",
                        "engagement_capacite",
                        "dossier_complet"
                    ]
                }
            )

        if not generated_files:
            return JSONResponse(
                status_code=400,
                content={"error": "Aucun fichier généré. Vérifiez les templates."}
            )

        with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
            for file_path in generated_files:
                zipf.write(file_path, arcname=file_path.name)

        return FileResponse(
            path=str(output_zip),
            media_type="application/zip",
            filename=output_zip.name
        )

    except FileNotFoundError as e:
        return JSONResponse(status_code=404, content={"error": str(e)})

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
