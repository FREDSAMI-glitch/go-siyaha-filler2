from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from pathlib import Path
from jinja2 import Environment, ChainableUndefined
from docx import Document
import tempfile
import zipfile
import uuid
import json
import re
import copy

app = FastAPI()

API_DIR = Path(__file__).resolve().parent
POSSIBLE_BASE_DIRS = [API_DIR, API_DIR.parent]

BASE_DIR = None
for d in POSSIBLE_BASE_DIRS:
    if (d / "templates").exists() or (d / "mappings").exists():
        BASE_DIR = d
        break

if BASE_DIR is None:
    BASE_DIR = API_DIR.parent

TEMPLATE_DIR = BASE_DIR / "templates"
MAPPING_DIR = BASE_DIR / "mappings"
DOCX_JINJA_ENV = Environment(undefined=ChainableUndefined)


def safe_name(value):
    if not value:
        return "DOSSIER"
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in str(value))


def as_dict(value):
    return value if isinstance(value, dict) else {}


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

    text = str(value).strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")

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

    if value_type == "money":
        n = to_number(value)
        return f"{n:,.0f}".replace(",", " ")

    return value if value is not None else ""


def add_docx_defaults(context):
    context.setdefault("dossier", {})
    context.setdefault("entreprise", {})
    context.setdefault("dirigeant", {})
    context.setdefault("projet", {})
    context.setdefault("banque", {})
    context.setdefault("emplois", {})
    context.setdefault("investissements", {})
    context.setdefault("financement_pi", {})
    context.setdefault("financement_expert", {})
    context.setdefault("hypotheses_financieres", {})
    context.setdefault("financement_checkbox", {})

    for key in [
        "dossier",
        "entreprise",
        "dirigeant",
        "projet",
        "banque",
        "emplois",
        "financement_pi",
        "financement_expert",
        "hypotheses_financieres",
        "financement_checkbox",
    ]:
        context[key] = as_dict(context.get(key))

    common_defaults = {
        "identifiant": "",
        "date_dossier": "",
        "lieu_signature": "",
        "date_signature": "",
        "numero_dossier": "",
        "programme": "GO SIYAHA",

        "raison_sociale": "",
        "denomination": "",
        "forme_juridique": "",
        "date_creation": "",
        "rc": "",
        "numero_rc": "",
        "ice": "",
        "cnss": "",
        "identifiant_fiscal": "",
        "patente": "",
        "adresse_siege": "",
        "adresse": "",
        "telephone": "",
        "tel": "",
        "email": "",
        "site_web": "",
        "capital_social": "",
        "capital_social_mad": 0,
        "actionnaires": "",
        "activite": "",
        "activite_selection": "",
        "activite_autre": "",
        "activite_detaillee": "",
        "objet_social": "",
        "attestation_classement": "",
        "classement": "",
        "categorie_classement": "",
        "type_categorie": "",
        "secteur": "",
        "secteur_activite": "",

        "nom": "",
        "prenom": "",
        "cin": "",
        "qualite": "",
        "fonction": "",
        "mobile": "",
        "gsm": "",
        "telephone_fixe": "",
        "profil": "",
        "affaires_gerees": "",

        "objet": "",
        "objectif": "",
        "description": "",
        "ville_region": "",
        "region": "",
        "province": "",
        "commune": "",
        "adresse_installations": "",
        "adresse_site": "",
        "lieu_realisation": "",
        "coordonnees_gps": "",
        "latitude": "",
        "longitude": "",
        "branche_activite": "",
        "filieres": "",
        "ecosystemes": "",
        "activites_envisagees": "",
        "offre_animation": "",
        "fiches_projet": "",
        "autorisations": "",
        "biens_services_produits": "",
        "investissement_total": 0,
        "investissement_total_mad": 0,
        "mode_financement_detail": "",
        "surface_terrain": "",
        "nature_terrain": "",
        "titre_foncier": "",
        "statut_foncier": "",
        "surface_mode_occupation": "",
        "planning_realisation": "",
        "date_demarrage_prevue": "",
        "annee_demarrage": "",
        "responsable_projet": "",
        "responsable_mobile": "",
        "role_region": "",
        "role_balance_commerciale": "",

        "banque_partenaire": "",
        "banque": "",
        "forme_juridique_banque": "",
        "capital": "",
        "capital_social_banque": "",
        "siege": "",
        "siege_social": "",
        "fonds_propres": 0,
        "credit_bancaire": 0,
        "cmt": 0,
        "leasing": 0,
        "credit_fournisseur": 0,
        "prime": 0,
        "prime_istitmar": 0,
        "montant_prime": 0,
        "mode_financement": "",

        "emplois_directs": "",
        "emplois_indirects": "",
        "emplois_stables": "",
        "effectif": "",
        "ca_prevu_annee_1": 0,
        "croissance_ca": 0,
    }

    for key, value in common_defaults.items():
        context.setdefault(key, value)

    for section in [
        "dossier",
        "entreprise",
        "dirigeant",
        "projet",
        "banque",
        "emplois",
        "financement_pi",
        "financement_expert",
        "hypotheses_financieres",
    ]:
        for key, value in common_defaults.items():
            context[section].setdefault(key, value)

    context["entreprise"].setdefault("numero_rc", context["entreprise"].get("rc", ""))
    context["entreprise"].setdefault("denomination", context["entreprise"].get("raison_sociale", ""))
    context["entreprise"].setdefault("adresse", context["entreprise"].get("adresse_siege", ""))
    context["entreprise"].setdefault("secteur_activite", context["entreprise"].get("activite", ""))

    context["dirigeant"].setdefault("fonction", context["dirigeant"].get("qualite", ""))
    context["dirigeant"].setdefault("gsm", context["dirigeant"].get("mobile", ""))

    context["projet"].setdefault("objectif", context["projet"].get("objet", ""))
    context["projet"].setdefault("description", context["projet"].get("objet", ""))
    context["projet"].setdefault("investissement_total_mad", context["projet"].get("investissement_total", 0))
    context["projet"].setdefault("adresse_site", context["projet"].get("adresse_installations", ""))
    context["projet"].setdefault("secteur", context["projet"].get("branche_activite", ""))

    context["banque"].setdefault("banque_partenaire", context["banque"].get("nom", ""))
    context["banque"].setdefault("forme_juridique", context["banque"].get("forme_juridique_banque", "Société Anonyme"))
    context["banque"].setdefault(
        "capital",
        context["banque"].get("capital_social", context["banque"].get("capital_social_banque", ""))
    )
    context["banque"].setdefault("siege", context["banque"].get("siege_social", ""))

    mode = context["projet"].get("mode_financement", {})

    if isinstance(mode, dict):
        fonds = to_number(mode.get("fonds_propres", mode.get("autofinancement", 0)))
        cmt = to_number(mode.get("credit_bancaire", mode.get("cmt", 0)))
        fp = to_number(mode.get("financement_participatif", 0))
        cf = to_number(mode.get("credit_fournisseur", 0))
        leasing = to_number(mode.get("leasing", 0))
    else:
        fonds = to_number(context["financement_expert"].get("fonds_propres_mad", 0))
        cmt = to_number(context["financement_expert"].get("credit_bancaire_mad", 0))
        fp = to_number(context["financement_expert"].get("financement_participatif_mad", 0))
        cf = to_number(context["financement_expert"].get("credit_fournisseur_mad", 0))
        leasing = to_number(context["financement_expert"].get("leasing_mad", 0))

    context["financement_checkbox"].setdefault("autofinancement", "☑" if fonds > 0 else "☐")
    context["financement_checkbox"].setdefault("cmt", "☑" if cmt > 0 else "☐")
    context["financement_checkbox"].setdefault("financement_participatif", "☑" if fp > 0 else "☐")
    context["financement_checkbox"].setdefault("credit_fournisseur", "☑" if cf > 0 else "☐")
    context["financement_checkbox"].setdefault("leasing", "☑" if leasing > 0 else "☐")

    return context


def write_cell(ws, cell_ref, value, value_type="text"):
    cell = ws[cell_ref]

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


def normalize_investissements(data):
    investissements = data.get("investissements", {})

    if isinstance(investissements, dict):
        return investissements

    result = {
        "terrain": [],
        "constructions": [],
        "amenagement_agencement": [],
        "materiel_equipement": [],
        "frais_preliminaires": [],
        "divers_imprevus": [],
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

                write_cell(ws, cell_ref, value, value_type)


def normalize_label(text):
    text = str(text or "").lower().replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^\wÀ-ÿ%()./ -]", "", text)
    return text.strip()


def set_cell_value(cell, value):
    cell.text = "" if value is None else str(value)


def fill_cell_next_to_label(doc, label, value, position="right", occurrence=1):
    wanted = normalize_label(label)
    seen = 0

    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                if wanted and wanted in normalize_label(cell.text):
                    seen += 1

                    if seen < occurrence:
                        continue

                    if position == "below" and r_idx + 1 < len(table.rows):
                        set_cell_value(table.rows[r_idx + 1].cells[c_idx], value)
                        return True

                    offset = 1

                    if isinstance(position, str) and position.startswith("right"):
                        try:
                            offset = int(position.replace("right", "") or "1")
                        except Exception:
                            offset = 1

                    target_index = min(c_idx + offset, len(row.cells) - 1)

                    if target_index != c_idx:
                        set_cell_value(row.cells[target_index], value)
                        return True

    return False


def replace_text_everywhere(doc, search, replace):
    replace = "" if replace is None else str(replace)

    for paragraph in doc.paragraphs:
        if search in paragraph.text:
            paragraph.text = paragraph.text.replace(search, replace)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if search in cell.text:
                    cell.text = cell.text.replace(search, replace)


def apply_dap_default_mappings(doc, context):
    default_mappings = [
        {"label": "Identifiant", "field": "dossier.identifiant"},
        {"label": "Raison sociale", "field": "entreprise.raison_sociale"},
        {"label": "Forme juridique", "field": "entreprise.forme_juridique"},
        {"label": "Date de création", "field": "entreprise.date_creation"},
        {"label": "Secteur d’activité", "field": "entreprise.activite"},
        {"label": "Type et Catégorie", "field": "entreprise.type_categorie"},
        {"label": "Attestation de classement", "field": "entreprise.attestation_classement"},
        {"label": "Capital social (MAD)", "field": "entreprise.capital_social_mad"},
        {"label": "Actionnaires (%)", "field": "entreprise.actionnaires"},
        {"label": "Registre de commerce", "field": "entreprise.rc"},
        {"label": "Identifiant Commun de l'Entreprise", "field": "entreprise.ice"},
        {"label": "CNSS", "field": "entreprise.cnss"},
        {"label": "Adresse du siège social", "field": "entreprise.adresse_siege"},
        {"label": "Adresse du site d’implantation", "field": "projet.adresse_site"},
        {"label": "Site Web", "field": "entreprise.site_web"},
        {"label": "Tél.", "field": "entreprise.telephone", "occurrence": 1},

        {"label": "Dirigeant", "field": "dirigeant.nom"},
        {"label": "Profil", "field": "dirigeant.profil"},
        {"label": "Affaire(s) gérée(s) par le dirigeant", "field": "dirigeant.affaires_gerees"},
        {"label": "Mobile (Cellulaire)", "field": "dirigeant.mobile", "occurrence": 1},
        {"label": "Fixe (Tél.)", "field": "dirigeant.telephone_fixe"},
        {"label": "Courrier électronique", "field": "dirigeant.email"},

        {"label": "Filière(s) du projet", "field": "projet.filieres"},
        {"label": "Objet du projet", "field": "projet.objet"},
        {"label": "Adresse des installations", "field": "projet.adresse_installations"},
        {"label": "Ville/Région du Projet", "field": "projet.ville_region"},
        {"label": "Superficie et mode d’occupation du site", "field": "projet.surface_mode_occupation"},
        {"label": "Effectif à embaucher", "field": "projet.effectif"},
        {"label": "Planning de réalisation", "field": "projet.planning_realisation"},
        {"label": "Date prévue de démarrage", "field": "projet.date_demarrage_prevue"},
        {"label": "Responsable de projet", "field": "projet.responsable_projet"},
        {"label": "Mobile (Cellulaire) du resp. de projet", "field": "projet.responsable_mobile"},
        {"label": "Partenaire financier", "field": "banque.nom"},
        {"label": "Investissement total", "field": "projet.investissement_total"},
        {"label": "Secteur d’activité du projet", "field": "projet.secteur"},
    ]

    for item in default_mappings:
        value = deep_get(context, item["field"], "")

        if value not in ["", None, 0]:
            fill_cell_next_to_label(
                doc,
                item["label"],
                value,
                item.get("position", "right"),
                int(item.get("occurrence", 1)),
            )

    checkbox_values = {
        "Autofinancement": deep_get(context, "financement_checkbox.autofinancement", "☐"),
        "CMT": deep_get(context, "financement_checkbox.cmt", "☐"),
        "Financement participatif": deep_get(context, "financement_checkbox.financement_participatif", "☐"),
        "Crédit fournisseur": deep_get(context, "financement_checkbox.credit_fournisseur", "☐"),
        "Leasing": deep_get(context, "financement_checkbox.leasing", "☐"),
    }

    for label, mark in checkbox_values.items():
        replace_text_everywhere(doc, f"☐ {label}", f"{mark} {label}")


def find_docx_table(doc, anchor_label):
    anchor = normalize_label(anchor_label)

    for table in doc.tables:
        text = normalize_label(" ".join(cell.text for row in table.rows for cell in row.cells))

        if anchor in text:
            return table

    return None


def apply_docx_table_mappings(doc, context, table_mappings):
    for mapping in table_mappings:
        table_label = mapping.get("table_label") or mapping.get("anchor_label") or mapping.get("label")
        source_array = mapping.get("source_array", "")
        columns = mapping.get("columns", [])
        start_row = mapping.get("start_row", mapping.get("start_row_index", 0))

        if not table_label or not source_array or not columns:
            continue

        table = find_docx_table(doc, table_label)

        if table is None:
            continue

        try:
            start_idx = int(start_row)
        except Exception:
            start_idx = 0

        if mapping.get("one_based", True) and start_idx > 0:
            start_idx -= 1

        rows_data = deep_get(context, source_array, [])

        if not isinstance(rows_data, list):
            continue

        for i, row_data in enumerate(rows_data):
            r = start_idx + i

            if r >= len(table.rows):
                break

            for col in columns:
                idx = col.get("index", col.get("column_index", col.get("col")))
                field = col.get("field")

                if idx is None or field is None:
                    continue

                try:
                    idx = int(idx)
                except Exception:
                    continue

                if idx < 0 or idx >= len(table.rows[r].cells):
                    continue

                value = row_data.get(field, "")
                set_cell_value(table.rows[r].cells[idx], value)


def apply_dap_mapping_file(doc, context):
    mapping_path = MAPPING_DIR / "mapping_dap_istitmar.json"

    if not mapping_path.exists():
        apply_dap_default_mappings(doc, context)
        return

    with open(mapping_path, "r", encoding="utf-8") as f:
        mapping = json.load(f)

    label_mappings = (
        mapping.get("cell_mappings", [])
        + mapping.get("label_mappings", [])
        + mapping.get("field_mappings", [])
    )

    for item in label_mappings:
        label = item.get("label") or item.get("search")
        field = item.get("field")
        position = item.get("position", "right")
        occurrence = int(item.get("occurrence", 1))
        value_type = item.get("type", "text")

        if not label or not field:
            continue

        value = normalize_value(deep_get(context, field, ""), value_type)

        if value in ["", None]:
            continue

        fill_cell_next_to_label(doc, label, value, position, occurrence)

    for item in mapping.get("checkbox_mappings", []):
        label = item.get("label")
        field = item.get("field")

        if not label or not field:
            continue

        mark = deep_get(context, field, "☐")
        checked = mark is True or str(mark).strip() in ["☑", "true", "True", "1", "oui", "Oui"]

        replace_text_everywhere(doc, f"☐ {label}", f"{'☑' if checked else '☐'} {label}")

    for item in mapping.get("text_replacements", []) + mapping.get("literal_replacements", []):
        search = item.get("search")
        field = item.get("field")
        replacement = item.get("replacement")

        if not search:
            continue

        if field:
            replacement = deep_get(context, field, "")

        replace_text_everywhere(doc, search, replacement or "")

    apply_docx_table_mappings(doc, context, mapping.get("table_mappings", []))
    apply_dap_default_mappings(doc, context)


def render_dap_with_mapping(output_path, context):
    template_path = TEMPLATE_DIR / "DAP_template.docx"

    if not template_path.exists():
        raise FileNotFoundError("Template Word introuvable : DAP_template.docx")

    safe_context = add_docx_defaults(copy.deepcopy(context))

    tmp_docx = Path(tempfile.gettempdir()) / f"tmp_dap_{uuid.uuid4()}.docx"

    doc_tpl = DocxTemplate(str(template_path))
    doc_tpl.render(safe_context, jinja_env=DOCX_JINJA_ENV)
    doc_tpl.save(str(tmp_docx))

    doc = Document(str(tmp_docx))
    apply_dap_mapping_file(doc, safe_context)
    doc.save(str(output_path))


def render_docx(template_name, output_path, context):
    if template_name == "DAP_template.docx":
        render_dap_with_mapping(output_path, context)
        return

    template_path = TEMPLATE_DIR / template_name

    if not template_path.exists():
        raise FileNotFoundError(f"Template Word introuvable : {template_name}")

    safe_context = add_docx_defaults(copy.deepcopy(context))

    doc = DocxTemplate(str(template_path))
    doc.render(safe_context, jinja_env=DOCX_JINJA_ENV)
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

    apply_simple_mappings(wb, context, mapping.get("scalar_mappings", []))
    apply_simple_mappings(wb, context, mapping.get("financement_pi_mappings", []))
    apply_table_mappings(wb, context, mapping.get("table_mappings", []))
    apply_simple_mappings(wb, context, mapping.get("cpc_mappings", []))
    apply_simple_mappings(wb, context, mapping.get("bilan_mappings", []))
    apply_simple_mappings(wb, context, mapping.get("impacts_mappings", []))

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

    wb.save(output_path)


@app.get("/")
def root():
    return {
        "status": "ok",
        "message": "GO SIYAHA filler API",
        "base_dir": str(BASE_DIR),
        "templates_dir": str(TEMPLATE_DIR),
        "mappings_dir": str(MAPPING_DIR),
    }


@app.get("/health")
def health():
    return {
        "status": "ok",
        "base_dir": str(BASE_DIR),
        "templates": {
            "DAP_template.docx": (TEMPLATE_DIR / "DAP_template.docx").exists(),
            "BP_template.xlsx": (TEMPLATE_DIR / "BP_template.xlsx").exists(),
            "demande_participation.docx": (TEMPLATE_DIR / "demande_participation.docx").exists(),
            "engagement_capacite_financiere.docx": (TEMPLATE_DIR / "engagement_capacite_financiere.docx").exists(),
            "engagement_autorisations.docx": (TEMPLATE_DIR / "engagement_autorisations.docx").exists(),
            "declaration_honneur_justificatifs.docx": (TEMPLATE_DIR / "declaration_honneur_justificatifs.docx").exists(),
        },
        "mappings": {
            "mapping_bp_istitmar.json": (MAPPING_DIR / "mapping_bp_istitmar.json").exists(),
            "mapping_dap_istitmar.json": (MAPPING_DIR / "mapping_dap_istitmar.json").exists(),
        },
    }


@app.post("/fill")
async def fill(request: Request):
    try:
        payload = await request.json()
        data = payload.get("data", payload)

        selected_template = payload.get("selected_template") or data.get("selected_template") or "dossier_complet"

        aliases = {
            "demande_honneur": "declaration_honneur_justificatifs",
            "engagement_capacite": "engagement_capacite_financiere",
            "declaration_factures": "declaration_honneur_justificatifs",
        }

        selected_template = aliases.get(selected_template, selected_template)

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
            "financement_expert": data.get("financement_expert", {}),
            "financement_checkbox": data.get("financement_checkbox", {}),
        }

        context = add_docx_defaults(context)

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

        elif selected_template == "demande_participation":
            add_docx("demande_participation.docx", f"{identifiant}_{societe}_demande_participation.docx")

        elif selected_template == "engagement_capacite_financiere":
            add_docx(
                "engagement_capacite_financiere.docx",
                f"{identifiant}_{societe}_engagement_capacite_financiere.docx",
            )

        elif selected_template == "engagement_autorisations":
            add_docx(
                "engagement_autorisations.docx",
                f"{identifiant}_{societe}_engagement_autorisations.docx",
            )

        elif selected_template == "declaration_honneur_justificatifs":
            add_docx(
                "declaration_honneur_justificatifs.docx",
                f"{identifiant}_{societe}_declaration_honneur_justificatifs.docx",
            )

        elif selected_template == "dossier_complet":
            if (TEMPLATE_DIR / "DAP_template.docx").exists():
                add_docx("DAP_template.docx", f"{identifiant}_{societe}_DAP.docx")

            if (TEMPLATE_DIR / "BP_template.xlsx").exists():
                add_bp()

            if (TEMPLATE_DIR / "demande_participation.docx").exists():
                add_docx("demande_participation.docx", f"{identifiant}_{societe}_demande_participation.docx")

            if (TEMPLATE_DIR / "engagement_capacite_financiere.docx").exists():
                add_docx(
                    "engagement_capacite_financiere.docx",
                    f"{identifiant}_{societe}_engagement_capacite_financiere.docx",
                )

            if (TEMPLATE_DIR / "engagement_autorisations.docx").exists():
                add_docx(
                    "engagement_autorisations.docx",
                    f"{identifiant}_{societe}_engagement_autorisations.docx",
                )

            if (TEMPLATE_DIR / "declaration_honneur_justificatifs.docx").exists():
                add_docx(
                    "declaration_honneur_justificatifs.docx",
                    f"{identifiant}_{societe}_declaration_honneur_justificatifs.docx",
                )

        else:
            return JSONResponse(
                status_code=400,
                content={
                    "error": f"selected_template inconnu : {selected_template}",
                    "allowed": [
                        "bp_excel",
                        "dap_word",
                        "demande_participation",
                        "engagement_capacite_financiere",
                        "engagement_autorisations",
                        "declaration_honneur_justificatifs",
                        "dossier_complet",
                    ],
                },
            )

        if not generated_files:
            return JSONResponse(
                status_code=400,
                content={"error": "Aucun fichier généré. Vérifiez les templates."},
            )

        with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
            for file_path in generated_files:
                zipf.write(file_path, arcname=file_path.name)

        return FileResponse(
            path=str(output_zip),
            media_type="application/zip",
            filename=output_zip.name,
        )

    except FileNotFoundError as e:
        return JSONResponse(status_code=404, content={"error": str(e)})

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
