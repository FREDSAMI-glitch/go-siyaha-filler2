from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, column_index_from_string
from pathlib import Path
from jinja2 import Environment, ChainableUndefined
from docx import Document
from docx.table import _Cell
from docx.shared import Cm
from docx.oxml.ns import qn
import tempfile
import zipfile
import uuid
import json
import re
import copy

app = FastAPI()

API_DIR = Path(__file__).resolve().parent
BASE_DIR = API_DIR.parent

TEMPLATE_DIR = BASE_DIR / "templates"
MAPPING_DIR = BASE_DIR / "mappings"

DOCX_JINJA_ENV = Environment(undefined=ChainableUndefined)


# =========================================================
# OUTILS GÉNÉRAUX
# =========================================================

def safe_name(value):
    if not value:
        return "DOSSIER"
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in str(value))


def as_dict(value):
    return value if isinstance(value, dict) else {}


def is_empty(value):
    return value is None or value == "" or value == [] or value == {}


def find_file(filename):
    possible_paths = [
        TEMPLATE_DIR / filename,
        MAPPING_DIR / filename,
        BASE_DIR / filename,
        API_DIR / filename,
    ]

    for path in possible_paths:
        if path.exists():
            return path

    for folder in [TEMPLATE_DIR, MAPPING_DIR, BASE_DIR, API_DIR]:
        if folder.exists():
            for path in folder.rglob(filename):
                if path.exists():
                    return path

    return None


def deep_get(data, path, default=None):
    if not path:
        return default

    current = data

    for part in str(path).split("."):
        if isinstance(current, dict):
            current = current.get(part, default)
        elif isinstance(current, list):
            try:
                current = current[int(part)]
            except Exception:
                return default
        else:
            return default

    return current if current is not None else default


def get_context_value(context, path, default=""):
    value = deep_get(context, path, None)

    if value not in [None, "", [], {}]:
        return value

    dap = context.get("dap", {})
    if isinstance(dap, dict):
        value = deep_get(dap, path, None)
        if value not in [None, "", [], {}]:
            return value

    return default


def parse_number(value):
    if value is None or value == "":
        return None

    if isinstance(value, bool):
        return int(value)

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip()
    text = text.replace("\xa0", " ")
    text = re.sub(r"[^\d,.\-]", "", text)

    if not text:
        return None

    if text.count(",") > 1 and "." not in text:
        text = text.replace(",", "")
    elif text.count(".") > 1 and "," not in text:
        text = text.replace(".", "")
    elif "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text and "." not in text:
        text = text.replace(",", ".")

    try:
        number = float(text)
        if number.is_integer():
            return int(number)
        return number
    except Exception:
        return None


def to_number(value):
    number = parse_number(value)
    return 0 if number is None else number


def format_display_value(value, fmt=None):
    if value is None:
        return ""

    if fmt in ["mad", "money"]:
        number = to_number(value)
        if number == 0:
            return ""
        return f"{number:,.0f}".replace(",", " ") + " MAD"

    if fmt == "kmad":
        number = to_number(value)
        if number == 0:
            return ""
        return f"{number:,.0f}".replace(",", " ") + " KMAD"

    if fmt == "percent":
        number = to_number(value)
        if number == 0:
            return ""
        if number <= 1:
            number = number * 100
        return f"{number:.0f}%"

    if fmt in ["number", "integer", "float"]:
        number = to_number(value)
        return str(int(number)) if number == int(number) else str(number)

    return str(value)


def normalize_excel_value(value, mapping_item, source_path=None, category=None):
    if is_empty(value):
        if "default_by_category" in mapping_item:
            defaults = mapping_item["default_by_category"]
            value = defaults.get(category, defaults.get("default"))
        elif "default" in mapping_item:
            value = mapping_item.get("default")
        else:
            return None

    value_type = mapping_item.get("type")
    unit = str(mapping_item.get("unit", "")).upper()

    if value_type == "list":
        return normalize_list_value(value, mapping_item.get("allowed_values", []))

    if value_type in ["number", "integer", "float", "money"]:
        number = parse_number(value)

        if number is None:
            return None

        if unit == "KMAD":
            if source_path and source_path.endswith("_mad"):
                number = number / 1000
            elif abs(number) >= 100000:
                number = number / 1000

        return number

    return value


def normalize_list_value(value, allowed_values):
    if is_empty(value):
        return None

    raw = str(value).strip()
    raw_low = raw.lower()

    for allowed in allowed_values:
        if raw == allowed:
            return allowed

    for allowed in allowed_values:
        if raw_low == allowed.lower():
            return allowed

    if "animation" in raw_low:
        for allowed in allowed_values:
            if "animation" in allowed.lower():
                return allowed

    if "restaurant" in raw_low or "restauration" in raw_low:
        for allowed in allowed_values:
            if "restaurant" in allowed.lower():
                return allowed

    if "hébergement" in raw_low or "hebergement" in raw_low or "hotel" in raw_low:
        for allowed in allowed_values:
            if "hébergement" in allowed.lower() or "hebergement" in allowed.lower():
                return allowed

    if "création" in raw_low or "creation" in raw_low:
        for allowed in allowed_values:
            if "création" in allowed.lower() or "creation" in allowed.lower():
                return allowed

    if "extension" in raw_low:
        for allowed in allowed_values:
            if "extension" in allowed.lower():
                return allowed

    return raw


# =========================================================
# CONTEXTE PAR DÉFAUT
# =========================================================

def add_context_defaults(context):
    context.setdefault("dossier", {})
    context.setdefault("entreprise", {})
    context.setdefault("dirigeant", {})
    context.setdefault("projet", {})
    context.setdefault("banque", {})
    context.setdefault("emplois", {})
    context.setdefault("investissements", {})
    context.setdefault("financement_pi", {})
    context.setdefault("financement_expert", {})
    context.setdefault("financement_checkbox", {})
    context.setdefault("hypotheses_financieres", {})
    context.setdefault("dap", {})

    for key in [
        "dossier",
        "entreprise",
        "dirigeant",
        "projet",
        "banque",
        "emplois",
        "financement_pi",
        "financement_expert",
        "financement_checkbox",
        "hypotheses_financieres",
        "dap",
    ]:
        context[key] = as_dict(context.get(key))

    common_defaults = {
        "identifiant": "",
        "date_dossier": "",
        "lieu_signature": "",
        "date_signature": "",
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
        "type_activite": "",
        "categorie": "",
        "secteur": "",
        "secteur_activite": "",

        "nom": "",
        "prenom": "",
        "cin": "",
        "qualite": "",
        "fonction": "",
        "mobile": "",
        "gsm": "",
        "fixe": "",
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
        "description_offre_animation": "",
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
        "superficie_mode_occupation": "",
        "planning_realisation": "",
        "date_demarrage_prevue": "",
        "annee_demarrage": "",
        "responsable_projet": "",
        "responsable_nom": "",
        "responsable_mobile": "",
        "secteur_activite_projet": "",
        "role_region": "",
        "role_balance_commerciale": "",
        "facteurs_differenciation": "",
        "attractivite_touristique": "",

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
    context["dirigeant"].setdefault("fixe", context["dirigeant"].get("telephone_fixe", ""))

    context["projet"].setdefault("objectif", context["projet"].get("objet", ""))
    context["projet"].setdefault("description", context["projet"].get("objet", ""))
    context["projet"].setdefault("investissement_total_mad", context["projet"].get("investissement_total", 0))
    context["projet"].setdefault("adresse_site", context["projet"].get("adresse_installations", ""))
    context["projet"].setdefault("secteur", context["projet"].get("branche_activite", "Animation touristique"))
    context["projet"].setdefault("secteur_activite_projet", context["projet"].get("secteur", "Animation touristique"))
    context["projet"].setdefault("responsable_nom", context["dirigeant"].get("nom", ""))
    context["projet"].setdefault("responsable_mobile", context["dirigeant"].get("mobile", ""))

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


# =========================================================
# BP EXCEL
# =========================================================

def build_candidate_paths(mapping_item):
    paths = []

    for path in mapping_item.get("field_candidates", []):
        if path and path not in paths:
            paths.append(path)

    field = mapping_item.get("field")
    if field and field not in paths:
        paths.append(field)

    extra = []

    for path in paths:
        if path.endswith("_kmad"):
            extra.append(path[:-5] + "_mad")
        if ".financement_expert." in path:
            extra.append(path.replace(".financement_expert.", ".financement."))
        if path.startswith("financement_expert."):
            extra.append(path.replace("financement_expert.", "financement."))

    for path in extra:
        if path not in paths:
            paths.append(path)

    return paths


def get_value_from_mapping(data, mapping_item):
    for path in build_candidate_paths(mapping_item):
        value = get_context_value(data, path, None)

        if not is_empty(value):
            return value, path

    if "default" in mapping_item:
        return mapping_item.get("default"), "__default__"

    return None, None


def cell_to_row_col(cell_ref):
    match = re.match(r"^([A-Z]+)([0-9]+)$", str(cell_ref).upper())

    if not match:
        return None, None

    col = column_index_from_string(match.group(1))
    row = int(match.group(2))

    return row, col


def build_blocked_ranges(mapping):
    blocked = {}

    for rule in mapping.get("never_write", []):
        sheet = rule.get("sheet")

        if not sheet:
            continue

        blocked.setdefault(sheet, [])

        for rng in rule.get("ranges", []):
            try:
                min_col, min_row, max_col, max_row = range_boundaries(rng)
                blocked[sheet].append((min_row, min_col, max_row, max_col, rng))
            except Exception:
                pass

    return blocked


def is_blocked_cell(sheet_name, cell_ref, blocked_ranges):
    row, col = cell_to_row_col(cell_ref)

    if row is None:
        return True

    for min_row, min_col, max_row, max_col, _ in blocked_ranges.get(sheet_name, []):
        if min_row <= row <= max_row and min_col <= col <= max_col:
            return True

    return False


def write_excel_cell(ws, cell_ref, value, blocked_ranges):
    if is_blocked_cell(ws.title, cell_ref, blocked_ranges):
        return False

    cell = ws[cell_ref]

    if cell.data_type == "f" or (isinstance(cell.value, str) and cell.value.startswith("=")):
        return False

    if value is None:
        return False

    cell.value = value

    return True


def write_mapping_section(wb, data, mapping, section_name, blocked_ranges):
    count = 0

    for item in mapping.get(section_name, []):
        sheet_name = item.get("sheet")
        cell_ref = item.get("cell")

        if not sheet_name or not cell_ref:
            continue

        if sheet_name not in wb.sheetnames:
            continue

        value, source_path = get_value_from_mapping(data, item)
        value = normalize_excel_value(value, item, source_path=source_path)

        if value is None:
            continue

        if write_excel_cell(wb[sheet_name], cell_ref, value, blocked_ranges):
            count += 1

    return count


def normalize_category_name(value):
    if not value:
        return ""

    text = str(value).lower()
    text = text.replace("é", "e").replace("è", "e").replace("ê", "e")
    text = text.replace("à", "a").replace("â", "a")
    text = text.replace("ç", "c")
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")

    if "terrain" in text:
        return "terrain"
    if "construction" in text:
        return "constructions"
    if "amenagement" in text or "agencement" in text:
        return "amenagement_agencement"
    if "materiel" in text or "equipement" in text:
        return "materiel_equipement"
    if "frais" in text and "prelim" in text:
        return "frais_preliminaires"
    if "divers" in text or "imprevu" in text:
        return "divers_imprevus"

    return text


def get_bp_source_array(data, source_array, category):
    arr = get_context_value(data, source_array, None)

    if isinstance(arr, list):
        return arr

    if isinstance(arr, dict):
        return list(arr.values())

    investissements = data.get("investissements")

    if isinstance(investissements, dict):
        possible = investissements.get(category)

        if isinstance(possible, list):
            return possible

    if isinstance(investissements, list):
        filtered = []

        for item in investissements:
            if not isinstance(item, dict):
                continue

            item_category = normalize_category_name(
                item.get("categorie")
                or item.get("category")
                or item.get("type")
                or item.get("rubrique")
            )

            if item_category == category:
                filtered.append(item)

        return filtered

    return []


def write_bp_table_mappings(wb, data, mapping, blocked_ranges):
    count = 0

    for table in mapping.get("table_mappings", []):
        sheet_name = table.get("sheet")
        category = table.get("category")
        start_row = table.get("start_row")
        end_row = table.get("end_row")

        if sheet_name not in wb.sheetnames:
            continue

        if start_row is None or end_row is None:
            continue

        ws = wb[sheet_name]
        rows = get_bp_source_array(data, table.get("source_array"), category)

        if not isinstance(rows, list):
            continue

        max_rows = max(0, int(end_row) - int(start_row) + 1)

        for index, row_data in enumerate(rows[:max_rows]):
            excel_row = int(start_row) + index

            for col_map in table.get("columns", []):
                col_letter = col_map.get("column")
                field = col_map.get("field")

                if not col_letter or not field:
                    continue

                value = deep_get(row_data, field, None)

                if is_empty(value):
                    if "default_by_category" in col_map:
                        defaults = col_map["default_by_category"]
                        value = defaults.get(category, defaults.get("default"))
                    elif "default" in col_map:
                        value = col_map["default"]
                    else:
                        continue

                value = normalize_excel_value(value, col_map, source_path=field, category=category)
                cell_ref = f"{col_letter}{excel_row}"

                if write_excel_cell(ws, cell_ref, value, blocked_ranges):
                    count += 1

    return count


def force_excel_recalculation(wb):
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass


def render_bp_excel(output_path, context):
    template_path = find_file("BP_template.xlsx")
    mapping_path = find_file("mapping_bp_istitmar.json")

    if template_path is None:
        raise FileNotFoundError("BP_template.xlsx introuvable dans /templates")

    if mapping_path is None:
        raise FileNotFoundError("mapping_bp_istitmar.json introuvable dans /mappings")

    with open(mapping_path, "r", encoding="utf-8") as f:
        mapping = json.load(f)

    wb = load_workbook(template_path)
    blocked_ranges = build_blocked_ranges(mapping)

    written = {
        "scalar_mappings": write_mapping_section(wb, context, mapping, "scalar_mappings", blocked_ranges),
        "financement_pi_mappings": write_mapping_section(wb, context, mapping, "financement_pi_mappings", blocked_ranges),
        "table_mappings": write_bp_table_mappings(wb, context, mapping, blocked_ranges),
        "cpc_mappings": write_mapping_section(wb, context, mapping, "cpc_mappings", blocked_ranges),
        "bilan_mappings": write_mapping_section(wb, context, mapping, "bilan_mappings", blocked_ranges),
        "impacts_mappings": write_mapping_section(wb, context, mapping, "impacts_mappings", blocked_ranges),
    }

    force_excel_recalculation(wb)
    wb.save(output_path)

    return written


# =========================================================
# WORD / OUTILS ROBUSTES
# =========================================================

def normalize_label(text):
    text = str(text or "").lower().replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^\wÀ-ÿ%()./ -]", "", text)
    return text.strip()


def safe_table_rows(table):
    try:
        return list(table.rows)
    except Exception:
        return []


def safe_row_cells(row, table):
    cells = []

    try:
        for child in row._tr.iterchildren():
            if child.tag == qn("w:tc"):
                cells.append(_Cell(child, row))
            elif child.tag == qn("w:sdt"):
                for sdt_child in child.iter():
                    if sdt_child.tag == qn("w:tc"):
                        cells.append(_Cell(sdt_child, row))
    except Exception:
        pass

    if cells:
        return cells

    try:
        return list(row.cells)
    except Exception:
        return []


def get_table_by_index(doc, index):
    try:
        return doc.tables[int(index)]
    except Exception:
        return None


def iter_all_paragraphs(doc):
    for paragraph in doc.paragraphs:
        yield paragraph

    for table in doc.tables:
        rows = safe_table_rows(table)

        for row in rows:
            cells = safe_row_cells(row, table)

            for cell in cells:
                try:
                    for paragraph in cell.paragraphs:
                        yield paragraph
                except Exception:
                    continue


def set_cell_value(cell, value):
    try:
        cell.text = "" if value is None else str(value)
    except Exception:
        pass


def set_paragraph_text(paragraph, value):
    try:
        paragraph.text = "" if value is None else str(value)
    except Exception:
        pass


def replace_text_everywhere(doc, search, replace):
    if not search:
        return

    replace = "" if replace is None else str(replace)

    for paragraph in iter_all_paragraphs(doc):
        try:
            if search in paragraph.text:
                paragraph.text = paragraph.text.replace(search, replace)
        except Exception:
            continue


def fill_cell_next_to_label(doc, label, value, position="right", occurrence=1):
    wanted = normalize_label(label)
    seen = 0

    for table in doc.tables:
        rows = safe_table_rows(table)

        for r_idx, row in enumerate(rows):
            cells = safe_row_cells(row, table)

            for c_idx, cell in enumerate(cells):
                try:
                    cell_text = normalize_label(cell.text)
                except Exception:
                    continue

                if wanted and wanted in cell_text:
                    seen += 1

                    if seen < occurrence:
                        continue

                    if position == "below" and r_idx + 1 < len(rows):
                        below_cells = safe_row_cells(rows[r_idx + 1], table)

                        if c_idx < len(below_cells):
                            set_cell_value(below_cells[c_idx], value)
                            return True

                    offset = 1

                    if isinstance(position, str) and position.startswith("right"):
                        try:
                            offset = int(position.replace("right", "") or "1")
                        except Exception:
                            offset = 1

                    target_index = c_idx + offset

                    if target_index < len(cells):
                        set_cell_value(cells[target_index], value)
                        return True

    return False


def find_docx_table(doc, anchor_label):
    anchor = normalize_label(anchor_label)

    for table in doc.tables:
        texts = []
        rows = safe_table_rows(table)

        for row in rows:
            cells = safe_row_cells(row, table)

            for cell in cells:
                try:
                    texts.append(cell.text)
                except Exception:
                    continue

        table_text = normalize_label(" ".join(texts))

        if anchor in table_text:
            return table

    return None


# =========================================================
# DAP WORD AVANCÉ
# =========================================================

def get_expert_default(mapping, field):
    defaults = mapping.get("expert_generation_defaults", {})

    if field in defaults:
        return defaults[field]

    return ""


def generate_default_rows(source_array, context):
    key = str(source_array).split(".")[-1]

    projet = context.get("projet", {})
    entreprise = context.get("entreprise", {})
    emplois = context.get("emplois", {})

    objet = projet.get("objet") or projet.get("description") or "activité d’animation touristique"
    ville = projet.get("ville_region") or projet.get("region") or "la zone d’implantation"
    ca = to_number(projet.get("ca_prevu_annee_1", 0))
    effectif = int(to_number(projet.get("effectif") or emplois.get("directs") or emplois.get("emplois_directs") or 5))

    if key == "gamme_services":
        return [
            {
                "domaine_activite": "Animation touristique",
                "description": objet,
                "marche_adressable": f"Touristes nationaux, touristes étrangers et clientèle locale à {ville}",
                "pourcentage_ca": "100%",
                "image_slot": "Image illustrative à ajouter",
            }
        ]

    if key == "marche_cibles":
        return [
            {
                "domaine_activite": "Animation touristique",
                "marche_cible": "Touristes nationaux, touristes internationaux, familles, groupes et clientèle locale",
                "taille_marche_mmad": "",
                "tcam": "",
                "source_taille_marche": "Estimation interne",
                "source_tcam": "Estimation interne",
            }
        ]

    if key == "concurrents":
        return [
            {
                "nom": "Opérateurs touristiques locaux",
                "implantation": ville,
                "categorie": "Concurrence directe et indirecte",
                "part_marche": "",
                "principaux_services": "Activités touristiques, loisirs, animation et expériences clients",
            }
        ]

    if key == "fournisseurs":
        return [
            {
                "nom": "Fournisseurs locaux et nationaux",
                "categorie": "Équipements, consommables et maintenance",
                "produits_services": "Matériel d’exploitation, maintenance, fournitures et prestations de support",
                "part_achats": "",
                "delai_reglement_jours": "30",
                "modalites_achat": "Achat sur devis, bon de commande et règlement selon conditions négociées",
            }
        ]

    if key == "clients_creneaux":
        return [
            {
                "creneau_service": "Clientèle touristique",
                "principaux_clients": "Touristes nationaux, touristes étrangers, familles, groupes et agences partenaires",
                "categorie": "B2C/B2B",
                "part_ca": "100%",
                "delai_recouvrement_jours": "0",
                "modalites_vente": "Paiement comptant, réservation en ligne et partenariats commerciaux",
            }
        ]

    if key == "politique_prix":
        return [
            {
                "famille_service": "Animation touristique",
                "strategie_prix": "Prix aligné sur le marché avec différenciation par la qualité de l’expérience",
                "commentaire": "Tarification modulable selon la saison, les groupes, les offres packagées et les partenariats",
            }
        ]

    if key == "capacites_animation":
        return [
            {
                "domaine_activite": "Animation touristique",
                "capacite_journaliere": "À préciser",
                "duree_cycle_service": "À préciser",
                "prix_moyen_service": "À préciser",
                "taux_occupation": "50%",
            }
        ]

    if key == "emplois_directs_profils":
        return [
            {
                "profil_domaine": "Exploitation et animation",
                "emplois_permanents": effectif,
                "emplois_feminins": "",
                "contrat_cdi": effectif,
                "contrat_anapec": "",
            }
        ]

    if key == "lignes":
        return [
            {
                "libelle": "Animation touristique",
                "taux_croissance": format_display_value(projet.get("croissance_ca", 0.15), "percent"),
                "2026_kmad": round(ca / 1000) if ca else "",
                "2027_kmad": round(ca * 1.15 / 1000) if ca else "",
                "2028_kmad": round(ca * 1.15**2 / 1000) if ca else "",
                "2029_kmad": round(ca * 1.15**3 / 1000) if ca else "",
                "2030_kmad": round(ca * 1.15**4 / 1000) if ca else "",
                "2031_kmad": round(ca * 1.15**5 / 1000) if ca else "",
            }
        ]

    return []


def get_dap_source_array(context, source_array, mapping_item=None):
    data = get_context_value(context, source_array, None)

    if isinstance(data, list):
        return data

    if isinstance(data, dict):
        return [data]

    if mapping_item and mapping_item.get("generate_defaults_if_empty"):
        return generate_default_rows(source_array, context)

    generated = generate_default_rows(source_array, context)

    if generated:
        return generated

    return []


def apply_scalar_text_replacements(doc, context, mapping):
    for item in mapping.get("scalar_text_replacements", []):
        placeholder = item.get("placeholder") or item.get("search")
        field = item.get("field")
        fmt = item.get("format") or item.get("type")

        if not placeholder:
            continue

        value = get_context_value(context, field, item.get("default", "")) if field else item.get("replacement", "")

        if value in [None, ""]:
            value = item.get("default", "")

        value = format_display_value(value, fmt)
        replace_text_everywhere(doc, placeholder, value)


def apply_legacy_text_replacements(doc, context, mapping):
    for item in mapping.get("text_replacements", []) + mapping.get("literal_replacements", []):
        search = item.get("search")
        field = item.get("field")
        replacement = item.get("replacement")

        if not search:
            continue

        if field:
            replacement = get_context_value(context, field, "")

        replace_text_everywhere(doc, search, replacement or "")


def apply_legacy_label_mappings(doc, context, mapping):
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
        fmt = item.get("format") or item.get("type")

        if not label or not field:
            continue

        value = get_context_value(context, field, "")

        if value in ["", None]:
            continue

        fill_cell_next_to_label(doc, label, format_display_value(value, fmt), position, occurrence)


def apply_table_cell_mappings(doc, context, mapping):
    for item in mapping.get("table_cell_mappings", []):
        table = get_table_by_index(doc, item.get("table_index"))

        if table is None:
            continue

        rows = safe_table_rows(table)
        row_idx = int(item.get("row", 0))

        if row_idx < 0 or row_idx >= len(rows):
            continue

        cells = safe_row_cells(rows[row_idx], table)

        if not cells:
            continue

        target = item.get("target", "right")
        fmt = item.get("format") or item.get("type")

        if target == "split_cells":
            fields = item.get("fields", [])

            for i, field in enumerate(fields):
                cell_idx = i + 1

                if cell_idx < len(cells):
                    value = get_context_value(context, field, "")
                    set_cell_value(cells[cell_idx], format_display_value(value, fmt))

            continue

        if target == "third_cell_append":
            field = item.get("field")
            value = get_context_value(context, field, "")

            if len(cells) >= 3 and value not in ["", None]:
                old = cells[2].text.strip()
                new_text = (old + " " + format_display_value(value, fmt)).strip()
                set_cell_value(cells[2], new_text)

            continue

        if target == "authorization_row":
            array_path = item.get("array")
            array_index = int(item.get("array_index", 0))
            rows_data = get_dap_source_array(context, array_path)

            if 0 <= array_index < len(rows_data):
                row_data = rows_data[array_index]

                if isinstance(row_data, dict):
                    value = " | ".join(str(v) for v in row_data.values() if v not in ["", None])
                else:
                    value = str(row_data)

                if len(cells) > 1:
                    set_cell_value(cells[1], value)

            continue

        field = item.get("field")
        value = get_context_value(context, field, "")

        if item.get("generate_if_empty") and value in ["", None]:
            value = get_expert_default(mapping, field)

        value = format_display_value(value, fmt)

        if len(cells) < 2:
            continue

        set_cell_value(cells[1], value)


def fill_image_cell(cell, image_path, fallback_text, width_cm=4.5, height_cm=3.0):
    if image_path:
        p = Path(str(image_path))

        if p.exists():
            try:
                cell.text = ""
                paragraph = cell.paragraphs[0]
                run = paragraph.add_run()
                run.add_picture(str(p), width=Cm(float(width_cm)), height=Cm(float(height_cm)))
                return
            except Exception:
                pass

    set_cell_value(cell, fallback_text or "Image illustrative à ajouter")


def apply_repeat_table_mappings(doc, context, mapping):
    image_settings_by_table = {}

    for img in mapping.get("image_mappings", []):
        if "table_index" in img:
            image_settings_by_table[int(img["table_index"])] = img

    for item in mapping.get("repeat_table_mappings", []):
        table = get_table_by_index(doc, item.get("table_index"))

        if table is None:
            continue

        rows = safe_table_rows(table)
        source_array = item.get("source_array", "")
        data_rows = get_dap_source_array(context, source_array, item)

        if not data_rows:
            continue

        start_row = int(item.get("start_row", 0))
        row_step = int(item.get("row_step", 1))
        max_rows = int(item.get("max_rows", len(data_rows)))
        columns = item.get("columns", [])

        for i, row_data in enumerate(data_rows[:max_rows]):
            row_idx = start_row + i * row_step

            if row_idx >= len(rows):
                break

            cells = safe_row_cells(rows[row_idx], table)

            for col in columns:
                col_idx = col.get("col", col.get("index", col.get("column_index")))
                field = col.get("field")
                fmt = col.get("format") or col.get("type")

                try:
                    col_idx = int(col_idx)
                except Exception:
                    continue

                if col_idx < 0 or col_idx >= len(cells):
                    continue

                if col.get("type") == "image":
                    img_conf = image_settings_by_table.get(int(item.get("table_index", -1)), {})
                    image_field = img_conf.get("image_field", field)
                    image_path = row_data.get(image_field, "")
                    fallback = img_conf.get("fallback_text", "Image illustrative à ajouter")
                    fill_image_cell(
                        cells[col_idx],
                        image_path,
                        fallback,
                        img_conf.get("insert_width_cm", 4.5),
                        img_conf.get("insert_height_cm", 3.0),
                    )
                    continue

                value = row_data.get(field, "")
                set_cell_value(cells[col_idx], format_display_value(value, fmt))

            source_row_columns = item.get("source_row_columns", [])

            if source_row_columns and row_step >= 2 and row_idx + 1 < len(rows):
                source_cells = safe_row_cells(rows[row_idx + 1], table)

                for src in source_row_columns:
                    col_idx = src.get("col")
                    field = src.get("field")
                    prefix = src.get("prefix", "")

                    try:
                        col_idx = int(col_idx)
                    except Exception:
                        continue

                    if col_idx < 0 or col_idx >= len(source_cells):
                        continue

                    value = row_data.get(field, "")

                    if value not in ["", None]:
                        set_cell_value(source_cells[col_idx], prefix + str(value))

        total_row = item.get("total_row")

        if isinstance(total_row, dict):
            row_idx = int(total_row.get("row", 0))

            if 0 <= row_idx < len(rows):
                cells = safe_row_cells(rows[row_idx], table)

                for col in total_row.get("columns", []):
                    col_idx = col.get("col")
                    field = col.get("field")
                    fmt = col.get("format") or col.get("type")

                    try:
                        col_idx = int(col_idx)
                    except Exception:
                        continue

                    if col_idx < 0 or col_idx >= len(cells):
                        continue

                    value = get_context_value(context, field, "")

                    if value not in ["", None]:
                        set_cell_value(cells[col_idx], format_display_value(value, fmt))


def apply_single_row_table_mappings(doc, context, mapping):
    for item in mapping.get("single_row_table_mappings", []):
        table = get_table_by_index(doc, item.get("table_index"))

        if table is None:
            continue

        rows = safe_table_rows(table)

        if not rows:
            continue

        row_idx = int(item.get("row", 1))

        if row_idx >= len(rows):
            row_idx = 0

        cells = safe_row_cells(rows[row_idx], table)

        target_col = item.get("target_col")
        field = item.get("field")

        try:
            target_col = int(target_col)
        except Exception:
            target_col = None

        if target_col is not None and 0 <= target_col < len(cells):
            value = get_context_value(context, field, "")
            set_cell_value(cells[target_col], format_display_value(value, item.get("format")))

        target_col_2 = item.get("target_col_2")
        field_2 = item.get("field_2")

        try:
            target_col_2 = int(target_col_2)
        except Exception:
            target_col_2 = None

        if target_col_2 is not None and 0 <= target_col_2 < len(cells):
            value_2 = get_context_value(context, field_2, "")

            if item.get("generate_field_2_if_empty") and value_2 in ["", None]:
                value_2 = get_expert_default(mapping, field_2)

            set_cell_value(cells[target_col_2], value_2)


def apply_paragraph_mappings(doc, context, mapping):
    for item in mapping.get("paragraph_mappings", []):
        search = item.get("search_contains")
        field = item.get("field")

        if not search or not field:
            continue

        value = get_context_value(context, field, "")

        if item.get("generate_if_empty") and value in ["", None]:
            value = get_expert_default(mapping, field)

        if value in ["", None]:
            continue

        search_norm = normalize_label(search)

        for paragraph in iter_all_paragraphs(doc):
            try:
                if search_norm in normalize_label(paragraph.text):
                    set_paragraph_text(paragraph, value)
            except Exception:
                continue


def apply_checkbox_mappings(doc, context, mapping):
    for item in mapping.get("checkbox_mappings", []):
        label = item.get("label")
        field = item.get("field")

        if not label or not field:
            continue

        raw = get_context_value(context, field, "☐")
        checked = raw is True or str(raw).strip() in ["☑", "true", "True", "1", "oui", "Oui"]

        checked_value = item.get("checked_value", "☑")
        unchecked_value = item.get("unchecked_value", "☐")
        mark = checked_value if checked else unchecked_value

        label_norm = normalize_label(label)

        for paragraph in iter_all_paragraphs(doc):
            try:
                txt_norm = normalize_label(paragraph.text)

                if label_norm in txt_norm and ("☐" in paragraph.text or "□" in paragraph.text or "☑" in paragraph.text):
                    paragraph.text = f"{mark} {label}"
            except Exception:
                continue


def apply_dap_default_mappings(doc, context):
    default_mappings = [
        {"label": "Identifiant", "field": "dossier.identifiant"},
        {"label": "Raison sociale", "field": "entreprise.raison_sociale"},
        {"label": "Forme juridique", "field": "entreprise.forme_juridique"},
        {"label": "Date de création", "field": "entreprise.date_creation"},
        {"label": "Secteur d’activité", "field": "entreprise.activite"},
        {"label": "Capital social (MAD)", "field": "entreprise.capital_social_mad"},
        {"label": "Registre de commerce", "field": "entreprise.rc"},
        {"label": "Identifiant Commun de l'Entreprise", "field": "entreprise.ice"},
        {"label": "CNSS", "field": "entreprise.cnss"},
        {"label": "Adresse du siège social", "field": "entreprise.adresse_siege"},
        {"label": "Dirigeant", "field": "dirigeant.nom"},
        {"label": "Mobile (Cellulaire)", "field": "dirigeant.mobile"},
        {"label": "Courrier électronique", "field": "dirigeant.email"},
        {"label": "Objet du projet", "field": "projet.objet"},
        {"label": "Ville/Région du Projet", "field": "projet.ville_region"},
        {"label": "Effectif à embaucher", "field": "projet.effectif"},
        {"label": "Partenaire financier", "field": "banque.nom"},
        {"label": "Investissement total", "field": "projet.investissement_total"},
    ]

    for item in default_mappings:
        value = get_context_value(context, item["field"], "")

        if value not in ["", None, 0]:
            fill_cell_next_to_label(doc, item["label"], value)


def apply_dap_mapping_file(doc, context):
    mapping_path = find_file("mapping_dap_istitmar.json")

    if mapping_path is None:
        apply_dap_default_mappings(doc, context)
        return {"mapping_used": False, "reason": "mapping_dap_istitmar.json introuvable"}

    with open(mapping_path, "r", encoding="utf-8") as f:
        mapping = json.load(f)

    apply_scalar_text_replacements(doc, context, mapping)
    apply_legacy_text_replacements(doc, context, mapping)
    apply_legacy_label_mappings(doc, context, mapping)
    apply_table_cell_mappings(doc, context, mapping)
    apply_repeat_table_mappings(doc, context, mapping)
    apply_single_row_table_mappings(doc, context, mapping)
    apply_paragraph_mappings(doc, context, mapping)
    apply_checkbox_mappings(doc, context, mapping)
    apply_dap_default_mappings(doc, context)

    if mapping.get("cleanup_enabled") is True:
        for placeholder in mapping.get("cleanup_placeholders", []):
            replace_text_everywhere(doc, placeholder, "")

    return {
        "mapping_used": True,
        "mapping_path": str(mapping_path),
        "mapping_name": mapping.get("mapping_name"),
        "version": mapping.get("version"),
    }


def render_dap_docx(output_path, context):
    template_path = find_file("DAP_template.docx")

    if template_path is None:
        raise FileNotFoundError("DAP_template.docx introuvable dans /templates")

    safe_context = add_context_defaults(copy.deepcopy(context))
    tmp_docx = Path(tempfile.gettempdir()) / f"tmp_dap_{uuid.uuid4()}.docx"

    doc_tpl = DocxTemplate(str(template_path))
    doc_tpl.render(safe_context, jinja_env=DOCX_JINJA_ENV)
    doc_tpl.save(str(tmp_docx))

    doc = Document(str(tmp_docx))
    info = apply_dap_mapping_file(doc, safe_context)
    doc.save(str(output_path))

    return info


# =========================================================
# WORD / DOCUMENTS JURIDIQUES
# =========================================================

def apply_legal_doc_defaults(doc, context):
    dossier = context.get("dossier", {})
    entreprise = context.get("entreprise", {})
    dirigeant = context.get("dirigeant", {})
    banque = context.get("banque", {})

    nom = dirigeant.get("nom", "")
    qualite = dirigeant.get("qualite", dirigeant.get("fonction", ""))
    cin = dirigeant.get("cin", "")
    raison = entreprise.get("raison_sociale", "")
    forme = entreprise.get("forme_juridique", "")
    rc = entreprise.get("rc", "")
    lieu = dossier.get("lieu_signature", "")
    date = dossier.get("date_signature", "")

    banque_nom = banque.get("nom", banque.get("banque_partenaire", ""))
    banque_forme = banque.get("forme_juridique", banque.get("forme_juridique_banque", "Société Anonyme"))
    banque_capital = banque.get("capital", banque.get("capital_social_banque", ""))
    banque_siege = banque.get("siege", banque.get("siege_social", ""))

    for paragraph in iter_all_paragraphs(doc):
        text = paragraph.text

        if "Je soussigné(e)" in text and "agissant au nom" in text:
            paragraph.text = (
                f"Je soussigné(e) {nom} ({qualite}), "
                f"agissant au nom et pour le compte de {raison} {forme} :"
            )

        elif "Je soussigné (prénom, nom)" in text:
            paragraph.text = (
                f"Je soussigné {nom} en sa qualité de {qualite}, "
                f"signataire de la convention d’investissement avec MAROC PME, "
                f"titulaire de la CIN N° {cin} et agissant au nom et pour le compte "
                f"de la société {raison} {forme}, inscrite au registre du commerce N° {rc} ;"
            )

        elif "Autorise ma banque partenaire" in text:
            paragraph.text = (
                f"• Autorise ma banque partenaire, {banque_nom}, {banque_forme} "
                f"au capital de {banque_capital} dirhams et dont le siège social est sis au {banque_siege}, "
                f"à transférer à Maroc PME, dans le cadre du Programme GO SIYAHA – volet soutien à l’investissement, "
                f"mon dossier de candidature, le dossier bancaire (rating, accord subordonné pour le prêt bancaire, "
                f"et le schéma de financement bancaire retenu) et une copie du contrat de crédit signé qu’elle aura constitué ;"
            )

        elif "Objet" in text and "justificatifs déposés" in text:
            paragraph.text = (
                f"Objet : Déclaration sur l’honneur relative aux justificatifs déposés par la société {raison} "
                f"pour sa candidature au programme GO SIYAHA – Volet soutien à l’investissement."
            )

        elif "Fait à" in text and ("Le" in text or "le" in text or "………" in text):
            paragraph.text = f"Fait à : {lieu}    Le : {date}"


def render_legal_docx(template_name, output_path, context):
    template_path = find_file(template_name)

    if template_path is None:
        raise FileNotFoundError(f"{template_name} introuvable dans /templates")

    safe_context = add_context_defaults(copy.deepcopy(context))
    tmp_docx = Path(tempfile.gettempdir()) / f"tmp_docx_{uuid.uuid4()}.docx"

    doc_tpl = DocxTemplate(str(template_path))
    doc_tpl.render(safe_context, jinja_env=DOCX_JINJA_ENV)
    doc_tpl.save(str(tmp_docx))

    doc = Document(str(tmp_docx))
    apply_legal_doc_defaults(doc, safe_context)
    doc.save(str(output_path))


# =========================================================
# ROUTES DEBUG
# =========================================================

@app.get("/")
def root():
    return {
        "status": "ok",
        "service": "GO SIYAHA filler API",
        "base_dir": str(BASE_DIR),
    }


@app.get("/health")
def health():
    return {
        "status": "ok",
        "base_dir": str(BASE_DIR),
        "templates": {
            "BP_template.xlsx": find_file("BP_template.xlsx") is not None,
            "DAP_template.docx": find_file("DAP_template.docx") is not None,
            "demande_participation.docx": find_file("demande_participation.docx") is not None,
            "engagement_capacite_financiere.docx": find_file("engagement_capacite_financiere.docx") is not None,
            "engagement_autorisations.docx": find_file("engagement_autorisations.docx") is not None,
            "declaration_honneur_justificatifs.docx": find_file("declaration_honneur_justificatifs.docx") is not None,
        },
        "mappings": {
            "mapping_bp_istitmar.json": find_file("mapping_bp_istitmar.json") is not None,
            "mapping_dap_istitmar.json": find_file("mapping_dap_istitmar.json") is not None,
        },
    }


@app.get("/debug-bp")
def debug_bp():
    template_path = find_file("BP_template.xlsx")
    mapping_path = find_file("mapping_bp_istitmar.json")

    if template_path is None or mapping_path is None:
        return JSONResponse(
            status_code=404,
            content={
                "status": "error",
                "BP_template.xlsx": template_path is not None,
                "mapping_bp_istitmar.json": mapping_path is not None,
            }
        )

    try:
        with open(mapping_path, "r", encoding="utf-8") as f:
            mapping = json.load(f)

        return {
            "status": "ok",
            "template_path": str(template_path),
            "mapping_path": str(mapping_path),
            "mapping_name": mapping.get("mapping_name"),
            "version": mapping.get("version"),
            "sections": {
                "scalar_mappings": len(mapping.get("scalar_mappings", [])),
                "financement_pi_mappings": len(mapping.get("financement_pi_mappings", [])),
                "table_mappings": len(mapping.get("table_mappings", [])),
                "cpc_mappings": len(mapping.get("cpc_mappings", [])),
                "bilan_mappings": len(mapping.get("bilan_mappings", [])),
                "impacts_mappings": len(mapping.get("impacts_mappings", [])),
            },
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"status": "error", "error": str(e)}
        )


@app.get("/debug-mapping")
def debug_mapping():
    mapping_path = find_file("mapping_dap_istitmar.json")

    if mapping_path is None:
        return JSONResponse(
            status_code=404,
            content={"status": "error", "error": "mapping_dap_istitmar.json introuvable"}
        )

    try:
        text = mapping_path.read_text(encoding="utf-8")
        data = json.loads(text)

        return {
            "status": "ok",
            "mapping_path": str(mapping_path),
            "version": data.get("version"),
            "mapping_name": data.get("mapping_name"),
            "sections": list(data.keys()),
        }

    except Exception as e:
        lines = text.splitlines() if "text" in locals() else []

        return JSONResponse(
            status_code=500,
            content={
                "status": "error",
                "mapping_path": str(mapping_path),
                "error": str(e),
                "around_line_281": lines[276:286],
            }
        )


# =========================================================
# ROUTE PRINCIPALE
# =========================================================

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

        aliases = {
            "bp": "bp_excel",
            "dap": "dap_word",
            "demande_honneur": "declaration_honneur_justificatifs",
            "engagement_capacite": "engagement_capacite_financiere",
            "declaration_factures": "declaration_honneur_justificatifs",
        }

        selected_template = aliases.get(selected_template, selected_template)

        context = add_context_defaults(copy.deepcopy(data))
        context["selected_template"] = selected_template

        dossier = context["dossier"]
        entreprise = context["entreprise"]

        identifiant = safe_name(dossier.get("identifiant", "DOSSIER"))
        societe = safe_name(entreprise.get("raison_sociale", "SOCIETE"))

        job_id = str(uuid.uuid4())
        tmp_dir = Path(tempfile.gettempdir()) / f"go_siyaha_{job_id}"
        tmp_dir.mkdir(parents=True, exist_ok=True)

        output_zip = tmp_dir / f"{identifiant}_{societe}_GO_SIYAHA.zip"
        generated_files = []
        generation_errors = []
        debug_info = {
            "selected_template": selected_template,
            "generated_files": [],
            "errors": [],
        }

        def add_bp_safe():
            try:
                output_path = tmp_dir / f"{identifiant}_{societe}_BP_GO_SIYAHA.xlsx"
                written = render_bp_excel(output_path, context)
                generated_files.append(output_path)
                debug_info["bp_written"] = written
                debug_info["generated_files"].append(output_path.name)
            except Exception as e:
                error = f"BP_template.xlsx : {str(e)}"
                generation_errors.append(error)
                debug_info["errors"].append(error)

        def add_dap_safe():
            try:
                output_path = tmp_dir / f"{identifiant}_{societe}_DAP_GO_SIYAHA.docx"
                info = render_dap_docx(output_path, context)
                generated_files.append(output_path)
                debug_info["dap_info"] = info
                debug_info["generated_files"].append(output_path.name)
            except Exception as e:
                error = f"DAP_template.docx : {str(e)}"
                generation_errors.append(error)
                debug_info["errors"].append(error)

        def add_legal_safe(template_name, output_suffix):
            try:
                output_path = tmp_dir / f"{identifiant}_{societe}_{output_suffix}.docx"
                render_legal_docx(template_name, output_path, context)
                generated_files.append(output_path)
                debug_info["generated_files"].append(output_path.name)
            except Exception as e:
                error = f"{template_name} : {str(e)}"
                generation_errors.append(error)
                debug_info["errors"].append(error)

        if selected_template == "bp_excel":
            add_bp_safe()

        elif selected_template == "dap_word":
            add_dap_safe()

        elif selected_template == "demande_participation":
            add_legal_safe("demande_participation.docx", "demande_participation")

        elif selected_template == "engagement_capacite_financiere":
            add_legal_safe("engagement_capacite_financiere.docx", "engagement_capacite_financiere")

        elif selected_template == "engagement_autorisations":
            add_legal_safe("engagement_autorisations.docx", "engagement_autorisations")

        elif selected_template == "declaration_honneur_justificatifs":
            add_legal_safe("declaration_honneur_justificatifs.docx", "declaration_honneur_justificatifs")

        elif selected_template in ["dossier_complet", "all", "tout"]:
            add_dap_safe()
            add_bp_safe()
            add_legal_safe("demande_participation.docx", "demande_participation")
            add_legal_safe("engagement_capacite_financiere.docx", "engagement_capacite_financiere")
            add_legal_safe("engagement_autorisations.docx", "engagement_autorisations")
            add_legal_safe("declaration_honneur_justificatifs.docx", "declaration_honneur_justificatifs")

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

        if generation_errors:
            error_report = tmp_dir / "rapport_erreurs_generation.txt"
            error_report.write_text(
                "Erreurs de génération GO SIYAHA\n\n" + "\n".join(generation_errors),
                encoding="utf-8",
            )
            generated_files.append(error_report)

        debug_path = tmp_dir / "debug_generation.json"
        debug_path.write_text(
            json.dumps(debug_info, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        generated_files.append(debug_path)

        if not generated_files:
            return JSONResponse(
                status_code=400,
                content={"error": "Aucun fichier généré."},
            )

        with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
            for file_path in generated_files:
                zipf.write(file_path, arcname=file_path.name)

        return FileResponse(
            path=str(output_zip),
            media_type="application/zip",
            filename=output_zip.name,
        )

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e)},
        )
