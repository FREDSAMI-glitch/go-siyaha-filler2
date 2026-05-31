from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from pathlib import Path
from jinja2 import Environment, ChainableUndefined
from docx import Document
from docx.table import _Cell
from docx.shared import Cm
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


# =========================================================
# OUTILS GÉNÉRAUX
# =========================================================

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

    if value is not None and value != "":
        return value

    dap = context.get("dap", {})

    if isinstance(dap, dict):
        value = deep_get(dap, path, None)

        if value is not None and value != "":
            return value

    return default


def to_number(value):
    if value is None or value == "":
        return 0

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip()
    text = text.replace(" ", "").replace("\u00a0", "").replace(",", ".")

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


def format_display_value(value, fmt=None):
    if value is None:
        return ""

    if fmt in ["mad", "money"]:
        n = to_number(value)
        if n == 0:
            return ""
        return f"{n:,.0f}".replace(",", " ") + " MAD"

    if fmt == "percent":
        n = to_number(value)
        if n == 0:
            return ""
        if n <= 1:
            n = n * 100
        return f"{n:.0f}%"

    if fmt == "mmad_an":
        n = to_number(value)
        if n == 0:
            return ""
        return f"{n:,.1f}".replace(",", " ") + " MMAD/an"

    if fmt in ["number", "integer", "float"]:
        n = to_number(value)
        return str(int(n)) if n == int(n) else str(n)

    return str(value)


def normalize_value(value, value_type="text"):
    if value_type in ["number", "integer", "float"]:
        return to_number(value)

    if value_type == "boolean":
        return bool(value)

    if value_type in ["money", "mad"]:
        return format_display_value(value, "mad")

    return value if value is not None else ""


# =========================================================
# CONTEXTE PAR DÉFAUT
# =========================================================

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
    context["projet"].setdefault("secteur", context["projet"].get("branche_activite", ""))
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
# EXCEL / BP
# =========================================================

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
        value = get_context_value(context, field, "")
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
    try:
        return list(row.cells)
    except Exception:
        try:
            return [_Cell(tc, table) for tc in row._tr.tc_lst]
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
# DAP AVANCÉ
# =========================================================

def get_expert_default(mapping, field):
    defaults = mapping.get("expert_generation_defaults", {})

    if field in defaults:
        return defaults[field]

    return ""


def generate_default_rows(source_array, context):
    projet = context.get("projet", {})
    entreprise = context.get("entreprise", {})
    emplois = context.get("emplois", {})

    objet = projet.get("objet") or projet.get("description") or "activité touristique"
    ville = projet.get("ville_region") or "la région ciblée"
    ca = to_number(projet.get("ca_prevu_annee_1", 0))
    effectif = int(to_number(projet.get("effectif") or emplois.get("directs") or 5))

    if source_array == "gamme_services":
        return [
            {
                "domaine_activite": "Animation touristique",
                "description": objet,
                "marche_adressable": f"Touristes et clientèle locale à {ville}",
                "pourcentage_ca": "100%",
                "image_slot": "Image illustrative à ajouter",
            }
        ]

    if source_array == "marche_cibles":
        return [
            {
                "domaine_activite": "Animation touristique",
                "marche_cible": "Touristes nationaux et internationaux",
                "taille_marche_mmad": "",
                "tcam": "",
                "source_taille_marche": "Estimation interne",
                "source_tcam": "Estimation interne",
            }
        ]

    if source_array == "concurrents":
        return [
            {
                "nom": "Opérateurs locaux",
                "implantation": ville,
                "categorie": "Concurrence directe/indirecte",
                "part_marche": "",
                "principaux_services": "Activités touristiques et loisirs",
            }
        ]

    if source_array == "fournisseurs":
        return [
            {
                "nom": "Fournisseurs locaux",
                "categorie": "Équipements et consommables",
                "produits_services": "Matériel, maintenance, services opérationnels",
                "part_achats": "",
                "delai_reglement_jours": "30",
                "modalites_achat": "Devis, bon de commande et règlement selon conditions négociées",
            }
        ]

    if source_array == "clients_creneaux":
        return [
            {
                "creneau_service": "Clientèle touristique",
                "principaux_clients": "Touristes nationaux, touristes étrangers, familles et groupes",
                "categorie": "B2C/B2B",
                "part_ca": "100%",
                "delai_recouvrement_jours": "0",
                "modalites_vente": "Paiement comptant, réservation en ligne et partenariats",
            }
        ]

    if source_array == "politique_prix":
        return [
            {
                "famille_service": "Animation touristique",
                "strategie_prix": "Prix aligné sur le marché avec différenciation par qualité d’expérience",
                "commentaire": "Tarification modulable selon saison, groupes et partenariats touristiques",
            }
        ]

    if source_array == "capacites_animation":
        return [
            {
                "domaine_activite": "Animation touristique",
                "capacite_journaliere": "À préciser",
                "duree_cycle_service": "À préciser",
                "prix_moyen_service": "À préciser",
                "taux_occupation": "50%",
            }
        ]

    if source_array == "emplois_directs_profils":
        return [
            {
                "profil_domaine": "Exploitation et animation",
                "emplois_permanents": effectif,
                "emplois_feminins": "",
                "contrat_cdi": effectif,
                "contrat_anapec": "",
            }
        ]

    if source_array == "ca_previsionnel_dap.lignes":
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


def get_source_array(context, source_array, mapping_item=None):
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
        placeholder = item.get("placeholder")
        field = item.get("field")
        fmt = item.get("format") or item.get("type")

        if not placeholder:
            continue

        value = get_context_value(context, field, item.get("default", "")) if field else item.get("replacement", "")

        if value in [None, ""]:
            value = item.get("default", "")

        value = format_display_value(value, fmt)
        replace_text_everywhere(doc, placeholder, value)


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
            rows_data = get_source_array(context, array_path)

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

        if target == "right":
            target_idx = 1 if len(cells) > 1 else 0
        else:
            target_idx = 1 if len(cells) > 1 else 0

        set_cell_value(cells[target_idx], value)


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
        data_rows = get_source_array(context, source_array, item)

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

        replace_text_everywhere(doc, f"☐ {label}", f"{mark} {label}")
        replace_text_everywhere(doc, f"□ {label}", f"{mark} {label}")


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
        value_type = item.get("type", "text")

        if not label or not field:
            continue

        value = normalize_value(get_context_value(context, field, ""), value_type)

        if value in ["", None]:
            continue

        fill_cell_next_to_label(doc, label, value, position, occurrence)


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


def apply_cleanup_placeholders(doc, mapping):
    for placeholder in mapping.get("cleanup_placeholders", []):
        if placeholder and len(str(placeholder)) > 2:
            replace_text_everywhere(doc, placeholder, "")


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
    mapping_path = MAPPING_DIR / "mapping_dap_istitmar.json"

    if not mapping_path.exists():
        apply_dap_default_mappings(doc, context)
        return

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
    apply_cleanup_placeholders(doc, mapping)


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


# =========================================================
# WORD / DOCUMENTS JURIDIQUES
# =========================================================

def apply_legal_doc_defaults(doc, context, template_name):
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

    banque_nom = banque.get("nom", "")
    banque_forme = banque.get("forme_juridique", "Société Anonyme")
    banque_capital = banque.get("capital", "")
    banque_siege = banque.get("siege", "")

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
                f"pour sa candidature au programme Go SIYAHA – Volet soutien à l’investissement :"
            )

        elif "Fait à" in text and ("Le" in text or "le" in text or "………" in text):
            paragraph.text = f"Fait à : {lieu}    Le : {date}"


def render_docx(template_name, output_path, context):
    if template_name == "DAP_template.docx":
        render_dap_with_mapping(output_path, context)
        return

    template_path = TEMPLATE_DIR / template_name

    if not template_path.exists():
        raise FileNotFoundError(f"Template Word introuvable : {template_name}")

    safe_context = add_docx_defaults(copy.deepcopy(context))
    tmp_docx = Path(tempfile.gettempdir()) / f"tmp_docx_{uuid.uuid4()}.docx"

    doc_tpl = DocxTemplate(str(template_path))
    doc_tpl.render(safe_context, jinja_env=DOCX_JINJA_ENV)
    doc_tpl.save(str(tmp_docx))

    doc = Document(str(tmp_docx))
    apply_legal_doc_defaults(doc, safe_context, template_name)
    doc.save(str(output_path))


# =========================================================
# ROUTES
# =========================================================

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

@app.get("/debug-mapping")
def debug_mapping():
    mapping_path = MAPPING_DIR / "mapping_dap_istitmar.json"

    try:
        text = mapping_path.read_text(encoding="utf-8")
        data = json.loads(text)

        return {
            "status": "ok",
            "mapping_path": str(mapping_path),
            "version": data.get("version"),
            "mapping_name": data.get("mapping_name"),
            "sections": list(data.keys())
        }

    except Exception as e:
        lines = text.splitlines() if "text" in locals() else []

        return JSONResponse(
            status_code=500,
            content={
                "status": "error",
                "mapping_path": str(mapping_path),
                "error": str(e),
                "around_line_281": lines[276:286]
            }
        )
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
            "dap": data.get("dap", {}),
            "gamme_services": data.get("gamme_services", []),
            "marche_cibles": data.get("marche_cibles", []),
            "concurrents": data.get("concurrents", []),
            "fournisseurs": data.get("fournisseurs", []),
            "clients_creneaux": data.get("clients_creneaux", []),
            "politique_prix": data.get("politique_prix", []),
            "capacites_animation": data.get("capacites_animation", []),
            "emplois_directs_profils": data.get("emplois_directs_profils", []),
            "recettes_devises": data.get("recettes_devises", []),
            "achats_devises": data.get("achats_devises", {}),
            "ca_previsionnel_dap": data.get("ca_previsionnel_dap", {}),
            "synthese_etude_marche": data.get("synthese_etude_marche", ""),
            "strategie_approvisionnement": data.get("strategie_approvisionnement", {}),
            "strategie_commerciale": data.get("strategie_commerciale", {}),
            "facteurs_differenciation": data.get("facteurs_differenciation", ""),
            "attractivite_touristique": data.get("attractivite_touristique", ""),
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
        generation_errors = []

        def add_docx_safe(template_name, output_name):
            try:
                output_path = tmp_dir / output_name
                render_docx(template_name, output_path, context)
                generated_files.append(output_path)
            except Exception as e:
                generation_errors.append(f"{template_name} : {str(e)}")

        def add_bp_safe():
            try:
                output_path = tmp_dir / f"{identifiant}_{societe}_BP.xlsx"
                render_bp_excel(output_path, context)
                generated_files.append(output_path)
            except Exception as e:
                generation_errors.append(f"BP_template.xlsx : {str(e)}")

        if selected_template == "bp_excel":
            add_bp_safe()

        elif selected_template == "dap_word":
            add_docx_safe("DAP_template.docx", f"{identifiant}_{societe}_DAP.docx")

        elif selected_template == "demande_participation":
            add_docx_safe("demande_participation.docx", f"{identifiant}_{societe}_demande_participation.docx")

        elif selected_template == "engagement_capacite_financiere":
            add_docx_safe(
                "engagement_capacite_financiere.docx",
                f"{identifiant}_{societe}_engagement_capacite_financiere.docx",
            )

        elif selected_template == "engagement_autorisations":
            add_docx_safe(
                "engagement_autorisations.docx",
                f"{identifiant}_{societe}_engagement_autorisations.docx",
            )

        elif selected_template == "declaration_honneur_justificatifs":
            add_docx_safe(
                "declaration_honneur_justificatifs.docx",
                f"{identifiant}_{societe}_declaration_honneur_justificatifs.docx",
            )

        elif selected_template == "dossier_complet":
            if (TEMPLATE_DIR / "DAP_template.docx").exists():
                add_docx_safe("DAP_template.docx", f"{identifiant}_{societe}_DAP.docx")

            if (TEMPLATE_DIR / "BP_template.xlsx").exists():
                add_bp_safe()

            if (TEMPLATE_DIR / "demande_participation.docx").exists():
                add_docx_safe("demande_participation.docx", f"{identifiant}_{societe}_demande_participation.docx")

            if (TEMPLATE_DIR / "engagement_capacite_financiere.docx").exists():
                add_docx_safe(
                    "engagement_capacite_financiere.docx",
                    f"{identifiant}_{societe}_engagement_capacite_financiere.docx",
                )

            if (TEMPLATE_DIR / "engagement_autorisations.docx").exists():
                add_docx_safe(
                    "engagement_autorisations.docx",
                    f"{identifiant}_{societe}_engagement_autorisations.docx",
                )

            if (TEMPLATE_DIR / "declaration_honneur_justificatifs.docx").exists():
                add_docx_safe(
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

        if generation_errors:
            error_report = tmp_dir / "rapport_erreurs_generation.txt"
            error_report.write_text(
                "Erreurs de génération GO SIYAHA\n\n" + "\n".join(generation_errors),
                encoding="utf-8"
            )
            generated_files.append(error_report)

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
