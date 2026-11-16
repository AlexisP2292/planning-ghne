import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import datetime
import re
import random
import calendar
import unicodedata
import math
import json
import io

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ==============================================================================
# 1. CONFIGURATION STREAMLIT & CHARTE GRAPHIQUE
# ==============================================================================
st.set_page_config(
    page_title="Planificateur Automatique GHNE Paris Saclay",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

PRIMARY_COLOR = "#003A54"
SECONDARY_COLOR = "#1A758E"
ACCENT_COLOR = "#00A896"
RED_ALERT = "#D9534F"
BG_LIGHT = "#F4F7F9"
BG_URGENCES = "#E3F2FD"
BG_SMUR = "#E8EAF6"
BORDER_URGENCES = "#90CAF9"
BORDER_SMUR = "#C5CAE9"

FRENCH_MONTHS = {
    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril",
    5: "Mai", 6: "Juin", 7: "Juillet", 8: "Août",
    9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
}

FRENCH_DAYS_FULL = [
    "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"
]
FRENCH_DAYS_SHORT = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]

st.markdown(f"""
<style>
.main {{
    background-color: {BG_LIGHT};
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
}}
.ghne-header {{
    background: linear-gradient(135deg, {PRIMARY_COLOR} 0%, {SECONDARY_COLOR} 100%);
    color: white;
    padding: 12px 20px;
    border-radius: 8px;
    margin-bottom: 15px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    display: flex;
    justify-content: space-between;
    align-items: center;
}}
.ghne-title {{
    font-size: 20px;
    font-weight: 700;
    margin: 0;
    letter-spacing: 0.5px;
}}
.ghne-subtitle {{
    font-size: 12px;
    opacity: 0.9;
    margin-top: 2px;
}}
.metric-card {{
    background-color: white;
    padding: 10px 14px;
    border-radius: 8px;
    border-left: 4px solid {SECONDARY_COLOR};
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    margin-bottom: 10px;
}}
.metric-title {{
    font-size: 10px;
    text-transform: uppercase;
    color: #6C757D;
    font-weight: 700;
    margin-bottom: 2px;
}}
.metric-value {{
    font-size: 18px;
    font-weight: bold;
    color: {PRIMARY_COLOR};
}}
.stTabs [data-baseweb="tab-list"] {{
    gap: 8px;
}}
.stTabs [data-baseweb="tab"] {{
    background-color: #E2E8F0;
    border-radius: 6px 6px 0px 0px;
    padding: 8px 16px;
    font-weight: 600;
    color: {PRIMARY_COLOR};
}}
.stTabs [aria-selected="true"] {{
    background-color: {PRIMARY_COLOR} !important;
    color: white !important;
}}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="ghne-header">
    <div>
        <div class="ghne-title">🏥 GROUPE HOSPITALIER NORD-ESSONNE</div>
        <div class="ghne-subtitle">Hôpital Paris-Saclay | Planificateur Automatique Urgences &amp; SMUR</div>
    </div>
    <div style="text-align: right; font-size: 11px; background: rgba(255,255,255,0.15); padding: 6px 12px; border-radius: 6px;">
        <b>Pôle Urgences / SAMU</b><br/>Génération &amp; Reporting
    </div>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. DÉFINITION DES POSTES, HORAIRES ET COMBINAISONS STRICTEMENT AUTORISÉES
# ==============================================================================
SHIFTS_CONFIG = {
    # SMUR
    "SMUR_JOUR": {
        "service": "SMUR", "type": "JOUR", "debut": "08:00", "fin": "18:00",
        "heures": 10.0, "nom": "SMUR Journée"
    },
    "SMUR_NUIT": {
        "service": "SMUR", "type": "NUIT", "debut": "18:00", "fin": "08:00",
        "heures": 14.0, "nom": "SMUR Nuit"
    },
    # URGENCES JOURNÉE
    "UHCD_1": {
        "service": "URGENCES", "type": "JOUR", "debut": "08:00", "fin": "18:00",
        "heures": 10.0, "nom": "UHCD 1"
    },
    "UHCD_2": {
        "service": "URGENCES", "type": "JOUR", "debut": "08:00", "fin": "18:00",
        "heures": 10.0, "nom": "UHCD 2"
    },
    "MRA": {
        "service": "URGENCES", "type": "JOUR", "debut": "08:30", "fin": "18:00",
        "heures": 9.5, "nom": "MRA", "req_mra": True
    },
    "COURS_DE_CIRCUIT": {
        "service": "URGENCES", "type": "JOUR", "debut": "08:30", "fin": "18:00",
        "heures": 9.5, "nom": "Circuit Court"
    },
    "CIRCUIT_LONG_1": {
        "service": "URGENCES", "type": "JOUR", "debut": "08:30", "fin": "18:00",
        "heures": 9.5, "nom": "Circuit Long 1"
    },
    "CIRCUIT_LONG_2": {
        "service": "URGENCES", "type": "JOUR", "debut": "08:30", "fin": "18:00",
        "heures": 9.5, "nom": "Circuit Long 2"
    },
    "SAUV": {
        "service": "URGENCES", "type": "JOUR", "debut": "08:30", "fin": "18:00",
        "heures": 9.5, "nom": "SAUV", "req_sauv": True
    },
    # URGENCES NUIT & DEMI-GARDES
    "DEMI_TRANS": {
        "service": "URGENCES", "type": "DEMI_GARDE", "debut": "18:00", "fin": "00:00",
        "heures": 6.0, "nom": "Demi Garde Trans"
    },
    "DEMI_COURT": {
        "service": "URGENCES", "type": "DEMI_GARDE", "debut": "18:00", "fin": "00:00",
        "heures": 6.0, "nom": "Demi Garde Court"
    },
    "GARDE_1": {
        "service": "URGENCES", "type": "GARDE", "debut": "18:00", "fin": "08:30",
        "heures": 14.5, "nom": "Garde 1"
    },
    "GARDE_2": {
        "service": "URGENCES", "type": "GARDE", "debut": "18:00", "fin": "08:30",
        "heures": 14.5, "nom": "Garde 2"
    },
    "GARDE_3_SAUV": {
        "service": "URGENCES", "type": "GARDE", "debut": "18:00", "fin": "08:30",
        "heures": 14.5, "nom": "Garde 3 SAUV", "req_sauv": True
    }
}

MATRIX_COLUMNS_ORDER = [
    "SMUR_JOUR", "SMUR_NUIT",
    "UHCD_1", "UHCD_2", "MRA", "COURS_DE_CIRCUIT",
    "CIRCUIT_LONG_1", "CIRCUIT_LONG_2", "SAUV",
    "DEMI_TRANS", "DEMI_COURT",
    "GARDE_1", "GARDE_2", "GARDE_3_SAUV"
]

VALID_COMBOS = [
    (["CIRCUIT_LONG_1", "CIRCUIT_LONG_2"], ["DEMI_TRANS"], "accepte_demi_garde"),
    (["MRA", "SAUV", "COURS_DE_CIRCUIT"], ["DEMI_COURT"], "accepte_demi_garde"),
    (["UHCD_1", "UHCD_2"], ["GARDE_1", "GARDE_2", "GARDE_3_SAUV"], "accepte_24h"),
    (["UHCD_1", "UHCD_2", "MRA"], ["SMUR_NUIT"], "mix_urg_smur_and_24h"),
]

DAY_SHIFTS_ORDER = [
    "SMUR_JOUR", "UHCD_1", "UHCD_2", "MRA", "SAUV",
    "COURS_DE_CIRCUIT", "CIRCUIT_LONG_1", "CIRCUIT_LONG_2"
]

ETAGE_DAY_SHIFTS = [
    "UHCD_1", "UHCD_2", "COURS_DE_CIRCUIT", "CIRCUIT_LONG_1", "CIRCUIT_LONG_2"
]

NIGHT_GARDES_ORDER = [
    "SMUR_NUIT", "GARDE_3_SAUV", "GARDE_1", "GARDE_2"
]

DEMI_GARDES_ORDER = [
    "DEMI_TRANS", "DEMI_COURT"
]

# ==============================================================================
# 3. OUTILS DE NORMALISATION, VALIDATION DES MÉDECINS & PARSING EXCEL
# ==============================================================================
EXCLUDED_DOCTOR_KEYWORDS = {
    "dit", "dits", "dite", "dites", "veille", "veilles", "total", "totaux",
    "soustotal", "statut", "legende", "observation", "remarque", "commentaire",
    "poste", "postes", "garde", "gardes", "demigarde", "smur", "urgences",
    "urgence", "horaire", "horaires", "date", "dates", "semaine", "mois",
    "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche",
    "stagiaire", "interne", "externe", "remplacant", "interimaire",
    "effectif", "couverture", "nonpourvu", "vacant", "noms", "nom", "prenom"
}


def strip_accents(text):
    """Supprime les accents d'une chaîne de caractères."""
    if not isinstance(text, str):
        return ""
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def clean_str(s):
    """Normalise une chaîne en minuscules sans accents ni espaces."""
    if not isinstance(s, str):
        return ""
    s_no_acc = strip_accents(str(s)).strip().lower()
    return re.sub(r"[^a-z0-9]+", "", s_no_acc)


def tokenize_name(s):
    """Découpe un nom en jetons alphanumériques normalisés sans accents."""
    if not isinstance(s, str):
        return []
    s_no_acc = strip_accents(str(s)).strip().lower()
    return [tok for tok in re.split(r"[^a-z0-9]+", s_no_acc) if tok]


def is_valid_doctor_name(name):
    """Vérifie strictement qu'un libellé correspond à un vrai médecin."""
    if name is None:
        return False
    raw = str(name).strip()
    if not raw or raw.lower() in {"nan", "none", "nat", "null", "-"}:
        return False
    cleaned = clean_str(raw)
    if len(cleaned) < 2:
        return False
    if cleaned.isdigit():
        return False
    if cleaned in EXCLUDED_DOCTOR_KEYWORDS:
        return False
    tokens = tokenize_name(raw)
    if not tokens:
        return False
    forbidden_tokens = {
        "dit", "dits", "dite", "dites", "veille", "veilles", "total", "totaux",
        "soustotal", "legende", "statut", "observation", "remarque", "commentaire",
        "horaire", "horaires", "effectif", "couverture", "nonpourvu", "vacant",
        "stagiaire", "interne", "externe", "remplacant", "interimaire",
        "noms", "nom", "prenom"
    }
    if any(t in forbidden_tokens for t in tokens):
        return False
    if tokens[0] in EXCLUDED_DOCTOR_KEYWORDS:
        return False
    return True


def match_doctor_name(name_c, list_desid_names, already_matched=None):
    """
    Apparie un prénom/nom du fichier Contraintes (ex: 'Roxane', 'Elodie Flore', 'Loik')
    avec la liste 'Nom Prénom' du fichier Désidératas (ex: 'PARIN Roxane').
    """
    if already_matched is None:
        already_matched = set()
    c_tokens = tokenize_name(name_c)
    c_clean = clean_str(name_c)
    if not c_clean:
        return None

    available_desid = [d for d in list_desid_names if d not in already_matched]
    search_pool = available_desid if available_desid else list(list_desid_names)

    # 1. Correspondance exacte normalisée
    for d_full in search_pool:
        if clean_str(d_full) == c_clean:
            return d_full

    # 2. Correspondance exacte sur un ou plusieurs tokens complets
    best_match = None
    best_score = 0
    for d_full in search_pool:
        d_tokens = tokenize_name(d_full)
        common = set(c_tokens).intersection(set(d_tokens))
        if common:
            score = sum(len(t) for t in common) * 10
            if len(common) == len(c_tokens):
                score += 60
            # Bonus si le prénom correspond aux derniers tokens de "NOM Prénom"
            if len(d_tokens) >= len(c_tokens) and d_tokens[-len(c_tokens):] == c_tokens:
                score += 40
            if score > best_score:
                best_score = score
                best_match = d_full

    if best_match is not None:
        return best_match

    # 3. Correspondance par inclusion si prénom composé / tiret
    if len(c_clean) >= 3:
        for d_full in search_pool:
            d_clean = clean_str(d_full)
            if c_clean in d_clean or (len(d_clean) >= 3 and d_clean in c_clean):
                return d_full

    return None


def is_yes_value(val):
    """Interprète une cellule Oui/Non de manière souple et insensible aux accents."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return False
    s = strip_accents(str(val)).strip().lower()
    return s in {"oui", "yes", "true", "1", "o", "vrai", "x"} or s.startswith("oui")


def is_no_value(val):
    """Interprète une cellule Non de manière souple."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return False
    s = strip_accents(str(val)).strip().lower()
    return s in {"non", "n", "false", "0", "faux"} or s.startswith("non")


def parse_mois_travailles(file_obj, doc_constraints):
    """Lit le fichier des mois travaillés et retourne les mois indisponibles par médecin."""
    if file_obj is None:
        return {}
    df_m = pd.read_excel(file_obj)
    df_m.columns = [str(c).strip() for c in df_m.columns]
    name_col = df_m.columns[0]
    unavail_dict = {}
    month_map = {}
    for c in df_m.columns[1:]:
        c_norm = strip_accents(c).lower()
        for m_idx, m_name in FRENCH_MONTHS.items():
            if strip_accents(m_name).lower() in c_norm or str(m_idx) == c_norm or f"{m_idx:02d}" in c_norm:
                month_map[c] = m_idx
                break

    # Dictionnaire de recherche combinant nom court (Contraintes) et nom complet (Désidératas)
    lookup_to_c_doc = {}
    for c_doc, info in doc_constraints.items():
        lookup_to_c_doc[c_doc] = c_doc
        desid_full = info.get("desid_matched_name")
        if desid_full:
            lookup_to_c_doc[desid_full] = c_doc

    all_candidates = list(lookup_to_c_doc.keys())
    for _, row in df_m.iterrows():
        raw_name = str(row.get(name_col, "")).strip()
        if not is_valid_doctor_name(raw_name):
            continue
        matched_key = match_doctor_name(raw_name, all_candidates)
        if matched_key:
            c_doc = lookup_to_c_doc[matched_key]
            unavail = set()
            for col, m_idx in month_map.items():
                if is_no_value(row.get(col)):
                    unavail.add(m_idx)
            unavail_dict[c_doc] = unavail
    return unavail_dict


def parse_prev_weekend(file_obj, doc_constraints):
    """
    Lit le planning du dernier week-end précédent (format NOM + Initiale du prénom,
    y compris multi-noms dans une même cellule) et déduit les repos de sécurité
    en tout début de période (Lundi D0 / Mardi D1).
    """
    if file_obj is None:
        return {}
    try:
        df_pw = pd.read_excel(file_obj, header=None)
    except Exception:
        return {}

    # Construire les signatures "NOM + Initiale" pour chaque médecin
    doc_signatures = {}
    for c_doc, info in doc_constraints.items():
        sigs = set()
        full_name = info.get("desid_matched_name") or c_doc
        toks = tokenize_name(full_name)
        if len(toks) >= 2:
            # Format standard "NOM Prénom" -> "nomp" et "nom p"
            nom_part = "".join(toks[:-1])
            prenom_init = toks[-1][0]
            sigs.add(f"{nom_part}{prenom_init}")
            # Cas inverse "Prénom NOM"
            sigs.add(f"{''.join(toks[1:])}{toks[0][0]}")
        elif len(toks) == 1:
            sigs.add(toks[0])
        c_toks = tokenize_name(c_doc)
        if c_toks:
            sigs.add("".join(c_toks))
        doc_signatures[c_doc] = {s for s in sigs if len(s) >= 3}

    # Repérer les colonnes Samedi / Dimanche et les lignes Jour / Nuit
    sat_cols, sun_cols = set(), set()
    for r in range(min(15, df_pw.shape[0])):
        for c in range(df_pw.shape[1]):
            val = df_pw.iloc[r, c]
            if pd.isna(val):
                continue
            s_norm = strip_accents(str(val)).lower()
            if "samedi" in s_norm or re.search(r"\bsam\b", s_norm):
                sat_cols.add(c)
            elif "dimanche" in s_norm or re.search(r"\bdim\b", s_norm):
                sun_cols.add(c)

    # Si les colonnes Samedi/Dimanche ne sont pas explicites, prendre la dernière colonne comme Dimanche
    if not sun_cols and df_pw.shape[1] >= 2:
        sun_cols.add(df_pw.shape[1] - 1)
        if df_pw.shape[1] >= 3:
            sat_cols.add(df_pw.shape[1] - 2)

    prev_we_rest = {
        c_doc: {"sat_24h": False, "sun_any": False, "sun_night": False}
        for c_doc in doc_constraints.keys()
    }

    for r in range(df_pw.shape[0]):
        row_label = " ".join(
            strip_accents(str(df_pw.iloc[r, c_lbl])).lower()
            for c_lbl in range(min(3, df_pw.shape[1]))
            if pd.notnull(df_pw.iloc[r, c_lbl])
        )
        is_night_row = any(k in row_label for k in ["nuit", "garde", "18h", "18:00", "demi"])
        is_day_row = any(k in row_label for k in ["jour", "08h", "08:00", "08:30", "matin"])

        for c in range(df_pw.shape[1]):
            cell = df_pw.iloc[r, c]
            if pd.isna(cell):
                continue
            raw_cell = str(cell).strip()
            if not raw_cell or raw_cell.lower() in {"nan", "none", "nat"}:
                continue

            # Découper les cellules contenant plusieurs médecins (retours ligne, virgules, slashs)
            entries = [e.strip() for e in re.split(r"[\n\r,;/]+", raw_cell) if e.strip()]
            for entry in entries:
                entry_clean = clean_str(entry)
                if len(entry_clean) < 3:
                    continue
                for c_doc, sigs in doc_signatures.items():
                    matched_sig = any(
                        entry_clean == sig or entry_clean.startswith(sig)
                        for sig in sigs
                    )
                    if matched_sig:
                        if c in sun_cols:
                            prev_we_rest[c_doc]["sun_any"] = True
                            if is_night_row or not is_day_row:
                                prev_we_rest[c_doc]["sun_night"] = True
                        elif c in sat_cols and is_night_row:
                            prev_we_rest[c_doc]["sat_24h"] = True

    return prev_we_rest


def parse_contraintes_excel(file_obj):
    """
    Lit le fichier Excel des Contraintes et extrait strictement la liste
    des médecins réels ainsi que toutes leurs contraintes.
    """
    df_c = pd.read_excel(file_obj)
    df_c.columns = [str(c).strip() for c in df_c.columns]
    cols = list(df_c.columns)

    name_col = cols[0]
    for c in cols:
        c_norm = strip_accents(c).lower()
        if c_norm in {"noms", "nom", "medecin", "medecins", "praticien", "praticiens"}:
            name_col = c
            break

    col_g_name = None
    col_h_name = None
    col_etage_name = None
    col_garde_name = "Garde"
    col_demi_garde_name = "Demi garde"
    col_24h_name = "24H"
    known_other_headers = {
        "noms", "nom", "temps urgences", "temps smur", "garde", "gardes",
        "demi garde", "demi-garde", "demi gardes", "demi-gardes",
        "24h", "preference nuit", "mix urg/smur", "jeudi off",
        "mercredi off", "mra", "sauv", "temps additionnel"
    }

    for c in cols:
        c_norm = strip_accents(c).lower()
        if c_norm in {"garde", "gardes"}:
            col_garde_name = c
        elif c_norm in {"demi garde", "demi-garde", "demi gardes", "demi-gardes"}:
            col_demi_garde_name = c
        elif c_norm == "24h":
            col_24h_name = c
        elif any(k in c_norm for k in [
            "nuit uniquement", "uniquement nuit", "garde uniquement",
            "uniquement garde", "que des nuits", "que des gardes",
            "seulement nuit", "seulement garde"
        ]):
            col_g_name = c
        elif any(k in c_norm for k in ["enchainement", "enchainer"]):
            col_h_name = c
        elif any(k in c_norm for k in ["medecin etage", "medecins etages", "etage"]):
            col_etage_name = c

    if col_g_name is None and len(cols) > 6:
        if strip_accents(cols[6]).lower() not in known_other_headers:
            col_g_name = cols[6]
    if col_h_name is None and len(cols) > 7:
        if strip_accents(cols[7]).lower() not in known_other_headers:
            col_h_name = cols[7]

    doc_dict = {}
    for _, row in df_c.iterrows():
        raw_name = str(row.get(name_col, "")).strip()
        if not is_valid_doctor_name(raw_name):
            continue

        is_etage = is_yes_value(row.get(col_etage_name, "Non")) if col_etage_name else False

        raw_urg = row.get("Temps Urgences", np.nan)
        raw_smur = row.get("Temps SMUR", np.nan)
        if not is_etage and pd.isna(raw_urg) and pd.isna(raw_smur):
            continue

        try:
            ratio_urg = float(raw_urg) if pd.notnull(raw_urg) else 0.0
        except (ValueError, TypeError):
            ratio_urg = 0.0
        try:
            ratio_smur = float(raw_smur) if pd.notnull(raw_smur) else 0.0
        except (ValueError, TypeError):
            ratio_smur = 0.0

        if not is_etage and (ratio_urg + ratio_smur) <= 0.0:
            continue

        t_add_raw = strip_accents(str(row.get("Temps additionnel", "Non"))).strip().lower()
        if "++" in t_add_raw or "oui++" in t_add_raw:
            t_add_pct = 0.225
            t_add_label = "Oui++ (+22.5%)"
        elif "+" in t_add_raw or "oui+" in t_add_raw:
            t_add_pct = 0.175
            t_add_label = "Oui+ (+17.5%)"
        elif "oui" in t_add_raw:
            t_add_pct = 0.125
            t_add_label = "Oui (+12.5%)"
        else:
            t_add_pct = 0.00
            t_add_label = "Non (+0%)"

        name_cleaned = clean_str(raw_name)
        nuit_uniquement = (
            (is_yes_value(row.get(col_g_name, "Non")) if col_g_name else False)
            or name_cleaned == "said"
        )
        no_enchainement_demi = is_yes_value(row.get(col_h_name, "Non")) if col_h_name else False

        is_kenfack = (
            "kenfack" in name_cleaned
            or "elodieflore" in name_cleaned
            or name_cleaned == "elodie"
        )

        accepte_garde = is_yes_value(row.get(col_garde_name, "Oui"))
        accepte_demi_garde = is_yes_value(row.get(col_demi_garde_name, "Oui"))
        accepte_24h = is_yes_value(row.get(col_24h_name, "Non"))

        if nuit_uniquement and not is_etage:
            accepte_garde = True
            accepte_demi_garde = False
            accepte_24h = False

        if is_kenfack:
            accepte_garde = True
            accepte_24h = True
            accepte_demi_garde = False
            nuit_uniquement = False

        if is_etage:
            ratio_urg = 0.0
            ratio_smur = 0.0
            accepte_garde = False
            accepte_demi_garde = True
            accepte_24h = False
            nuit_uniquement = False

        pref_nuit_raw = strip_accents(str(row.get("Préférence nuit", "Mix"))).strip().lower()
        if "demi" in pref_nuit_raw:
            pref_nuit = "demi_garde"
        elif "garde" in pref_nuit_raw:
            pref_nuit = "garde"
        else:
            pref_nuit = "mix"

        doc_dict[raw_name] = {
            "raw_name": raw_name,
            "ratio_urg": ratio_urg,
            "ratio_smur": ratio_smur,
            "is_etage": is_etage,
            "is_kenfack": is_kenfack,
            "accepte_garde": accepte_garde,
            "accepte_demi_garde": accepte_demi_garde,
            "accepte_24h": accepte_24h,
            "nuit_uniquement": nuit_uniquement,
            "no_enchainement_demi_garde": no_enchainement_demi,
            "pref_nuit": pref_nuit,
            "mix_urg_smur": is_yes_value(row.get("Mix urg/smur", "Non")) and not is_etage,
            "jeudi_off": is_yes_value(row.get("Jeudi off", "Non")),
            "mercredi_off": is_yes_value(row.get("Mercredi off", "Non")),
            "capa_mra": is_yes_value(row.get("MRA", "Oui")) and not is_etage,
            "capa_sauv": is_yes_value(row.get("SAUV", "Oui")) and not is_etage,
            "temps_add_pct": t_add_pct,
            "temps_add_label": t_add_label,
            "desid_matched_name": None,
        }
    return doc_dict


def parse_cell_desiderata(cell_val, is_etage=False):
    """
    Analyse détaillée d'une cellule du planning des désidératas (mono ou multi-ligne).
    Pour les médecins d'étages (is_etage=True), une case vide signifie indisponible ('BLOCKED').
    """
    empty_status = "BLOCKED" if is_etage else "FREE"
    if cell_val is None or (isinstance(cell_val, float) and np.isnan(cell_val)):
        return empty_status, empty_status, False, False, False, False
    cell_str = str(cell_val).strip()
    if not cell_str or cell_str.lower() in {"nan", "none", "nat"}:
        return empty_status, empty_status, False, False, False, False

    cell_upper = strip_accents(cell_str).upper()
    lines = [l.strip() for l in cell_upper.split("\n") if l.strip()]

    am_status, nu_status = empty_status, empty_status
    is_vac, is_rtt, is_form, is_mal = False, False, False, False

    blocked_keywords = [
        "INDISPONIBLE", "INDISPO", "FORMATION", "CONGE", "RTT", "R.T.T",
        "ARRET", "MALADIE", "ABSENT", "ASTREINTE", "EXTRA-CLINIQUE",
        "EXTRA CLINIQUE", "GARDE OU ASTREINTE", "NE PEUT PAS", "IMPOSSIBLE",
        "RECUP", "REPOS COMPENSATEUR"
    ]
    avoid_keywords = ["PREFERE NE PAS", "NE PAS TRAVAILLER", "EVITER"]
    demi_keywords = ["DEMI GARDE", "DEMI-GARDE", "DEMI_GARDE", "1/4"]

    for line in lines:
        has_am_tag = "(AM)" in line or " MATIN" in line or line.startswith("MATIN")
        has_nu_tag = "(NU)" in line or "(1/4" in line or " NUIT" in line or line.startswith("NUIT")

        if has_am_tag and not has_nu_tag:
            target_am, target_nu = True, False
        elif has_nu_tag and not has_am_tag:
            target_am, target_nu = False, True
        elif has_am_tag and has_nu_tag:
            target_am, target_nu = True, True
        else:
            if "FORMATION" in line or "RTT" in line or "R.T.T" in line:
                target_am, target_nu = True, True
            elif "ABSENT" in line and "NU" in line:
                target_am, target_nu = False, True
            elif any(x in line for x in demi_keywords):
                target_am, target_nu = False, True
            elif any(x in line for x in blocked_keywords):
                target_am, target_nu = True, True
            else:
                target_am, target_nu = True, True

        is_blocked = any(x in line for x in blocked_keywords)
        is_avoid = any(x in line for x in avoid_keywords)
        has_demi_kw = any(x in line for x in demi_keywords)
        is_want = (
            any(x in line for x in ["VEUT", "SOUHAIT FORT", "SOUHAITE TRAVAILLER"])
            and not is_avoid
            and not is_blocked
        )
        is_prefer = (
            any(x in line for x in ["PREFERE TRAVAILLER", "PREFERE"])
            and not is_avoid
            and not is_want
            and not is_blocked
        )
        is_demi_only = (
            (has_demi_kw or "INDIFFERENT" in line)
            and not is_want
            and not is_blocked
            and not is_avoid
        )
        is_dispo_etage = (
            is_etage
            and not is_blocked
            and not is_avoid
            and not is_want
            and not is_prefer
            and (has_am_tag or has_nu_tag or "DISPO" in line or "TRAVAILLER" in line)
        )

        if "CONGE" in line:
            is_vac = True
        if "RTT" in line or "R.T.T" in line:
            is_rtt = True
        if "FORMATION" in line:
            is_form = True
        if "ARRET" in line or "MALADIE" in line:
            is_mal = True

        if target_am:
            if is_blocked:
                am_status = "BLOCKED"
            elif am_status != "BLOCKED" or is_etage:
                if is_want:
                    am_status = "WANT"
                elif (is_prefer or is_dispo_etage) and am_status not in {"WANT"}:
                    am_status = "PREFER"
                elif is_avoid and am_status not in {"WANT", "PREFER"}:
                    am_status = "AVOID"

        if target_nu:
            if is_blocked:
                nu_status = "BLOCKED"
            elif nu_status != "BLOCKED" or is_etage:
                if is_want and has_demi_kw:
                    nu_status = "WANT_DEMI"
                elif is_want and nu_status != "WANT_DEMI":
                    nu_status = "WANT"
                elif (is_prefer or is_dispo_etage) and nu_status not in {"WANT_DEMI", "WANT"}:
                    nu_status = "PREFER"
                elif is_avoid and nu_status not in {"WANT_DEMI", "WANT", "PREFER"}:
                    nu_status = "AVOID"
                elif is_demi_only and nu_status not in {"WANT_DEMI", "WANT", "PREFER", "AVOID"}:
                    nu_status = "DEMI_ONLY"

    return am_status, nu_status, is_vac, is_rtt, is_form, is_mal


def parse_desiderata_excel(file_obj, contraintes_dict):
    """
    Lit le fichier Excel des Désidératas, associe chaque médecin du fichier Contraintes,
    et retourne également le bilan de correspondance (retrouvés vs manquants).
    """
    df_d = pd.read_excel(file_obj)

    date_row_idx = 3
    best_date_count = 0
    for r_scan in range(min(10, df_d.shape[0])):
        cnt = 0
        for c_idx in range(1, df_d.shape[1]):
            val = df_d.iloc[r_scan, c_idx]
            if isinstance(val, (datetime.datetime, pd.Timestamp)):
                cnt += 1
            elif isinstance(val, str):
                try:
                    parsed = pd.to_datetime(val, dayfirst=True)
                    if pd.notnull(parsed):
                        cnt += 1
                except Exception:
                    pass
        if cnt > best_date_count:
            best_date_count = cnt
            date_row_idx = r_scan

    date_cols, dates = {}, []
    for c_idx in range(1, df_d.shape[1]):
        val = df_d.iloc[date_row_idx, c_idx]
        if isinstance(val, (datetime.datetime, pd.Timestamp)):
            d = val.date()
            dates.append(d)
            date_cols[c_idx] = d
        elif isinstance(val, str):
            try:
                parsed = pd.to_datetime(val, dayfirst=True)
                if pd.notnull(parsed):
                    d = parsed.date()
                    dates.append(d)
                    date_cols[c_idx] = d
            except Exception:
                pass
    dates = sorted(list(set(dates)))

    raw_rows_by_doc = {}
    for r_idx in range(date_row_idx + 1, df_d.shape[0]):
        doc_full = str(df_d.iloc[r_idx, 0]).strip()
        if not is_valid_doctor_name(doc_full):
            continue
        raw_rows_by_doc[doc_full] = {
            date_val: df_d.iloc[r_idx, col_idx]
            for col_idx, date_val in date_cols.items()
        }

    final_desid, final_conges, final_rtt, final_absences, final_formation = {}, {}, {}, {}, {}
    docs_to_keep = {}
    desid_docs_list = list(raw_rows_by_doc.keys())
    already_matched = set()
    matched_pairs = []
    unmatched_docs = []

    for c_doc, c_info in contraintes_dict.items():
        is_etage = c_info.get("is_etage", False)
        matched = match_doctor_name(c_doc, desid_docs_list, already_matched=already_matched)
        info_copy = dict(c_info)
        if matched is not None and matched in raw_rows_by_doc:
            already_matched.add(matched)
            info_copy["desid_matched_name"] = matched
            matched_pairs.append((c_doc, matched))

            doc_desid = {}
            c_cnt, rtt_cnt, f_cnt, abs_cnt = 0, 0, 0, 0
            for d_val in dates:
                cell_val = raw_rows_by_doc[matched].get(d_val)
                am_st, nu_st, is_vac, is_rtt, is_form, is_mal = parse_cell_desiderata(
                    cell_val, is_etage=is_etage
                )
                if d_val.weekday() < 5:
                    if is_vac:
                        c_cnt += 1
                        abs_cnt += 1
                    elif is_rtt:
                        rtt_cnt += 1
                        abs_cnt += 1
                    elif is_form:
                        f_cnt += 1
                        abs_cnt += 1
                    elif is_mal:
                        abs_cnt += 1
                doc_desid[d_val] = {
                    "AM": am_st,
                    "NU": nu_st,
                    "is_conges": is_vac,
                    "is_rtt": is_rtt,
                    "is_formation": is_form,
                    "is_maladie": is_mal,
                }
            final_desid[c_doc] = doc_desid
            final_conges[c_doc] = c_cnt
            final_rtt[c_doc] = rtt_cnt
            final_formation[c_doc] = f_cnt
            final_absences[c_doc] = abs_cnt
        else:
            unmatched_docs.append(c_doc)
            empty_st = "BLOCKED" if is_etage else "FREE"
            final_desid[c_doc] = {
                d_val: {
                    "AM": empty_st,
                    "NU": empty_st,
                    "is_conges": False,
                    "is_rtt": False,
                    "is_formation": False,
                    "is_maladie": False,
                }
                for d_val in dates
            }
            final_conges[c_doc] = 0
            final_rtt[c_doc] = 0
            final_formation[c_doc] = 0
            final_absences[c_doc] = 0

        docs_to_keep[c_doc] = info_copy

    match_report = {
        "total": len(contraintes_dict),
        "matched_count": len(matched_pairs),
        "matched_pairs": matched_pairs,
        "unmatched_docs": unmatched_docs,
    }
    return (
        dates,
        final_desid,
        final_conges,
        final_rtt,
        final_absences,
        final_formation,
        docs_to_keep,
        match_report,
    )


# ==============================================================================
# 4. ALGORITHME DE RÉSOLUTION OPTIMISÉ
# ==============================================================================
def get_weekend_index(d_obj, first_date):
    """Retourne l'indice de semaine du week-end (samedi/dimanche) ou None."""
    if d_obj.weekday() == 5:
        sat = d_obj
    elif d_obj.weekday() == 6:
        sat = d_obj - datetime.timedelta(days=1)
    else:
        return None
    return (sat - first_date).days // 7


def can_work_weekend(doc_worked_weekends_set, w_idx):
    """Interdit de travailler plus de 2 week-ends consécutifs."""
    if w_idx is None:
        return True
    if (w_idx - 1) in doc_worked_weekends_set and (w_idx - 2) in doc_worked_weekends_set:
        return False
    if (w_idx - 1) in doc_worked_weekends_set and (w_idx + 1) in doc_worked_weekends_set:
        return False
    if (w_idx + 1) in doc_worked_weekends_set and (w_idx + 2) in doc_worked_weekends_set:
        return False
    return True


def check_combination_validity(shift_a, shift_b, doc_info):
    """Vérifie qu'un couple (Jour, Nuit/Demi-garde) sur une même journée est autorisé."""
    t1 = SHIFTS_CONFIG[shift_a]["type"]
    t2 = SHIFTS_CONFIG[shift_b]["type"]
    if t1 == "JOUR" and t2 == "JOUR":
        return False
    if t1 != "JOUR" and t2 != "JOUR":
        return False

    d_shift = shift_a if t1 == "JOUR" else shift_b
    n_shift = shift_b if t2 != "JOUR" else shift_a

    if doc_info.get("nuit_uniquement", False):
        return False

    if doc_info.get("is_etage", False):
        return (
            d_shift in ETAGE_DAY_SHIFTS
            and n_shift in DEMI_GARDES_ORDER
            and not doc_info.get("no_enchainement_demi_garde", False)
        )

    for day_list, night_list, req in VALID_COMBOS:
        if d_shift in day_list and n_shift in night_list:
            if req == "accepte_24h":
                if not doc_info.get("accepte_24h", False):
                    return False
                if doc_info.get("mix_urg_smur", False):
                    return False
                return True
            elif req == "accepte_demi_garde":
                if not doc_info.get("accepte_demi_garde", False):
                    return False
                if doc_info.get("no_enchainement_demi_garde", False):
                    return False
                return True
            elif req == "mix_urg_smur_and_24h":
                if not doc_info.get("mix_urg_smur", False) or not doc_info.get("accepte_24h", False):
                    return False
                return True
    return False


def solve_planning(
    doc_constraints,
    desiderata,
    dates,
    conges_days,
    absence_days,
    unavail_months,
    prev_we_data=None,
    allow_overtime=False,
):
    """
    Solveur complet :
    - Passe Étage : 100% des souhaits des médecins d'étages (Urgences Journée hors SAUV/MRA et Demi-gardes)
      avec repos obligatoire le lendemain d'une demi-garde.
    - Passe 0A : Souhaits 'Veut demi-garde (NU)' (avec journée si souhaitée/possible, sinon demi-garde seule)
      et interdiction formelle de placer une garde complète sur 'Veut demi-garde (NU)'.
    - Règle Kenfack Elodie Flore : Ven-Dim uniquement, 24H exclusivement, équilibrage Ven/Sam/Dim.
    - Équilibre des dimanches pour tous les praticiens et garantie du contrat minimum.
    """
    if prev_we_data is None:
        prev_we_data = {}

    schedule = {d: {s: None for s in SHIFTS_CONFIG.keys()} for d in dates}
    first_date = dates[0]

    doc_hours = {doc: 0.0 for doc in doc_constraints.keys()}
    doc_urg_hours = {doc: 0.0 for doc in doc_constraints.keys()}
    doc_smur_hours = {doc: 0.0 for doc in doc_constraints.keys()}
    doc_gardes = {doc: 0 for doc in doc_constraints.keys()}
    doc_smur_nights = {doc: 0 for doc in doc_constraints.keys()}
    doc_urg_nights = {doc: 0 for doc in doc_constraints.keys()}
    doc_demi_gardes = {doc: 0 for doc in doc_constraints.keys()}
    doc_wd_counts = {doc: {wd: 0 for wd in range(7)} for doc in doc_constraints.keys()}
    doc_worked_weekends = {doc: set() for doc in doc_constraints.keys()}
    doc_saturdays = {doc: 0 for doc in doc_constraints.keys()}
    doc_sundays = {doc: 0 for doc in doc_constraints.keys()}
    doc_days_worked = {doc: {w: set() for w in range(5)} for doc in doc_constraints.keys()}
    doc_shift_counts = {doc: {s: 0 for s in SHIFTS_CONFIG.keys()} for doc in doc_constraints.keys()}

    solo_night_wanted = set()

    doc_targets_base = {}
    doc_schedule_targets = {}
    doc_smur_targets = {}
    doc_urg_targets = {}
    doc_eff_fte = {}
    doc_scarcity_bonus = {}

    for doc, info in doc_constraints.items():
        if info.get("is_etage", False):
            doc_targets_base[doc] = 0.0
            doc_schedule_targets[doc] = 0.0
            doc_smur_targets[doc] = 0.0
            doc_urg_targets[doc] = 0.0
            doc_eff_fte[doc] = 0.0
            doc_scarcity_bonus[doc] = 0.0
            continue

        doc_unavail = unavail_months.get(doc, set())
        doc_available_dates = [d for d in dates if d.month not in doc_unavail]
        total_weekdays_doc = sum(1 for d in doc_available_dates if d.weekday() < 5)
        num_weeks_doc = len(doc_available_dates) / 7.0

        fte = info["ratio_urg"] + info["ratio_smur"]
        base_weekly = 39.0 * fte
        total_target_brut = base_weekly * num_weeks_doc
        c_abs = absence_days.get(doc, 0)
        target_base = max(0.0, total_target_brut - (c_abs * 7.8 * fte))
        doc_targets_base[doc] = target_base

        avail_ratio = max(0.25, (total_weekdays_doc - c_abs) / max(1, total_weekdays_doc))
        doc_eff_fte[doc] = max(0.1, fte * avail_ratio) if fte > 0 else 0.0

        ot_pct = info["temps_add_pct"] if allow_overtime else 0.0
        sched_target = target_base * (1.0 + ot_pct)
        doc_schedule_targets[doc] = sched_target

        if fte > 0:
            smur_share = info["ratio_smur"] / fte
            urg_share = info["ratio_urg"] / fte
        else:
            smur_share, urg_share = 0.0, 0.0
        doc_smur_targets[doc] = sched_target * smur_share
        doc_urg_targets[doc] = sched_target * urg_share

        open_days = 0
        for d_chk in doc_available_dates:
            if info.get("is_kenfack") and d_chk.weekday() < 4:
                continue
            if d_chk.weekday() == 3 and info["jeudi_off"]:
                continue
            if d_chk.weekday() == 2 and info["mercredi_off"]:
                continue
            des_chk = desiderata[doc].get(d_chk, {"AM": "FREE", "NU": "FREE"})
            can_am = (not info.get("nuit_uniquement", False)) and des_chk["AM"] != "BLOCKED"
            can_nu = (info["accepte_garde"] or info["accepte_demi_garde"]) and des_chk["NU"] != "BLOCKED"
            if can_am or can_nu:
                open_days += 1

        needed_shifts_est = max(1.0, sched_target / 11.0)
        scarcity_ratio = open_days / needed_shifts_est
        if sched_target > 0 and scarcity_ratio < 2.4:
            doc_scarcity_bonus[doc] = min(45000.0, (2.4 - scarcity_ratio) * 22000.0)
        else:
            doc_scarcity_bonus[doc] = 0.0

    regular_docs_list = [d for d, inf in doc_constraints.items() if not inf.get("is_etage", False)]
    etage_docs_list = [d for d, inf in doc_constraints.items() if inf.get("is_etage", False)]

    total_we_indices = {get_weekend_index(d, first_date) for d in dates if d.weekday() >= 5}
    num_total_weekends = max(1, len(total_we_indices))
    total_eff_fte = max(0.1, sum(doc_eff_fte.values()))
    max_eff_fte = max(0.1, max(doc_eff_fte.values()) if doc_eff_fte else 1.0)

    nuit_only_gardes_est = sum(
        doc_schedule_targets[d] / 14.25
        for d in regular_docs_list
        if doc_constraints[d].get("nuit_uniquement", False)
    )
    eligible_garde_eff_fte = max(
        0.5,
        sum(
            doc_eff_fte[d]
            for d in regular_docs_list
            if doc_constraints[d]["accepte_garde"] and not doc_constraints[d].get("nuit_uniquement", False)
        ),
    )
    fair_gardes_per_fte = max(1.0, (len(dates) * 4 - nuit_only_gardes_est) / eligible_garde_eff_fte)

    doc_max_gardes = {}
    for d_name, inf in doc_constraints.items():
        if inf.get("is_etage", False):
            doc_max_gardes[d_name] = 0
        elif inf.get("nuit_uniquement", False) or not inf["accepte_garde"] or inf.get("is_kenfack", False):
            doc_max_gardes[d_name] = 999
        elif inf["ratio_smur"] > 0 or inf["mix_urg_smur"]:
            doc_max_gardes[d_name] = max(1, math.ceil(fair_gardes_per_fte * doc_eff_fte[d_name] * 1.15))
        else:
            doc_max_gardes[d_name] = max(1, math.ceil(fair_gardes_per_fte * doc_eff_fte[d_name] * 0.96))

    def weekend_equity_score(doc, date_obj):
        """Équilibre les week-ends et veille à attribuer des dimanches à tous les praticiens."""
        info = doc_constraints[doc]
        if info.get("is_kenfack", False):
            if date_obj.weekday() in {4, 5, 6}:
                fri_c = doc_wd_counts[doc][4]
                sat_c = doc_saturdays[doc]
                sun_c = doc_sundays[doc]
                counts_map = {4: fri_c, 5: sat_c, 6: sun_c}
                min_c = min(counts_map.values())
                diff_day = counts_map[date_obj.weekday()] - min_c
                # Bonus fort sur le jour (Ven / Sam / Dim) le moins travaillé par Elodie
                return 35000.0 - (diff_day * 45000.0)
            return -999999.0

        if date_obj.weekday() < 5:
            return 0.0

        w_idx = get_weekend_index(date_obj, first_date)
        eff_fte = doc_eff_fte.get(doc, 0.0)
        if eff_fte <= 0:
            return -80000.0

        score_we = 0.0
        if w_idx not in doc_worked_weekends[doc]:
            total_we_assigned = sum(
                len(v) for k, v in doc_worked_weekends.items() if not doc_constraints[k].get("is_kenfack")
            )
            avg_we_per_fte = total_we_assigned / total_eff_fte
            current_rate = len(doc_worked_weekends[doc]) / eff_fte
            projected_rate = (len(doc_worked_weekends[doc]) + 1) / eff_fte
            if projected_rate > avg_we_per_fte:
                diff = projected_rate - avg_we_per_fte
                score_we -= diff * 50000.0 + (diff ** 2) * 12000.0
            else:
                diff = avg_we_per_fte - current_rate
                score_we += diff * 42000.0

        # Équilibre Samedi vs Dimanche (les dimanches sont payés double)
        sats = doc_saturdays[doc]
        suns = doc_sundays[doc]
        if date_obj.weekday() == 6:  # Dimanche
            if suns == 0 and sats > 0:
                score_we += 48000.0 + (sats * 12000.0)
            elif suns < sats:
                score_we += (sats - suns) * 28000.0
            else:
                score_we -= (suns - sats) * 14000.0
        elif date_obj.weekday() == 5:  # Samedi
            if sats > suns:
                score_we -= (sats - suns) * 26000.0 + (15000.0 if suns == 0 else 0.0)
            elif suns > sats:
                score_we += (suns - sats) * 18000.0

        return score_we

    def hour_priority_score(doc, base_weight=65000.0):
        target = doc_schedule_targets[doc]
        tot_deficit = target - doc_hours[doc]
        if tot_deficit > 0:
            pct_missing = tot_deficit / max(10.0, target)
            return base_weight + (pct_missing * 35000.0) + (tot_deficit * 60.0) + doc_scarcity_bonus.get(doc, 0.0)
        return -25000.0 + (tot_deficit * 200.0)

    def is_24h_on_date(doc_name, d_obj):
        if d_obj not in schedule:
            return False
        has_day = any(SHIFTS_CONFIG[s]["type"] == "JOUR" for s, a in schedule[d_obj].items() if a == doc_name)
        has_night = any(SHIFTS_CONFIG[s]["type"] in {"NUIT", "GARDE"} for s, a in schedule[d_obj].items() if a == doc_name)
        return has_day and has_night

    def worked_night_on_date(doc_name, d_obj):
        if d_obj not in schedule:
            return False
        return any(
            SHIFTS_CONFIG[s]["type"] in {"NUIT", "GARDE", "DEMI_GARDE"}
            for s, a in schedule[d_obj].items()
            if a == doc_name
        )

    def worked_full_night_on_date(doc_name, d_obj):
        if d_obj not in schedule:
            return False
        return any(
            SHIFTS_CONFIG[s]["type"] in {"NUIT", "GARDE"}
            for s, a in schedule[d_obj].items()
            if a == doc_name
        )

    def worked_any_on_date(doc_name, d_obj):
        if d_obj not in schedule:
            return False
        return any(a == doc_name for a in schedule[d_obj].values())

    def can_assign_base(
        doc,
        date_obj,
        shift_keys_list,
        enforce_90_night_cap=False,
        enforce_smur_night_ratio=True,
        is_explicit_want=False,
        avoid_day_before_night=False,
        relax_soft_limits=False,
        enforce_night_balance=False,
    ):
        if date_obj.month in unavail_months.get(doc, set()):
            return False

        day_idx = dates.index(date_obj)
        info = doc_constraints[doc]
        des = desiderata[doc].get(date_obj, {"AM": "FREE", "NU": "FREE"})

        # Règle spécifique Kenfack Elodie Flore : uniquement Ven/Sam/Dim et exclusivement en 24H
        if info.get("is_kenfack", False):
            if date_obj.weekday() < 4:
                return False
            existing_on_d = [s for s, a in schedule[date_obj].items() if a == doc]
            combined_on_d = existing_on_d + shift_keys_list
            if len(combined_on_d) != 2:
                return False
            types_set = {SHIFTS_CONFIG[s]["type"] for s in combined_on_d}
            if not ("JOUR" in types_set and any(t in {"NUIT", "GARDE"} for t in types_set)):
                return False

        # Règle stricte: 'Veut travailler (demi-garde seulement) (NU)' -> jamais de garde complète ce jour-là
        if des.get("NU") == "WANT_DEMI":
            if any(SHIFTS_CONFIG[sk]["type"] in {"NUIT", "GARDE"} for sk in shift_keys_list):
                return False

        # Médecin d'étage : uniquement Urgences Journée (hors SAUV/MRA) et/ou Demi-garde
        if info.get("is_etage", False):
            for sk in shift_keys_list:
                if sk not in ETAGE_DAY_SHIFTS and sk not in DEMI_GARDES_ORDER:
                    return False
                if SHIFTS_CONFIG[sk]["type"] == "JOUR" and des["AM"] not in {"WANT", "PREFER"}:
                    return False
                if SHIFTS_CONFIG[sk]["type"] == "DEMI_GARDE" and des["NU"] not in {"WANT_DEMI", "WANT", "PREFER"}:
                    return False

            # Repos obligatoire le lendemain d'une demi-garde (et vice-versa)
            if day_idx == 0 and prev_we_data.get(doc, {}).get("sun_night"):
                return False
            if day_idx > 0 and worked_night_on_date(doc, dates[day_idx - 1]):
                return False
            has_demi_new = any(SHIFTS_CONFIG[sk]["type"] == "DEMI_GARDE" for sk in shift_keys_list)
            if has_demi_new and day_idx + 1 < len(dates) and worked_any_on_date(doc, dates[day_idx + 1]):
                return False

            existing_shifts = [s for s, a in schedule[date_obj].items() if a == doc]
            all_shifts_for_day = existing_shifts + shift_keys_list
            if len(all_shifts_for_day) > 2:
                return False
            if len(all_shifts_for_day) == 2:
                if not check_combination_validity(all_shifts_for_day[0], all_shifts_for_day[1], info):
                    return False
            for sk in shift_keys_list:
                if schedule[date_obj][sk] is not None:
                    return False
            return True

        total_heures = sum(SHIFTS_CONFIG[sk]["heures"] for sk in shift_keys_list)
        fte = info["ratio_urg"] + info["ratio_smur"]
        sched_target = doc_schedule_targets[doc]

        # 1. Plafond horaire strict
        if fte <= 0.0:
            return False
        if fte < 0.35:
            if not is_explicit_want and doc_hours[doc] >= sched_target - 0.5:
                return False
            extra_small = 10.0 if is_explicit_want else (
                9.5 if (relax_soft_limits and doc_hours[doc] < sched_target - 4.0) else 6.5
            )
            max_cap = sched_target + extra_small
            if doc_hours[doc] + total_heures > max_cap:
                return False
        else:
            if is_explicit_want:
                max_cap = sched_target + max(24.0, 28.0 * fte)
            elif relax_soft_limits:
                max_cap = sched_target + max(14.5, 16.0 * fte)
            else:
                max_cap = sched_target + max(14.5, 18.0 * fte)
            if doc_hours[doc] + total_heures > max_cap:
                return False

        # 2. Équilibre Jour / Nuit et plafond équitable de gardes par ETP
        added_full_gardes = sum(
            1 for sk in shift_keys_list if SHIFTS_CONFIG[sk]["type"] in {"NUIT", "GARDE"}
        )
        if added_full_gardes > 0 and not is_explicit_want and not info.get("nuit_uniquement", False):
            if not (relax_soft_limits and doc_hours[doc] < sched_target - 14.0):
                if doc_gardes[doc] + added_full_gardes > doc_max_gardes[doc]:
                    return False

        if enforce_night_balance and not is_explicit_want and not info.get("nuit_uniquement", False):
            added_night_h = sum(
                SHIFTS_CONFIG[sk]["heures"]
                for sk in shift_keys_list
                if SHIFTS_CONFIG[sk]["type"] in {"NUIT", "GARDE"}
            )
            if added_night_h > 0:
                current_night_h = (doc_smur_nights[doc] * 14.0) + (doc_urg_nights[doc] * 14.5)
                if current_night_h + added_night_h > sched_target * 0.56 + 4.0:
                    return False
            if "SMUR_NUIT" in shift_keys_list:
                current_smur_night_h = doc_smur_nights[doc] * 14.0
                if current_smur_night_h + 14.0 > doc_smur_targets[doc] * 0.72 + 7.0:
                    return False

        # 3. Règles Week-end (non appliquées à Kenfack Elodie Flore)
        if not info.get("is_kenfack", False):
            w_idx = get_weekend_index(date_obj, first_date)
            if w_idx is not None and w_idx not in doc_worked_weekends[doc]:
                if not (relax_soft_limits and doc_hours[doc] < sched_target - 14.0):
                    if not can_work_weekend(doc_worked_weekends[doc], w_idx):
                        return False
                if not is_explicit_want and not relax_soft_limits:
                    eff_fte = doc_eff_fte.get(doc, 0.0)
                    max_we_allowed = max(1, math.ceil(num_total_weekends * 0.65 * (eff_fte / max_eff_fte)))
                    if len(doc_worked_weekends[doc]) >= max_we_allowed:
                        return False
            if date_obj.weekday() == 5:
                sun = date_obj + datetime.timedelta(days=1)
                if sun in schedule and worked_any_on_date(doc, sun):
                    return False
            elif date_obj.weekday() == 6:
                sat = date_obj - datetime.timedelta(days=1)
                if sat in schedule and worked_any_on_date(doc, sat):
                    return False

        # 4. Repos de sécurité après garde / nuit / 24h (WE précédent + D-1, D-2)
        has_nuit_in_new = any(SHIFTS_CONFIG[sk]["type"] in {"NUIT", "GARDE", "DEMI_GARDE"} for sk in shift_keys_list)
        has_full_garde_in_new = any(SHIFTS_CONFIG[sk]["type"] in {"NUIT", "GARDE"} for sk in shift_keys_list)
        if day_idx == 0 and prev_we_data.get(doc, {}).get("sun_night"):
            return False
        if day_idx == 1 and prev_we_data.get(doc, {}).get("sun_night") and has_nuit_in_new:
            return False

        if day_idx > 0:
            prev_1 = dates[day_idx - 1]
            if is_24h_on_date(doc, prev_1) or worked_night_on_date(doc, prev_1):
                return False
        if day_idx > 1:
            prev_2 = dates[day_idx - 2]
            if is_24h_on_date(doc, prev_2):
                return False
            if has_nuit_in_new:
                if is_explicit_want and not has_full_garde_in_new:
                    if worked_full_night_on_date(doc, prev_2):
                        return False
                elif worked_night_on_date(doc, prev_2):
                    return False

        # 5. Repos de sécurité futurs (D+1, D+2) et veille d'une matinée bloquée
        existing_day = any(SHIFTS_CONFIG[s]["type"] == "JOUR" for s, a in schedule[date_obj].items() if a == doc)
        will_be_24h = existing_day
        has_full_night_in_new = False
        has_day_in_new = False
        for sk in shift_keys_list:
            if SHIFTS_CONFIG[sk]["type"] == "JOUR":
                will_be_24h = True
                has_day_in_new = True
            elif SHIFTS_CONFIG[sk]["type"] in {"NUIT", "GARDE"}:
                has_full_night_in_new = True
        will_be_24h = will_be_24h and has_full_night_in_new

        if avoid_day_before_night and has_day_in_new and not has_nuit_in_new:
            if day_idx + 1 < len(dates) and worked_night_on_date(doc, dates[day_idx + 1]):
                if des["AM"] not in {"WANT", "PREFER"}:
                    return False

        if has_full_night_in_new and day_idx + 1 < len(dates):
            next_des = desiderata[doc].get(dates[day_idx + 1], {"AM": "FREE", "NU": "FREE"})
            if next_des["AM"] == "BLOCKED" and des["NU"] not in {"WANT", "WANT_DEMI"}:
                return False

        if will_be_24h:
            if day_idx + 1 < len(dates) and worked_any_on_date(doc, dates[day_idx + 1]):
                return False
            if day_idx + 2 < len(dates) and worked_any_on_date(doc, dates[day_idx + 2]):
                return False
        elif has_nuit_in_new:
            if day_idx + 1 < len(dates) and worked_any_on_date(doc, dates[day_idx + 1]):
                return False
            if day_idx + 2 < len(dates):
                if is_explicit_want and not has_full_garde_in_new:
                    if worked_full_night_on_date(doc, dates[day_idx + 2]):
                        return False
                elif worked_night_on_date(doc, dates[day_idx + 2]):
                    return False

        # 6. Vérification poste par poste
        for shift_key in shift_keys_list:
            cfg = SHIFTS_CONFIG[shift_key]
            if info.get("nuit_uniquement", False) and cfg["type"] not in {"NUIT", "GARDE"}:
                return False

            if cfg["type"] == "JOUR" and des["AM"] == "BLOCKED":
                return False
            if cfg["type"] in {"NUIT", "GARDE", "DEMI_GARDE"} and des["NU"] == "BLOCKED":
                return False

            if cfg["type"] == "JOUR" and (doc, date_obj) in solo_night_wanted:
                return False

            if cfg["service"] == "SMUR":
                if info["ratio_smur"] == 0 and not info["mix_urg_smur"]:
                    return False
                smur_margin = (
                    max(60.0, 120.0 * fte)
                    if (relax_soft_limits or is_explicit_want)
                    else max(8.0, 22.0 * fte)
                )
                if doc_smur_hours[doc] + cfg["heures"] > doc_smur_targets[doc] + smur_margin:
                    return False
            else:
                if info["ratio_urg"] == 0:
                    return False
                urg_margin = (
                    max(60.0, 120.0 * fte)
                    if (relax_soft_limits or is_explicit_want)
                    else max(6.0, 14.5 * fte)
                )
                if info["ratio_smur"] > 0 and doc_urg_hours[doc] + cfg["heures"] > doc_urg_targets[doc] + urg_margin:
                    return False

            if cfg["type"] == "GARDE" and not info["accepte_garde"]:
                return False
            if cfg["type"] == "DEMI_GARDE" and not info["accepte_demi_garde"]:
                return False
            if cfg.get("req_mra", False) and not info["capa_mra"]:
                return False
            if cfg.get("req_sauv", False) and not info["capa_sauv"]:
                return False

            if enforce_smur_night_ratio and shift_key in {"GARDE_1", "GARDE_2", "GARDE_3_SAUV"}:
                if info["ratio_smur"] > 0:
                    if doc_urg_nights[doc] + 1 > doc_smur_nights[doc] + 1:
                        return False

            if date_obj.weekday() == 3 and info["jeudi_off"]:
                return False
            if date_obj.weekday() == 2 and info["mercredi_off"]:
                return False
            if cfg["type"] in {"GARDE", "NUIT"}:
                if date_obj.weekday() == 2 and info["jeudi_off"]:
                    return False
                if date_obj.weekday() == 1 and info["mercredi_off"]:
                    return False

            if schedule[date_obj][shift_key] is not None:
                return False

        # 7. Vérification de la combinaison sur la journée
        existing_shifts = [s for s, a in schedule[date_obj].items() if a == doc]
        all_shifts_for_day = existing_shifts + shift_keys_list
        if len(all_shifts_for_day) > 1:
            if len(all_shifts_for_day) > 2:
                return False
            if not check_combination_validity(all_shifts_for_day[0], all_shifts_for_day[1], info):
                return False

        if len(shift_keys_list) == 1 and len(existing_shifts) == 0:
            sk = shift_keys_list[0]
            if SHIFTS_CONFIG[sk]["type"] in {"GARDE", "NUIT"} and not info["accepte_garde"]:
                return False

        return True

    def assign_shifts(doc, date_obj, shift_keys_list):
        was_working_date = worked_any_on_date(doc, date_obj)
        for shift_key in shift_keys_list:
            schedule[date_obj][shift_key] = doc
            cfg = SHIFTS_CONFIG[shift_key]
            doc_hours[doc] += cfg["heures"]
            doc_shift_counts[doc][shift_key] += 1
            if cfg["service"] == "URGENCES":
                doc_urg_hours[doc] += cfg["heures"]
            else:
                doc_smur_hours[doc] += cfg["heures"]
            if shift_key == "SMUR_NUIT":
                doc_gardes[doc] += 1
                doc_smur_nights[doc] += 1
            elif cfg["type"] == "GARDE":
                doc_gardes[doc] += 1
                doc_urg_nights[doc] += 1
            elif cfg["type"] == "DEMI_GARDE":
                doc_demi_gardes[doc] += 1

        w_idx = get_weekend_index(date_obj, first_date)
        if w_idx is not None:
            doc_worked_weekends[doc].add(w_idx)

        wd = date_obj.weekday()
        if not was_working_date:
            doc_wd_counts[doc][wd] += 1
            if wd == 5:
                doc_saturdays[doc] += 1
            elif wd == 6:
                doc_sundays[doc] += 1
            elif wd < 5:
                doc_days_worked[doc][wd].add(date_obj)

    def unassign_shift(doc, date_obj, shift_key):
        schedule[date_obj][shift_key] = None
        cfg = SHIFTS_CONFIG[shift_key]
        doc_hours[doc] -= cfg["heures"]
        doc_shift_counts[doc][shift_key] = max(0, doc_shift_counts[doc][shift_key] - 1)
        if cfg["service"] == "URGENCES":
            doc_urg_hours[doc] -= cfg["heures"]
        else:
            doc_smur_hours[doc] -= cfg["heures"]
        if shift_key == "SMUR_NUIT":
            doc_gardes[doc] = max(0, doc_gardes[doc] - 1)
            doc_smur_nights[doc] = max(0, doc_smur_nights[doc] - 1)
        elif cfg["type"] == "GARDE":
            doc_gardes[doc] = max(0, doc_gardes[doc] - 1)
            doc_urg_nights[doc] = max(0, doc_urg_nights[doc] - 1)
        elif cfg["type"] == "DEMI_GARDE":
            doc_demi_gardes[doc] = max(0, doc_demi_gardes[doc] - 1)

        if not worked_any_on_date(doc, date_obj):
            wd = date_obj.weekday()
            doc_wd_counts[doc][wd] = max(0, doc_wd_counts[doc][wd] - 1)
            if wd == 5:
                doc_saturdays[doc] = max(0, doc_saturdays[doc] - 1)
            elif wd == 6:
                doc_sundays[doc] = max(0, doc_sundays[doc] - 1)
            elif wd < 5:
                doc_days_worked[doc][wd].discard(date_obj)

        w_idx = get_weekend_index(date_obj, first_date)
        if w_idx is not None:
            still_on_we = any(
                get_weekend_index(dt, first_date) == w_idx and any(a == doc for a in schedule[dt].values())
                for dt in dates
                if dt.weekday() >= 5
            )
            if not still_on_we:
                doc_worked_weekends[doc].discard(w_idx)

    def calculate_equity_score(doc, info):
        if info.get("is_etage", False):
            return 0.0
        fte = info["ratio_urg"] + info["ratio_smur"]
        if fte <= 0:
            return 999.0
        if info.get("nuit_uniquement", False) or not info["accepte_garde"] or info.get("is_kenfack", False):
            return 0.0
        return (doc_gardes[doc] + 0.5 * doc_demi_gardes[doc]) / max(0.1, fte)

    # ==========================================================================
    # PASSE ÉTAGE: MÉDECINS D'ÉTAGES (100% DES SOUHAITS EXAUCÉS EN PRIORITÉ)
    # ==========================================================================
    if etage_docs_list:
        for d in dates:
            candidates_etage = sorted(etage_docs_list, key=lambda x: (doc_hours[x], random.random()))
            for doc in candidates_etage:
                if worked_any_on_date(doc, d):
                    continue
                des = desiderata[doc].get(d, {"AM": "BLOCKED", "NU": "BLOCKED"})
                want_am = des["AM"] in {"WANT", "PREFER"}
                want_nu = des["NU"] in {"WANT_DEMI", "WANT", "PREFER"}
                if not want_am and not want_nu:
                    continue

                if want_am and want_nu:
                    etage_combos = [
                        (["CIRCUIT_LONG_1", "CIRCUIT_LONG_2"], "DEMI_TRANS"),
                        (["COURS_DE_CIRCUIT"], "DEMI_COURT"),
                    ]
                    random.shuffle(etage_combos)
                    combo_ok = False
                    for day_opts, demi_sk in etage_combos:
                        if schedule[d][demi_sk] is not None:
                            continue
                        for day_sk in sorted(day_opts, key=lambda s: doc_shift_counts[doc][s]):
                            if schedule[d][day_sk] is None and can_assign_base(
                                doc, d, [day_sk, demi_sk], is_explicit_want=True
                            ):
                                assign_shifts(doc, d, [day_sk, demi_sk])
                                combo_ok = True
                                break
                        if combo_ok:
                            break
                    if combo_ok:
                        continue

                if want_am:
                    day_choices = sorted(ETAGE_DAY_SHIFTS, key=lambda s: doc_shift_counts[doc][s])
                    for day_sk in day_choices:
                        if schedule[d][day_sk] is None and can_assign_base(
                            doc, d, [day_sk], is_explicit_want=True
                        ):
                            assign_shifts(doc, d, [day_sk])
                            break
                elif want_nu:
                    for demi_sk in DEMI_GARDES_ORDER:
                        if schedule[d][demi_sk] is None and can_assign_base(
                            doc, d, [demi_sk], is_explicit_want=True
                        ):
                            assign_shifts(doc, d, [demi_sk])
                            break

    docs_list = list(regular_docs_list)
    weekdays_dates = [d for d in dates if d.weekday() < 5]
    weekend_dates = [d for d in dates if d.weekday() >= 5]
    ordered_want_dates = weekdays_dates + weekend_dates
    doc_wants_granted = {doc: 0 for doc in docs_list}

    # ==========================================================================
    # PASSE 0A: SOUHAITS 'VEUT DEMI-GARDE (NU)' (AVEC POTENTIELLEMENT LA JOURNÉE)
    # ==========================================================================
    demi_want_dates = sorted(
        dates,
        key=lambda d_item: (
            0
            if any(
                desiderata[doc].get(d_item, {}).get("AM") in {"WANT", "PREFER"}
                and desiderata[doc].get(d_item, {}).get("NU") == "WANT_DEMI"
                for doc in docs_list
            )
            else 1,
            d_item,
        ),
    )

    for d in demi_want_dates:
        candidates = sorted(docs_list, key=lambda x: (doc_wants_granted[x], random.random()))
        for doc in candidates:
            if worked_any_on_date(doc, d):
                continue
            info = doc_constraints[doc]
            if info.get("is_kenfack"):
                continue
            des = desiderata[doc].get(d, {"AM": "FREE", "NU": "FREE"})
            is_want_demi = (
                des["NU"] == "WANT_DEMI"
                or (des["NU"] == "WANT" and info["accepte_demi_garde"] and not info["accepte_garde"])
            )
            if not is_want_demi:
                continue

            assigned_ok = False
            can_pair_day = (
                not info.get("no_enchainement_demi_garde", False)
                and des["AM"] not in {"BLOCKED", "AVOID"}
                and (
                    des["AM"] in {"WANT", "PREFER"}
                    or doc_hours[doc] + 15.5 <= doc_schedule_targets[doc] + 10.0
                )
            )

            # 1. Tenter en priorité la combinaison Journée Urgences + Demi-garde si compatible
            if can_pair_day:
                demi_combos = [
                    (["CIRCUIT_LONG_1", "CIRCUIT_LONG_2"], "DEMI_TRANS"),
                    (["MRA", "SAUV", "COURS_DE_CIRCUIT"], "DEMI_COURT"),
                ]
                random.shuffle(demi_combos)
                for day_opts, demi_sk in demi_combos:
                    if schedule[d][demi_sk] is not None:
                        continue
                    day_opts_sorted = sorted(day_opts, key=lambda s: doc_shift_counts[doc][s])
                    for day_sk in day_opts_sorted:
                        if schedule[d][day_sk] is None and can_assign_base(
                            doc, d, [day_sk, demi_sk], enforce_90_night_cap=False, is_explicit_want=True
                        ):
                            assign_shifts(doc, d, [day_sk, demi_sk])
                            doc_wants_granted[doc] += (2 if des["AM"] == "WANT" else 1)
                            assigned_ok = True
                            break
                    if assigned_ok:
                        break

            # 2. Sinon, garantir impérativement la demi-garde seule
            if not assigned_ok:
                for demi_sk in DEMI_GARDES_ORDER:
                    if schedule[d][demi_sk] is None and can_assign_base(
                        doc, d, [demi_sk], enforce_90_night_cap=False, is_explicit_want=True
                    ):
                        assign_shifts(doc, d, [demi_sk])
                        doc_wants_granted[doc] += 1
                        break

    # ==========================================================================
    # PASSE 0B: DOUBLE SOUHAIT 'VEUT (AM) + VEUT (NU)' (GARDE COMPLÈTE)
    # ==========================================================================
    for d in ordered_want_dates:
        candidates = sorted(docs_list, key=lambda x: (doc_wants_granted[x], random.random()))
        for doc in candidates:
            if worked_any_on_date(doc, d):
                continue
            info = doc_constraints[doc]
            des = desiderata[doc].get(d, {"AM": "FREE", "NU": "FREE"})
            if not (des["AM"] == "WANT" and des["NU"] == "WANT"):
                continue

            combo_placed = False
            combos_priority = [c for c in VALID_COMBOS if c[2] != "accepte_demi_garde"] + [
                c for c in VALID_COMBOS if c[2] == "accepte_demi_garde"
            ]
            for day_list, night_list, _ in combos_priority:
                day_sorted = sorted(day_list, key=lambda s: doc_shift_counts[doc][s])
                night_sorted = sorted(night_list, key=lambda s: doc_shift_counts[doc][s])
                for sd in day_sorted:
                    for sn in night_sorted:
                        if schedule[d][sd] is None and schedule[d][sn] is None:
                            if can_assign_base(
                                doc,
                                d,
                                [sd, sn],
                                enforce_90_night_cap=False,
                                enforce_smur_night_ratio=False,
                                is_explicit_want=True,
                            ):
                                assign_shifts(doc, d, [sd, sn])
                                doc_wants_granted[doc] += 2
                                combo_placed = True
                                break
                    if combo_placed:
                        break
                if combo_placed:
                    break

            if combo_placed:
                continue

            if info.get("is_kenfack"):
                continue

            placed_partial = False
            if info["accepte_garde"]:
                night_choices = (
                    ["SMUR_NUIT", "GARDE_3_SAUV", "GARDE_1", "GARDE_2"]
                    if info["ratio_smur"] > 0
                    else sorted(["GARDE_3_SAUV", "GARDE_1", "GARDE_2"], key=lambda s: doc_shift_counts[doc][s])
                )
                for sk in night_choices:
                    if schedule[d][sk] is None and can_assign_base(
                        doc, d, [sk], enforce_90_night_cap=False, enforce_smur_night_ratio=False, is_explicit_want=True
                    ):
                        assign_shifts(doc, d, [sk])
                        doc_wants_granted[doc] += 1
                        placed_partial = True
                        break

            if not placed_partial and not info.get("nuit_uniquement", False):
                day_choices = sorted(DAY_SHIFTS_ORDER, key=lambda s: doc_shift_counts[doc][s])
                if info["ratio_smur"] > 0 and doc_smur_hours[doc] < doc_smur_targets[doc]:
                    day_choices = ["SMUR_JOUR"] + [s for s in day_choices if s != "SMUR_JOUR"]
                for sk in day_choices:
                    if schedule[d][sk] is None and can_assign_base(
                        doc, d, [sk], enforce_90_night_cap=False, is_explicit_want=True
                    ):
                        assign_shifts(doc, d, [sk])
                        doc_wants_granted[doc] += 1
                        placed_partial = True
                        break

            if not placed_partial and info["accepte_demi_garde"]:
                for sk in DEMI_GARDES_ORDER:
                    if schedule[d][sk] is None and can_assign_base(
                        doc, d, [sk], enforce_90_night_cap=False, is_explicit_want=True
                    ):
                        assign_shifts(doc, d, [sk])
                        doc_wants_granted[doc] += 1
                        break

    # ==========================================================================
    # PASSE 0C: SOUHAITS 'VEUT (NU)' SEUL -> GARDE COMPLÈTE
    # ==========================================================================
    for d in ordered_want_dates:
        candidates = sorted(docs_list, key=lambda x: (doc_wants_granted[x], random.random()))
        for doc in candidates:
            if worked_any_on_date(doc, d):
                continue
            info = doc_constraints[doc]
            des = desiderata[doc].get(d, {"AM": "FREE", "NU": "FREE"})
            if not (des["NU"] == "WANT" and des["AM"] != "WANT"):
                continue

            if info.get("is_kenfack"):
                for day_list, night_list, req in VALID_COMBOS:
                    if req != "accepte_24h":
                        continue
                    placed_k = False
                    for sd in day_list:
                        for sn in night_list:
                            if schedule[d][sd] is None and schedule[d][sn] is None and can_assign_base(
                                doc, d, [sd, sn], enforce_90_night_cap=False, enforce_smur_night_ratio=False, is_explicit_want=True
                            ):
                                assign_shifts(doc, d, [sd, sn])
                                doc_wants_granted[doc] += 1
                                placed_k = True
                                break
                        if placed_k:
                            break
                    if placed_k:
                        break
                continue

            solo_night_wanted.add((doc, d))
            if info["accepte_garde"]:
                if info["ratio_smur"] > 0:
                    night_choices = ["SMUR_NUIT", "GARDE_3_SAUV", "GARDE_1", "GARDE_2"]
                else:
                    night_choices = sorted(["GARDE_3_SAUV", "GARDE_1", "GARDE_2"], key=lambda s: doc_shift_counts[doc][s])
                for sk in night_choices:
                    if schedule[d][sk] is None and can_assign_base(
                        doc, d, [sk], enforce_90_night_cap=False, enforce_smur_night_ratio=False, is_explicit_want=True
                    ):
                        assign_shifts(doc, d, [sk])
                        doc_wants_granted[doc] += 1
                        break
            elif info["accepte_demi_garde"]:
                for sk in DEMI_GARDES_ORDER:
                    if schedule[d][sk] is None and can_assign_base(
                        doc, d, [sk], enforce_90_night_cap=False, is_explicit_want=True
                    ):
                        assign_shifts(doc, d, [sk])
                        doc_wants_granted[doc] += 1
                        break

    # ==========================================================================
    # PASSE 0D: SOUHAITS 'VEUT (AM)' SEUL
    # ==========================================================================
    for d in ordered_want_dates:
        candidates = sorted(docs_list, key=lambda x: (doc_wants_granted[x], random.random()))
        for doc in candidates:
            if worked_any_on_date(doc, d):
                continue
            info = doc_constraints[doc]
            if info.get("nuit_uniquement", False):
                continue
            des = desiderata[doc].get(d, {"AM": "FREE", "NU": "FREE"})
            if not (des["AM"] == "WANT" and des["NU"] not in {"WANT", "WANT_DEMI"}):
                continue

            if info.get("is_kenfack"):
                for day_list, night_list, req in VALID_COMBOS:
                    if req != "accepte_24h":
                        continue
                    placed_k = False
                    for sd in day_list:
                        for sn in night_list:
                            if schedule[d][sd] is None and schedule[d][sn] is None and can_assign_base(
                                doc, d, [sd, sn], enforce_90_night_cap=False, enforce_smur_night_ratio=False, is_explicit_want=True
                            ):
                                assign_shifts(doc, d, [sd, sn])
                                doc_wants_granted[doc] += 1
                                placed_k = True
                                break
                        if placed_k:
                            break
                    if placed_k:
                        break
                continue

            placed_am = False
            if (
                des["NU"] in {"PREFER", "DEMI_ONLY"}
                and info["accepte_demi_garde"]
                and not info.get("no_enchainement_demi_garde", False)
            ):
                demi_combos = [
                    (["CIRCUIT_LONG_1", "CIRCUIT_LONG_2"], "DEMI_TRANS"),
                    (["MRA", "SAUV", "COURS_DE_CIRCUIT"], "DEMI_COURT"),
                ]
                random.shuffle(demi_combos)
                for day_opts, demi_sk in demi_combos:
                    if schedule[d][demi_sk] is not None:
                        continue
                    for day_sk in sorted(day_opts, key=lambda s: doc_shift_counts[doc][s]):
                        if schedule[d][day_sk] is None and can_assign_base(
                            doc, d, [day_sk, demi_sk], enforce_90_night_cap=False, is_explicit_want=True
                        ):
                            assign_shifts(doc, d, [day_sk, demi_sk])
                            doc_wants_granted[doc] += 1
                            placed_am = True
                            break
                    if placed_am:
                        break

            if not placed_am:
                day_choices = sorted(DAY_SHIFTS_ORDER, key=lambda s: doc_shift_counts[doc][s])
                if info["ratio_smur"] > 0 and doc_smur_hours[doc] < doc_smur_targets[doc]:
                    day_choices = ["SMUR_JOUR"] + [s for s in day_choices if s != "SMUR_JOUR"]
                for sk in day_choices:
                    if schedule[d][sk] is None and can_assign_base(
                        doc, d, [sk], enforce_90_night_cap=False, is_explicit_want=True
                    ):
                        assign_shifts(doc, d, [sk])
                        doc_wants_granted[doc] += 1
                        break

    # ==========================================================================
    # PASSE 1A: PLACEMENT ÉQUILIBRÉ VEN / SAM / DIM (24H) POUR KENFACK ELODIE FLORE
    # ==========================================================================
    kenfack_docs = [d for d in docs_list if doc_constraints[d].get("is_kenfack", False)]
    for k_doc in kenfack_docs:
        k_target = doc_schedule_targets[k_doc]
        for _ in range(12):
            if doc_hours[k_doc] >= k_target - 10.0:
                break
            wd_counts_k = {
                4: doc_wd_counts[k_doc][4],
                5: doc_saturdays[k_doc],
                6: doc_sundays[k_doc],
            }
            ordered_wds = sorted([4, 5, 6], key=lambda w: (wd_counts_k[w], random.random()))
            placed_k_step = False
            for target_wd in ordered_wds:
                cand_dates = [d for d in dates if d.weekday() == target_wd and not worked_any_on_date(k_doc, d)]
                random.shuffle(cand_dates)
                for d_k in cand_dates:
                    for sd in ["UHCD_1", "UHCD_2"]:
                        for sn in ["GARDE_1", "GARDE_2", "GARDE_3_SAUV"]:
                            if schedule[d_k][sd] is None and schedule[d_k][sn] is None:
                                if can_assign_base(
                                    k_doc, d_k, [sd, sn],
                                    enforce_90_night_cap=False,
                                    enforce_smur_night_ratio=False,
                                    relax_soft_limits=True,
                                ):
                                    assign_shifts(k_doc, d_k, [sd, sn])
                                    placed_k_step = True
                                    break
                        if placed_k_step:
                            break
                    if placed_k_step:
                        break
                if placed_k_step:
                    break
            if not placed_k_step:
                break

    # ==========================================================================
    # PASSE 1B: COMBINAISONS VALIDES (24H & JOURNÉE + DEMI-GARDE)
    # ==========================================================================
    dates_combo_order = list(dates)
    random.shuffle(dates_combo_order)
    dates_combo_order.sort(
        key=lambda d: (
            0 if d.weekday() in {5, 6} else (1 if d.weekday() == 4 else 2),
            -sum(1 for a in schedule[d].values() if a is not None),
        )
    )

    for d in dates_combo_order:
        for day_list, night_list, req in VALID_COMBOS:
            day_iter = list(day_list)
            night_iter = list(night_list)
            random.shuffle(day_iter)
            random.shuffle(night_iter)
            for sd in day_iter:
                for sn in night_iter:
                    existing_doc_day = schedule[d][sd]
                    if existing_doc_day is not None and schedule[d][sn] is not None:
                        continue
                    candidate_docs = [existing_doc_day] if existing_doc_day is not None else docs_list
                    best_doc, best_score = None, -999999.0

                    for doc in candidate_docs:
                        if doc is None:
                            continue
                        info = doc_constraints[doc]
                        des = desiderata[doc].get(d, {"AM": "FREE", "NU": "FREE"})

                        if req == "accepte_24h" and not info["accepte_24h"]:
                            continue
                        if req == "accepte_demi_garde":
                            if not info["accepte_demi_garde"] or info.get("no_enchainement_demi_garde", False):
                                continue
                            if des["NU"] not in {"WANT_DEMI", "WANT", "PREFER", "DEMI_ONLY"}:
                                continue
                        if req == "mix_urg_smur_and_24h" and (
                            not info["mix_urg_smur"] or not info["accepte_24h"]
                        ):
                            continue

                        shifts_to_assign = [sn] if existing_doc_day == doc else [sd, sn]
                        if not can_assign_base(
                            doc,
                            d,
                            shifts_to_assign,
                            enforce_90_night_cap=True,
                            enforce_smur_night_ratio=True,
                            avoid_day_before_night=True,
                            enforce_night_balance=True,
                        ):
                            continue

                        score = hour_priority_score(doc, base_weight=60000.0)
                        score += 4000.0
                        if existing_doc_day == doc:
                            score += 6000.0

                        if des["AM"] == "WANT" or des["NU"] in {"WANT", "WANT_DEMI"}:
                            score += 70000.0
                        if des["AM"] == "PREFER":
                            score += 18000.0
                        if des["NU"] == "PREFER":
                            score += 18000.0
                        if "demi" in req and des["NU"] == "DEMI_ONLY":
                            score += 20000.0
                        if des["AM"] == "AVOID" or des["NU"] == "AVOID":
                            score -= 25000.0

                        eff_fte_doc = max(0.1, doc_eff_fte.get(doc, 0.1))
                        is_smur_doc = info["ratio_smur"] > 0 or info["mix_urg_smur"]
                        garde_weight = 11000.0 if is_smur_doc else 17500.0
                        score -= (doc_gardes[doc] / eff_fte_doc) * garde_weight
                        if is_smur_doc:
                            score += 6500.0

                        rep_penalty = sum(doc_shift_counts[doc][sk] for sk in shifts_to_assign) * 900.0
                        score -= rep_penalty

                        eq_val = calculate_equity_score(doc, info)
                        score -= eq_val * 500.0
                        score += weekend_equity_score(doc, d)

                        pref = info["pref_nuit"]
                        if "demi" in req and pref == "garde":
                            score -= 800.0
                        elif "24h" in req and pref == "demi_garde":
                            score -= 800.0
                        elif "demi" in req and pref == "demi_garde":
                            score += 900.0
                        elif "24h" in req and pref == "garde":
                            score += 900.0

                        score += random.uniform(-15.0, 15.0)
                        if score > best_score:
                            best_score = score
                            best_doc = doc

                    if best_doc is not None:
                        shifts_to_assign = [sn] if existing_doc_day == best_doc else [sd, sn]
                        assign_shifts(best_doc, d, shifts_to_assign)

    # ==========================================================================
    # PASSE 2: REMPLISSAGE DES GARDES DE NUIT (SMUR & URGENCES) PRIORITAIRES
    # ==========================================================================
    dates_nights = list(dates)
    random.shuffle(dates_nights)
    dates_nights.sort(key=lambda d: (0 if d.weekday() == 6 else (1 if d.weekday() == 5 else 2)))

    for d in dates_nights:
        for sk in NIGHT_GARDES_ORDER:
            if schedule[d][sk] is not None:
                continue
            best_doc, best_score = None, -999999.0
            for doc in docs_list:
                info = doc_constraints[doc]
                des = desiderata[doc].get(d, {"AM": "FREE", "NU": "FREE"})
                is_want_night = des["NU"] in {"WANT", "PREFER"}
                if not can_assign_base(
                    doc,
                    d,
                    [sk],
                    enforce_90_night_cap=False,
                    enforce_smur_night_ratio=True,
                    enforce_night_balance=not is_want_night,
                ):
                    continue

                is_night_only = info.get("nuit_uniquement", False)
                score = hour_priority_score(doc, base_weight=55000.0)
                if is_night_only:
                    score += 32000.0

                if sk == "SMUR_NUIT":
                    smur_def = doc_smur_targets[doc] - doc_smur_hours[doc]
                    score += smur_def * 120.0
                    if doc_smur_nights[doc] <= doc_urg_nights[doc]:
                        score += 12000.0

                if des["NU"] == "WANT":
                    score += 70000.0
                elif des["NU"] == "PREFER":
                    score += 20000.0
                elif des["NU"] == "AVOID":
                    score -= 25000.0

                if info["pref_nuit"] == "garde":
                    score += 800.0
                elif info["pref_nuit"] == "demi_garde":
                    score -= 600.0

                if not is_night_only:
                    eff_fte_doc = max(0.1, doc_eff_fte.get(doc, 0.1))
                    is_smur_doc = info["ratio_smur"] > 0 or info["mix_urg_smur"]
                    eq_val = calculate_equity_score(doc, info)
                    score -= eq_val * 700.0
                    garde_weight = 9500.0 if is_smur_doc else 15500.0
                    score -= (doc_gardes[doc] / eff_fte_doc) * garde_weight
                    if is_smur_doc:
                        score += 7500.0

                score += weekend_equity_score(doc, d)
                score -= doc_shift_counts[doc][sk] * 800.0
                score += random.uniform(-10.0, 10.0)
                if score > best_score:
                    best_score = score
                    best_doc = doc

            if best_doc is not None:
                assign_shifts(best_doc, d, [sk])

    # ==========================================================================
    # PASSE 3: DEMI-GARDES SOUHAITÉES OU INDIFFÉRENTES
    # ==========================================================================
    dates_demi = list(dates)
    random.shuffle(dates_demi)

    for d in dates_demi:
        for demi_sk in DEMI_GARDES_ORDER:
            if schedule[d][demi_sk] is not None:
                continue
            best_doc, best_score = None, -999999.0
            for doc in docs_list:
                info = doc_constraints[doc]
                if not info["accepte_demi_garde"]:
                    continue
                des = desiderata[doc].get(d, {"AM": "FREE", "NU": "FREE"})
                nu_st = des["NU"]
                is_no_chain_profile = info.get("no_enchainement_demi_garde", False)
                if nu_st not in {"WANT_DEMI", "PREFER", "DEMI_ONLY"} and not is_no_chain_profile:
                    continue

                tot_deficit = doc_schedule_targets[doc] - doc_hours[doc]
                if is_no_chain_profile and nu_st == "FREE" and tot_deficit < 6.0:
                    continue

                if not can_assign_base(doc, d, [demi_sk], enforce_90_night_cap=True, enforce_smur_night_ratio=True):
                    continue

                score = max(0.0, tot_deficit) * 70.0
                if nu_st == "WANT_DEMI":
                    score += 55000.0
                elif nu_st == "PREFER":
                    score += 32000.0
                elif nu_st == "DEMI_ONLY":
                    score += 24000.0
                if is_no_chain_profile:
                    score += 14000.0

                score -= doc_shift_counts[doc][demi_sk] * 600.0
                score += weekend_equity_score(doc, d)
                score += random.uniform(-10.0, 10.0)
                if score > best_score:
                    best_score = score
                    best_doc = doc

            if best_doc is not None:
                assign_shifts(best_doc, d, [demi_sk])

    # ==========================================================================
    # PASSE 4: REMPLISSAGE DES POSTES DE JOURNÉE RESTANTS
    # ==========================================================================
    remaining_day_shifts_order = [
        "SMUR_JOUR",
        "UHCD_1", "UHCD_2", "MRA", "SAUV", "COURS_DE_CIRCUIT",
        "CIRCUIT_LONG_1", "CIRCUIT_LONG_2"
    ]

    for _ in range(2):
        for shift_k in remaining_day_shifts_order:
            dates_s = list(dates)
            random.shuffle(dates_s)
            dates_s.sort(key=lambda d: (0 if d.weekday() == 6 else (1 if d.weekday() == 5 else 2)))
            for d in dates_s:
                if schedule[d][shift_k] is not None:
                    continue
                best_doc, best_score = None, -999999.0
                for doc in docs_list:
                    if not can_assign_base(
                        doc,
                        d,
                        [shift_k],
                        enforce_90_night_cap=True,
                        enforce_smur_night_ratio=True,
                        avoid_day_before_night=True,
                    ):
                        continue

                    des = desiderata[doc].get(d, {"AM": "FREE", "NU": "FREE"})
                    st_val = des["AM"]
                    score = hour_priority_score(doc, base_weight=60000.0)
                    if shift_k == "SMUR_JOUR" and doc_smur_hours[doc] < doc_smur_targets[doc]:
                        score += 22000.0
                    existing = [s for s, a in schedule[d].items() if a == doc]
                    if len(existing) > 0:
                        score += 12000.0
                    if st_val == "WANT":
                        score += 70000.0
                    elif st_val == "PREFER":
                        score += 20000.0
                    elif st_val == "AVOID":
                        score -= 25000.0

                    score -= doc_shift_counts[doc][shift_k] * 1250.0
                    score += weekend_equity_score(doc, d)
                    score -= 6.0 * (doc_wd_counts[doc][d.weekday()] ** 2)
                    score += random.uniform(-10.0, 10.0)
                    if score > best_score:
                        best_score = score
                        best_doc = doc

                if best_doc is not None:
                    assign_shifts(best_doc, d, [shift_k])

    # ==========================================================================
    # PASSE 5: GARANTIE DU TEMPS CONTRACTUEL MINIMUM (RATTRAPAGE SOUS-PROGRAMMÉS)
    # ==========================================================================
    for _ in range(10):
        under_docs = [
            doc for doc in docs_list if doc_hours[doc] < doc_schedule_targets[doc] - 3.5
        ]
        if not under_docs:
            break
        under_docs.sort(
            key=lambda x: (doc_schedule_targets[x] - doc_hours[x]) / max(1.0, doc_schedule_targets[x]),
            reverse=True,
        )
        progress_made = False
        for doc in under_docs:
            deficit = doc_schedule_targets[doc] - doc_hours[doc]
            if deficit <= 3.5:
                continue
            info = doc_constraints[doc]
            is_night_only = info.get("nuit_uniquement", False)
            if is_night_only:
                candidate_shifts = list(NIGHT_GARDES_ORDER)
            elif info.get("is_kenfack", False):
                candidate_shifts = []
            else:
                candidate_shifts = list(DAY_SHIFTS_ORDER) + list(NIGHT_GARDES_ORDER)
                if info["accepte_demi_garde"]:
                    candidate_shifts += list(DEMI_GARDES_ORDER)

            best_choice = None
            best_score = -999999.0

            if not is_night_only and (
                info.get("is_kenfack") or deficit >= 20.0 or doc_scarcity_bonus.get(doc, 0.0) > 12000.0
            ):
                for d in dates:
                    for day_list, night_list, _ in VALID_COMBOS:
                        for sd in day_list:
                            for sn in night_list:
                                if schedule[d][sd] is None and schedule[d][sn] is None:
                                    if can_assign_base(
                                        doc,
                                        d,
                                        [sd, sn],
                                        enforce_90_night_cap=False,
                                        enforce_smur_night_ratio=False,
                                        avoid_day_before_night=False,
                                        relax_soft_limits=True,
                                    ):
                                        des = desiderata[doc].get(d, {"AM": "FREE", "NU": "FREE"})
                                        c_score = 65000.0 + weekend_equity_score(doc, d)
                                        if des["AM"] == "AVOID" or des["NU"] == "AVOID":
                                            c_score -= 15000.0
                                        if c_score > best_score:
                                            best_score = c_score
                                            best_choice = (d, [sd, sn])

            for d in dates:
                for sk in candidate_shifts:
                    if schedule[d][sk] is not None:
                        continue
                    if not can_assign_base(
                        doc,
                        d,
                        [sk],
                        enforce_90_night_cap=False,
                        enforce_smur_night_ratio=False,
                        avoid_day_before_night=False,
                        relax_soft_limits=True,
                    ):
                        continue

                    cfg = SHIFTS_CONFIG[sk]
                    des = desiderata[doc].get(d, {"AM": "FREE", "NU": "FREE"})
                    st_val = des["AM"] if cfg["type"] == "JOUR" else des["NU"]

                    existing_on_d = [s for s, a in schedule[d].items() if a == doc]
                    combo_bonus = 28000.0 if len(existing_on_d) == 1 else 0.0

                    score = 50000.0 + combo_bonus + (cfg["heures"] * 600.0)
                    if st_val == "WANT":
                        score += 40000.0
                    elif st_val == "PREFER":
                        score += 20000.0
                    elif st_val == "AVOID":
                        score -= 12000.0

                    if cfg["service"] == "SMUR" and doc_smur_hours[doc] < doc_smur_targets[doc]:
                        score += 12000.0
                    elif cfg["service"] == "URGENCES" and doc_urg_hours[doc] < doc_urg_targets[doc]:
                        score += 12000.0

                    if not is_night_only and cfg["type"] in {"GARDE", "NUIT"}:
                        score -= calculate_equity_score(doc, info) * 100.0
                    score += weekend_equity_score(doc, d)
                    score -= doc_shift_counts[doc][sk] * 500.0
                    score += random.uniform(-10.0, 10.0)
                    if score > best_score:
                        best_score = score
                        best_choice = (d, [sk])

            if best_choice is not None:
                assign_shifts(doc, best_choice[0], best_choice[1])
                progress_made = True
        if not progress_made:
            break

    # ==========================================================================
    # PASSE 5B: RÉALLOCATION INTELLIGENTE POUR MÉDECINS À CONTRAINTES STRICTES
    # ==========================================================================
    for _ in range(8):
        still_under = [
            doc for doc in docs_list if doc_hours[doc] < doc_schedule_targets[doc] - 3.5
        ]
        if not still_under:
            break
        still_under.sort(
            key=lambda x: (doc_schedule_targets[x] - doc_hours[x]) / max(1.0, doc_schedule_targets[x]),
            reverse=True,
        )
        swapped_any = False
        for u_doc in still_under:
            if doc_hours[u_doc] >= doc_schedule_targets[u_doc] - 3.5:
                continue
            u_info = doc_constraints[u_doc]
            if u_info.get("is_kenfack", False):
                continue

            u_shifts = (
                list(NIGHT_GARDES_ORDER)
                if u_info.get("nuit_uniquement", False)
                else list(NIGHT_GARDES_ORDER)
                + list(DAY_SHIFTS_ORDER)
                + (list(DEMI_GARDES_ORDER) if u_info["accepte_demi_garde"] else [])
            )
            swap_done = False
            for d in dates:
                for sk in u_shifts:
                    o_doc = schedule[d][sk]
                    if o_doc is None or o_doc == u_doc or doc_constraints[o_doc].get("is_etage", False):
                        continue
                    o_shifts_on_d = [s for s, a in schedule[d].items() if a == o_doc]
                    if len(o_shifts_on_d) > 1:
                        continue

                    cfg_sk = SHIFTS_CONFIG[sk]
                    o_des = desiderata[o_doc].get(d, {"AM": "FREE", "NU": "FREE"})
                    o_st = o_des["AM"] if cfg_sk["type"] == "JOUR" else o_des["NU"]
                    if o_st in {"WANT", "WANT_DEMI"}:
                        continue

                    unassign_shift(o_doc, d, sk)
                    u_can_take = can_assign_base(
                        u_doc,
                        d,
                        [sk],
                        enforce_90_night_cap=False,
                        enforce_smur_night_ratio=False,
                        avoid_day_before_night=False,
                        relax_soft_limits=True,
                    )
                    if not u_can_take:
                        assign_shifts(o_doc, d, [sk])
                        continue

                    o_deficit_after = doc_schedule_targets[o_doc] - doc_hours[o_doc]
                    if o_deficit_after <= 2.5:
                        assign_shifts(u_doc, d, [sk])
                        swap_done = True
                        swapped_any = True
                        break

                    o_info = doc_constraints[o_doc]
                    o_alt_shifts = (
                        list(NIGHT_GARDES_ORDER)
                        if o_info.get("nuit_uniquement", False)
                        else list(DAY_SHIFTS_ORDER)
                        + list(NIGHT_GARDES_ORDER)
                        + (list(DEMI_GARDES_ORDER) if o_info["accepte_demi_garde"] else [])
                    )
                    alt_found = None
                    for d_alt in dates:
                        if d_alt == d:
                            continue
                        for sk_alt in o_alt_shifts:
                            if schedule[d_alt][sk_alt] is None and can_assign_base(
                                o_doc,
                                d_alt,
                                [sk_alt],
                                enforce_90_night_cap=False,
                                enforce_smur_night_ratio=False,
                                avoid_day_before_night=False,
                                relax_soft_limits=True,
                            ):
                                alt_found = (d_alt, sk_alt)
                                break
                        if alt_found is not None:
                            break

                    if alt_found is not None:
                        assign_shifts(u_doc, d, [sk])
                        assign_shifts(o_doc, alt_found[0], [alt_found[1]])
                        swap_done = True
                        swapped_any = True
                        break
                    else:
                        assign_shifts(o_doc, d, [sk])
                if swap_done:
                    break
        if not swapped_any:
            break

    # ==========================================================================
    # PASSE 5C: ÉQUILIBRAGE DES DIMANCHES ET DU VEN/SAM/DIM D'ELODIE FLORE
    # ==========================================================================
    for _ in range(10):
        zero_sun_docs = [
            d for d in docs_list
            if not doc_constraints[d].get("is_kenfack")
            and doc_saturdays[d] >= 1
            and doc_sundays[d] == 0
            and doc_eff_fte[d] >= 0.25
        ]
        if not zero_sun_docs:
            break
        progress_sun = False
        for doc_a in zero_sun_docs:
            if doc_sundays[doc_a] > 0:
                continue
            sat_dates_a = [dt for dt in dates if dt.weekday() == 5 and worked_any_on_date(doc_a, dt)]
            sun_dates_all = [dt for dt in dates if dt.weekday() == 6 and not worked_any_on_date(doc_a, dt)]
            swapped_sun = False
            for d_sat in sat_dates_a:
                shifts_a = [s for s, a in schedule[d_sat].items() if a == doc_a]
                if len(shifts_a) != 1:
                    continue
                sk_a = shifts_a[0]
                des_a_sat = desiderata[doc_a].get(d_sat, {"AM": "FREE", "NU": "FREE"})
                st_a_sat = des_a_sat["AM"] if SHIFTS_CONFIG[sk_a]["type"] == "JOUR" else des_a_sat["NU"]
                if st_a_sat in {"WANT", "WANT_DEMI"}:
                    continue

                for d_sun in sun_dates_all:
                    if d_sun == d_sat + datetime.timedelta(days=1):
                        continue
                    doc_b = schedule[d_sun][sk_a]
                    if doc_b is None or doc_b == doc_a:
                        continue
                    if doc_constraints[doc_b].get("is_etage") or doc_constraints[doc_b].get("is_kenfack"):
                        continue
                    if doc_sundays[doc_b] <= 1:
                        continue
                    shifts_b = [s for s, a in schedule[d_sun].items() if a == doc_b]
                    if len(shifts_b) != 1:
                        continue
                    des_b_sun = desiderata[doc_b].get(d_sun, {"AM": "FREE", "NU": "FREE"})
                    st_b_sun = des_b_sun["AM"] if SHIFTS_CONFIG[sk_a]["type"] == "JOUR" else des_b_sun["NU"]
                    if st_b_sun in {"WANT", "WANT_DEMI"}:
                        continue

                    unassign_shift(doc_a, d_sat, sk_a)
                    unassign_shift(doc_b, d_sun, sk_a)
                    if (
                        can_assign_base(doc_a, d_sun, [sk_a], relax_soft_limits=True)
                        and can_assign_base(doc_b, d_sat, [sk_a], relax_soft_limits=True)
                    ):
                        assign_shifts(doc_a, d_sun, [sk_a])
                        assign_shifts(doc_b, d_sat, [sk_a])
                        swapped_sun = True
                        progress_sun = True
                        break
                    else:
                        assign_shifts(doc_a, d_sat, [sk_a])
                        assign_shifts(doc_b, d_sun, [sk_a])
                if swapped_sun:
                    break
        if not progress_sun:
            break

    # Rééquilibrage Ven / Sam / Dim en 24H pour Kenfack Elodie Flore
    for k_doc in kenfack_docs:
        for _ in range(6):
            counts_k = {
                4: doc_wd_counts[k_doc][4],
                5: doc_saturdays[k_doc],
                6: doc_sundays[k_doc],
            }
            max_wd = max(counts_k, key=counts_k.get)
            min_wd = min(counts_k, key=counts_k.get)
            if counts_k[max_wd] - counts_k[min_wd] <= 1:
                break

            src_dates = [dt for dt in dates if dt.weekday() == max_wd and is_24h_on_date(k_doc, dt)]
            dst_dates = [dt for dt in dates if dt.weekday() == min_wd and not worked_any_on_date(k_doc, dt)]
            moved_k = False
            for d_src in src_dates:
                src_shifts = [s for s, a in schedule[d_src].items() if a == k_doc]
                if len(src_shifts) != 2:
                    continue
                des_k_src = desiderata[k_doc].get(d_src, {"AM": "FREE", "NU": "FREE"})
                if des_k_src["AM"] == "WANT" or des_k_src["NU"] == "WANT":
                    continue

                for d_dst in dst_dates:
                    for sd in ["UHCD_1", "UHCD_2"]:
                        for sn in ["GARDE_1", "GARDE_2", "GARDE_3_SAUV"]:
                            if schedule[d_dst][sd] is None and schedule[d_dst][sn] is None:
                                for s_old in src_shifts:
                                    unassign_shift(k_doc, d_src, s_old)
                                if can_assign_base(
                                    k_doc, d_dst, [sd, sn],
                                    enforce_90_night_cap=False,
                                    enforce_smur_night_ratio=False,
                                    relax_soft_limits=True,
                                ):
                                    assign_shifts(k_doc, d_dst, [sd, sn])
                                    moved_k = True
                                    break
                                else:
                                    assign_shifts(k_doc, d_src, src_shifts)
                        if moved_k:
                            break
                    if moved_k:
                        break
                if moved_k:
                    break
            if not moved_k:
                break

    # ==========================================================================
    # PASSE 5D: RÉALLOCATION ULTIME POUR LES MÉDECINS ENCORE EN DÉFICIT (> 15H)
    # ==========================================================================
    for _ in range(10):
        still_severely_under = [
            doc for doc in docs_list
            if doc_hours[doc] < doc_schedule_targets[doc] - 15.0 and not doc_constraints[doc].get("is_kenfack")
        ]
        if not still_severely_under:
            break
        still_severely_under.sort(key=lambda x: doc_schedule_targets[x] - doc_hours[x], reverse=True)
        swapped_force = False

        for u_doc in still_severely_under:
            u_info = doc_constraints[u_doc]
            u_shifts = (
                list(NIGHT_GARDES_ORDER)
                if u_info.get("nuit_uniquement", False)
                else list(NIGHT_GARDES_ORDER)
                + list(DAY_SHIFTS_ORDER)
                + (list(DEMI_GARDES_ORDER) if u_info["accepte_demi_garde"] else [])
            )
            for d in dates:
                for sk in u_shifts:
                    o_doc = schedule[d][sk]
                    if o_doc is None or o_doc == u_doc or doc_constraints[o_doc].get("is_etage", False):
                        continue
                    o_shifts_on_d = [s for s, a in schedule[d].items() if a == o_doc]
                    if len(o_shifts_on_d) > 1:
                        continue
                    cfg_sk = SHIFTS_CONFIG[sk]
                    o_des = desiderata[o_doc].get(d, {"AM": "FREE", "NU": "FREE"})
                    o_st = o_des["AM"] if cfg_sk["type"] == "JOUR" else o_des["NU"]
                    if o_st in {"WANT", "WANT_DEMI"}:
                        continue

                    if doc_hours[o_doc] - cfg_sk["heures"] < doc_schedule_targets[o_doc] - 6.0:
                        continue

                    unassign_shift(o_doc, d, sk)
                    if can_assign_base(
                        u_doc,
                        d,
                        [sk],
                        enforce_90_night_cap=False,
                        enforce_smur_night_ratio=False,
                        avoid_day_before_night=False,
                        relax_soft_limits=True,
                    ):
                        assign_shifts(u_doc, d, [sk])
                        swapped_force = True
                        break
                    else:
                        assign_shifts(o_doc, d, [sk])
                if swapped_force:
                    break
            if swapped_force:
                break
        if not swapped_force:
            break

    # ==========================================================================
    # CALCUL DES STATISTIQUES ET SYNTHÈSE
    # ==========================================================================
    total_combos = 0
    for d in dates:
        doc_shifts = {}
        for s_k, assigned_doc in schedule[d].items():
            if assigned_doc:
                doc_shifts.setdefault(assigned_doc, []).append(SHIFTS_CONFIG[s_k]["type"])
        for _, shift_types in doc_shifts.items():
            if any(t == "JOUR" for t in shift_types) and any(t in {"NUIT", "GARDE", "DEMI_GARDE"} for t in shift_types):
                total_combos += 1

    total_posts, assigned_posts = 0, 0
    stats_breakdown = {
        "JOUR_SMUR": {"nom": "SMUR Journée", "total": 0, "assigned": 0},
        "NUIT_SMUR": {"nom": "SMUR Nuit (Garde)", "total": 0, "assigned": 0},
        "JOUR_URG": {"nom": "Urgences Journée", "total": 0, "assigned": 0},
        "GARDE_URG": {"nom": "Urgences Nuit (Gardes complètes)", "total": 0, "assigned": 0},
        "DEMI_URG": {"nom": "Urgences Nuit (Demi-Gardes)", "total": 0, "assigned": 0},
    }

    for d in dates:
        for s_k, assigned in schedule[d].items():
            cfg = SHIFTS_CONFIG[s_k]
            total_posts += 1
            if cfg["service"] == "SMUR":
                cat = "JOUR_SMUR" if cfg["type"] == "JOUR" else "NUIT_SMUR"
            else:
                if cfg["type"] == "JOUR":
                    cat = "JOUR_URG"
                elif cfg["type"] == "GARDE":
                    cat = "GARDE_URG"
                else:
                    cat = "DEMI_URG"
            stats_breakdown[cat]["total"] += 1
            if assigned is not None:
                assigned_posts += 1
                stats_breakdown[cat]["assigned"] += 1

    return {
        "schedule": schedule,
        "doc_hours": doc_hours,
        "doc_urg_hours": doc_urg_hours,
        "doc_smur_hours": doc_smur_hours,
        "doc_targets_total": doc_targets_base,
        "doc_schedule_targets": doc_schedule_targets,
        "doc_smur_targets": doc_smur_targets,
        "doc_urg_targets": doc_urg_targets,
        "doc_gardes": doc_gardes,
        "doc_smur_nights": doc_smur_nights,
        "doc_urg_nights": doc_urg_nights,
        "doc_demi_gardes": doc_demi_gardes,
        "doc_weekends": {k: len(v) for k, v in doc_worked_weekends.items()},
        "doc_saturdays": doc_saturdays,
        "doc_sundays": doc_sundays,
        "doc_days_worked": {k: {w: len(v) for w, v in days.items()} for k, days in doc_days_worked.items()},
        "doc_shift_counts": doc_shift_counts,
        "total_posts": total_posts,
        "assigned_posts": assigned_posts,
        "vacant_posts": total_posts - assigned_posts,
        "stats_breakdown": stats_breakdown,
        "total_combos": total_combos,
    }


# ==============================================================================
# 5. AUDIT DE CONFORMITÉ & EXPORT EXCEL / JSON
# ==============================================================================
def run_audit_checks(schedule, doc_constraints, desiderata, dates, prev_we_data):
    """Vérifie automatiquement le respect des vœux forts, indisponibilités et repos."""
    anomalies = []
    wants_total, wants_ok = 0, 0
    wants_missed_list = []

    for doc, info in doc_constraints.items():
        for idx, d in enumerate(dates):
            des = desiderata.get(doc, {}).get(d, {})
            shifts_d = [s for s, a in schedule[d].items() if a == doc]
            has_day = any(SHIFTS_CONFIG[s]["type"] == "JOUR" for s in shifts_d)
            has_night = any(SHIFTS_CONFIG[s]["type"] in {"NUIT", "GARDE", "DEMI_GARDE"} for s in shifts_d)
            has_full_garde = any(SHIFTS_CONFIG[s]["type"] in {"NUIT", "GARDE"} for s in shifts_d)
            has_demi = any(SHIFTS_CONFIG[s]["type"] == "DEMI_GARDE" for s in shifts_d)

            if des.get("AM") == "BLOCKED" and has_day:
                anomalies.append(f"{doc} ({d.strftime('%d/%m/%Y')}): Affecté en journée malgré indisponibilité (AM).")
            if des.get("NU") == "BLOCKED" and has_night:
                anomalies.append(f"{doc} ({d.strftime('%d/%m/%Y')}): Affecté de nuit/demi-garde malgré indisponibilité (NU).")
            if des.get("NU") == "WANT_DEMI" and has_full_garde:
                anomalies.append(f"{doc} ({d.strftime('%d/%m/%Y')}): Placé de garde complète malgré 'Veut demi-garde seulement'.")

            if des.get("AM") == "WANT":
                wants_total += 1
                if has_day:
                    wants_ok += 1
                else:
                    wants_missed_list.append((doc, d.strftime("%d/%m/%Y"), "Veut (AM)"))

            if des.get("NU") in {"WANT", "WANT_DEMI"}:
                wants_total += 1
                if (des.get("NU") == "WANT_DEMI" and has_demi) or (des.get("NU") == "WANT" and has_night):
                    wants_ok += 1
                else:
                    wants_missed_list.append(
                        (doc, d.strftime("%d/%m/%Y"), "Veut demi-garde (NU)" if des.get("NU") == "WANT_DEMI" else "Veut (NU)")
                    )

            if has_night and idx + 1 < len(dates):
                next_shifts = [s for s, a in schedule[dates[idx + 1]].items() if a == doc]
                if next_shifts:
                    anomalies.append(
                        f"{doc} ({dates[idx + 1].strftime('%d/%m/%Y')}): Travaille le lendemain d'une nuit/demi-garde."
                    )

            if idx == 0 and prev_we_data.get(doc, {}).get("sun_night") and shifts_d:
                anomalies.append(f"{doc} ({d.strftime('%d/%m/%Y')}): Travaille le lundi après garde du dimanche précédent.")

    return {
        "anomalies": anomalies,
        "wants_total": wants_total,
        "wants_ok": wants_ok,
        "wants_missed_list": wants_missed_list,
    }


def serialize_simulation_to_json(state_dict):
    """Sérialise une simulation complète en chaîne JSON sauvegardable."""
    results = state_dict["results"]
    dates = state_dict["dates"]
    desiderata = state_dict["desiderata"]

    sched_ser = {d.isoformat(): shifts for d, shifts in results["schedule"].items()}
    desid_ser = {
        doc: {d.isoformat(): vals for d, vals in d_map.items()}
        for doc, d_map in desiderata.items()
    }
    payload = {
        "version": "GHNE_V2",
        "saved_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "allow_overtime": state_dict.get("allow_overtime", False),
        "dates": [d.isoformat() for d in dates],
        "doc_constraints": state_dict["doc_constraints"],
        "desiderata": desid_ser,
        "conges_days": state_dict.get("conges_days", {}),
        "rtt_days": state_dict.get("rtt_days", {}),
        "formation_days": state_dict.get("formation_days", {}),
        "prev_we_data": state_dict.get("prev_we_data", {}),
        "match_report": state_dict.get("match_report", {}),
        "results": {
            **results,
            "schedule": sched_ser,
        },
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def deserialize_simulation_from_json(json_str):
    """Restaure proprement une simulation depuis JSON (avec clés entières restaurées)."""
    payload = json.loads(json_str)
    dates = [datetime.date.fromisoformat(ds) for ds in payload["dates"]]
    res_raw = payload["results"]
    sched = {datetime.date.fromisoformat(ds): shifts for ds, shifts in res_raw["schedule"].items()}

    days_worked_fixed = {}
    for doc, wd_map in res_raw.get("doc_days_worked", {}).items():
        days_worked_fixed[doc] = {int(k): int(v) for k, v in wd_map.items()}

    results = {
        **res_raw,
        "schedule": sched,
        "doc_days_worked": days_worked_fixed,
    }
    desiderata = {
        doc: {datetime.date.fromisoformat(ds): vals for ds, vals in d_map.items()}
        for doc, d_map in payload.get("desiderata", {}).items()
    }
    return {
        "allow_overtime": payload.get("allow_overtime", False),
        "dates": dates,
        "doc_constraints": payload["doc_constraints"],
        "desiderata": desiderata,
        "conges_days": payload.get("conges_days", {}),
        "rtt_days": payload.get("rtt_days", {}),
        "formation_days": payload.get("formation_days", {}),
        "prev_we_data": payload.get("prev_we_data", {}),
        "match_report": payload.get("match_report", {}),
        "results": results,
    }


def build_formatted_excel_report(results, doc_constraints, dates, conges_days, rtt_days, formation_days):
    """Génère un classeur Excel multi-onglets mis en forme."""
    if not HAS_OPENPYXL:
        return None

    wb = openpyxl.Workbook()
    ws_rep = wb.active
    ws_rep.title = "Reporting Médecins"

    hdr_fill = PatternFill(start_color="003A54", end_color="003A54", fill_type="solid")
    hdr_font = Font(color="FFFFFF", bold=True, size=10)
    we_fill = PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid")
    vac_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    rep_headers = [
        "Médecin", "Profil", "ETP Contractuel", "ETP SMUR", "ETP Urgences",
        "Heures Cibles", "Heures Réalisées", "Écart (h)", "SMUR (h)", "Urgences (h)",
        "Gardes SMUR", "Gardes Urgences", "Demi-Gardes", "Week-ends",
        "Samedis", "Dimanches", "Congés (j)", "RTT (j)", "Formation (j)"
    ]
    ws_rep.append(rep_headers)
    for col_i in range(1, len(rep_headers) + 1):
        cell = ws_rep.cell(row=1, column=col_i)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for doc in sorted(doc_constraints.keys()):
        info = doc_constraints[doc]
        fte = info["ratio_urg"] + info["ratio_smur"]
        h_tot = results["doc_hours"][doc]
        t_base = results["doc_targets_total"][doc]
        profil = "Médecin Étage" if info.get("is_etage") else ("Urgences et Smur" if (info["ratio_smur"] > 0 or info["mix_urg_smur"]) else "Urgence")
        ws_rep.append([
            doc,
            profil,
            round(fte, 2),
            round(info["ratio_smur"], 2),
            round(info["ratio_urg"], 2),
            round(t_base, 1),
            round(h_tot, 1),
            round(h_tot - t_base, 1),
            round(results["doc_smur_hours"][doc], 1),
            round(results["doc_urg_hours"][doc], 1),
            results["doc_smur_nights"][doc],
            results["doc_urg_nights"][doc],
            results["doc_demi_gardes"][doc],
            results["doc_weekends"][doc],
            results.get("doc_saturdays", {}).get(doc, 0),
            results.get("doc_sundays", {}).get(doc, 0),
            conges_days.get(doc, 0),
            rtt_days.get(doc, 0),
            formation_days.get(doc, 0),
        ])

    for col in ws_rep.columns:
        ws_rep.column_dimensions[get_column_letter(col[0].column)].width = 16

    months_avail = sorted(list(set((d.year, d.month) for d in dates)))
    for y, m in months_avail:
        ws_m = wb.create_sheet(title=f"{FRENCH_MONTHS[m]} {y}"[:31])
        m_headers = ["Date", "Jour", "Semaine"] + [SHIFTS_CONFIG[c]["nom"] for c in MATRIX_COLUMNS_ORDER]
        ws_m.append(m_headers)
        for col_i in range(1, len(m_headers) + 1):
            c = ws_m.cell(row=1, column=col_i)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal="center", vertical="center")

        m_dates = [d for d in dates if d.year == y and d.month == m]
        for r_idx, d in enumerate(m_dates, start=2):
            row_vals = [
                d.strftime("%d/%m/%Y"),
                FRENCH_DAYS_FULL[d.weekday()],
                f"S {d.isocalendar()[1]:02d}",
            ] + [results["schedule"][d][c] or "NON POURVU" for c in MATRIX_COLUMNS_ORDER]
            ws_m.append(row_vals)
            is_we = d.weekday() >= 5
            for c_idx in range(1, len(row_vals) + 1):
                cell = ws_m.cell(row=r_idx, column=c_idx)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if cell.value == "NON POURVU":
                    cell.fill = vac_fill
                elif is_we:
                    cell.fill = we_fill

        for col in ws_m.columns:
            ws_m.column_dimensions[get_column_letter(col[0].column)].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ==============================================================================
# 6. RENDUS HTML: PLANNING MENSUEL INVERSÉ & CALENDRIER INDIVIDUEL
# ==============================================================================
def render_monthly_matrix_html(month_dates, schedule, highlight_doc=None):
    """Génère le tableau mensuel inversé du premier onglet."""
    smur_cols = [c for c in MATRIX_COLUMNS_ORDER if SHIFTS_CONFIG[c]["service"] == "SMUR"]
    urg_cols = [c for c in MATRIX_COLUMNS_ORDER if SHIFTS_CONFIG[c]["service"] == "URGENCES"]

    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        margin: 0;
        padding: 4px;
        background: #FFFFFF;
    }
    .table-wrap {
        width: 100%;
        overflow-x: auto;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
    }
    table {
        width: 100%;
        border-collapse: collapse;
        font-size: 11px;
        min-width: 1280px;
    }
    th, td {
        border: 1px solid #CBD5E1;
        padding: 4px 5px;
        text-align: center;
        vertical-align: middle;
    }
    th.group-smur {
        background-color: #F48FB1;
        color: #1E293B;
        font-weight: 800;
        font-size: 12px;
        letter-spacing: 0.5px;
    }
    th.group-urg {
        background-color: #81D4FA;
        color: #0F172A;
        font-weight: 800;
        font-size: 12px;
        letter-spacing: 0.5px;
    }
    th.col-date {
        background-color: #003A54;
        color: white;
        width: 96px;
        min-width: 96px;
        position: sticky;
        left: 0;
        z-index: 3;
    }
    th.sub-smur {
        background-color: #FCE4EC;
        color: #880E4F;
        font-weight: 700;
        font-size: 10.5px;
    }
    th.sub-urg {
        background-color: #E1F5FE;
        color: #01579B;
        font-weight: 700;
        font-size: 10.5px;
    }
    td.date-cell {
        font-weight: 700;
        text-align: left;
        padding-left: 7px;
        background-color: #F8FAFC;
        color: #1E293B;
        position: sticky;
        left: 0;
        z-index: 2;
        white-space: nowrap;
    }
    tr.weekend-row td {
        background-color: #B3B3F5 !important;
    }
    tr.weekend-row td.date-cell {
        background-color: #9FA0EC !important;
        color: #111827;
    }
    .week-badge {
        display: inline-block;
        margin-top: 2px;
        padding: 1px 5px;
        background-color: #003A54;
        color: #FFFFFF;
        border-radius: 3px;
        font-size: 9.5px;
        font-weight: 700;
    }
    .doc-pill {
        display: inline-block;
        padding: 2px 6px;
        border-radius: 4px;
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        color: #0F172A;
        font-weight: 600;
        font-size: 10.5px;
        box-shadow: 0 1px 1px rgba(0,0,0,0.04);
    }
    .doc-pill-night {
        background-color: #FFF8E1;
        border-color: #FFE082;
    }
    .doc-pill-highlight {
        background-color: #F59E0B !important;
        border: 2px solid #B45309 !important;
        color: #FFFFFF !important;
        font-weight: 800 !important;
        box-shadow: 0 2px 5px rgba(245, 158, 11, 0.45);
    }
    .vacant-pill {
        display: inline-block;
        padding: 2px 5px;
        border-radius: 4px;
        background-color: #FEE2E2;
        border: 1px solid #FCA5A5;
        color: #991B1B;
        font-weight: 700;
        font-size: 9.5px;
    }
    </style>
    </head>
    <body>
    <div class="table-wrap">
    <table>
    <thead>
    <tr>
        <th class="col-date" rowspan="2">Jour / Date</th>
        <th class="group-smur" colspan="2">SMUR</th>
        <th class="group-urg" colspan="12">URGENCES</th>
    </tr>
    <tr>
    """

    for col_k in smur_cols:
        cfg = SHIFTS_CONFIG[col_k]
        html += f"<th class='sub-smur'>{cfg['nom']}<br/><span style='font-weight:400;font-size:9px;'>{cfg['debut']}-{cfg['fin']}</span></th>"
    for col_k in urg_cols:
        cfg = SHIFTS_CONFIG[col_k]
        html += f"<th class='sub-urg'>{cfg['nom']}<br/><span style='font-weight:400;font-size:9px;'>{cfg['debut']}-{cfg['fin']}</span></th>"

    html += """
    </tr>
    </thead>
    <tbody>
    """

    for d in month_dates:
        is_we = d.weekday() >= 5
        row_cls = "weekend-row" if is_we else ""
        day_short = FRENCH_DAYS_SHORT[d.weekday()]
        iso_w = d.isocalendar()[1]
        week_tag = f"<br/><span class='week-badge'>[S {iso_w:02d}]</span>" if d.weekday() == 0 else ""

        html += f"<tr class='{row_cls}'>"
        html += f"<td class='date-cell'>{day_short} {d.strftime('%d/%m/%Y')}{week_tag}</td>"

        day_sched = schedule.get(d, {})
        for col_k in MATRIX_COLUMNS_ORDER:
            assigned = day_sched.get(col_k)
            is_night_type = SHIFTS_CONFIG[col_k]["type"] in {"NUIT", "GARDE", "DEMI_GARDE"}
            if assigned:
                if highlight_doc and assigned == highlight_doc:
                    pill_cls = "doc-pill doc-pill-highlight"
                elif is_night_type:
                    pill_cls = "doc-pill doc-pill-night"
                else:
                    pill_cls = "doc-pill"
                html += f"<td><span class='{pill_cls}' translate='no'>{assigned}</span></td>"
            else:
                html += "<td><span class='vacant-pill'>NON POURVU</span></td>"
        html += "</tr>"

    html += """
    </tbody>
    </table>
    </div>
    </body>
    </html>
    """
    return html


def render_monthly_calendar(doc, year, month, schedule, desiderata, dates_list):
    """Génère le calendrier mensuel individuel HTML avec Congés, R.T.T. et Formation."""
    cal = calendar.Calendar(firstweekday=0)
    month_days = cal.monthdatescalendar(year, month)
    dates_set = set(dates_list)

    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background: #FFFFFF;
        margin: 0;
        padding: 5px;
    }
    table {
        width: 100%;
        table-layout: fixed;
        border-collapse: collapse;
        font-size: 12px;
    }
    th {
        background-color: #003A54;
        color: white;
        text-align: center;
        padding: 8px;
        font-weight: 600;
    }
    th.sem-col {
        background-color: #1A758E;
        width: 56px;
    }
    td.sem-cell {
        background-color: #EEF2F6;
        color: #003A54;
        font-weight: 800;
        text-align: center;
        vertical-align: middle;
        width: 56px;
        font-size: 11px;
    }
    td.day-cell {
        border: 1px solid #DEE2E6;
        padding: 5px;
        vertical-align: top;
        height: 104px;
        width: 13.4%;
    }
    .day-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 4px;
        color: #495057;
        font-size: 11px;
    }
    .day-label {
        font-weight: 600;
        color: #6C757D;
    }
    .day-num {
        font-weight: 800;
        color: #003A54;
    }
    .content {
        text-align: center;
        line-height: 1.35;
    }
    .desid-badge {
        font-size: 9.5px;
        margin-top: 5px;
        padding: 2px 4px;
        background: rgba(0, 58, 84, 0.08);
        border-radius: 3px;
        color: #1E293B;
        font-weight: 600;
        display: block;
    }
    </style>
    </head>
    <body>
    <table>
    <tr>
        <th class="sem-col">Sem.</th>
        <th>Lundi</th>
        <th>Mardi</th>
        <th>Mercredi</th>
        <th>Jeudi</th>
        <th>Vendredi</th>
        <th>Samedi</th>
        <th>Dimanche</th>
    </tr>
    """

    for week in month_days:
        iso_week = week[0].isocalendar()[1]
        html += f"<tr><td class='sem-cell'>S {iso_week:02d}</td>"
        for d in week:
            day_short = FRENCH_DAYS_SHORT[d.weekday()]
            if d.month != month:
                html += "<td class='day-cell' style='background-color:#F8F9FA; color:#ADB5BD;'></td>"
                continue

            if d in dates_set:
                shifts = [s for s, assigned in schedule[d].items() if assigned == doc]
                desid = desiderata.get(doc, {}).get(d, {})
                am_st = desid.get("AM", "FREE")
                nu_st = desid.get("NU", "FREE")
                is_conges = desid.get("is_conges", False)
                is_rtt = desid.get("is_rtt", False)
                is_formation = desid.get("is_formation", False)
                is_maladie = desid.get("is_maladie", False)

                desid_labels = []
                if am_st == "WANT":
                    desid_labels.append("Veut (AM)")
                elif am_st == "PREFER":
                    desid_labels.append("Préfère (AM)")
                elif am_st == "AVOID":
                    desid_labels.append("Éviter (AM)")
                elif am_st == "BLOCKED" and not (is_conges or is_rtt or is_formation or is_maladie):
                    desid_labels.append("Indispo (AM)")

                if nu_st == "WANT_DEMI":
                    desid_labels.append("Veut demi-garde (NU)")
                elif nu_st == "WANT":
                    desid_labels.append("Veut (NU)")
                elif nu_st == "PREFER":
                    desid_labels.append("Préfère (NU)")
                elif nu_st == "AVOID":
                    desid_labels.append("Éviter (NU)")
                elif nu_st == "DEMI_ONLY":
                    desid_labels.append("Demi-garde (NU)")
                elif nu_st == "BLOCKED" and not (is_conges or is_rtt or is_formation or is_maladie):
                    desid_labels.append("Indispo (NU)")

                desid_html = f"<div class='desid-badge'>{' | '.join(desid_labels)}</div>" if desid_labels else ""

                if shifts:
                    bg_color = "#E3F2FD" if len(shifts) == 1 else "#D1FAE5"
                    day_html, night_html = [], []
                    for s in shifts:
                        cfg = SHIFTS_CONFIG[s]
                        if cfg["type"] == "JOUR":
                            day_html.append(f"☀️ <b>{cfg['nom']}</b>")
                        else:
                            night_html.append(f"🌙 <b>{cfg['nom']}</b>")
                    content = "<br/>".join(day_html + night_html)
                elif is_conges:
                    bg_color = "#D4EDDA"
                    content = "<span style='color:#155724; font-weight:bold;'>🌴 Congés</span>"
                elif is_rtt:
                    bg_color = "#CCFBF1"
                    content = "<span style='color:#115E59; font-weight:bold;'>🍃 R.T.T.</span>"
                elif is_formation:
                    bg_color = "#E2D9F3"
                    content = "<span style='color:#4A235A; font-weight:bold;'>🎓 Formation</span>"
                elif is_maladie:
                    bg_color = "#F8D7DA"
                    content = "<span style='color:#721C24; font-weight:bold;'>🏥 Arrêt</span>"
                elif am_st == "BLOCKED" and nu_st == "BLOCKED":
                    bg_color = "#F1F5F9"
                    content = "<span style='color:#475569; font-weight:bold;'>⛔ Indisponible</span>"
                else:
                    bg_color = "#F8FAFC" if d.weekday() < 5 else "#EEF2FF"
                    content = "<span style='color:#6C757D; font-style: italic;'>Repos</span>"

                content += desid_html
                html += f"<td class='day-cell' style='background-color:{bg_color};'>"
                html += f"<div class='day-header'><span class='day-label'>{day_short}</span><span class='day-num'>{d.day:02d}</span></div>"
                html += f"<div class='content'>{content}</div>"
                html += "</td>"
            else:
                html += "<td class='day-cell' style='background-color:#F8F9FA;'></td>"
        html += "</tr>"

    html += "</table></body></html>"
    return html


# ==============================================================================
# 7. POP-UP DE SYNTHÈSE DE CROISEMENT CONTRAINTES <-> DÉSIDÉRATAS
# ==============================================================================
@st.dialog("Synthèse du croisement Contraintes ↔ Désidératas")
def show_matching_dialog(match_report):
    tot = match_report.get("total", 0)
    ok_cnt = match_report.get("matched_count", 0)
    missing = match_report.get("unmatched_docs", [])
    if not missing:
        st.success(f"Tous les médecins du fichier Contraintes ont été retrouvés dans les Désidératas ({ok_cnt}/{tot}).")
    else:
        st.warning(
            f"**{ok_cnt}/{tot}** médecins retrouvés dans les Désidératas. "
            f"**{len(missing)}** praticien(s) non retrouvé(s) :"
        )
        for m_name in missing:
            st.markdown(f"- **{m_name}**")
    if st.button("Fermer", type="primary", use_container_width=True):
        st.rerun()


# ==============================================================================
# 8. INTERFACE STREAMLIT PRINCIPALE
# ==============================================================================
if "saved_simulations" not in st.session_state:
    st.session_state["saved_simulations"] = {}

with st.sidebar:
    st.markdown("### 💾 Sauvegarde & Chargement")
    if "results" in st.session_state:
        save_label = st.text_input(
            "Nom de la sauvegarde",
            value=f"Simu_{datetime.datetime.now().strftime('%H%M')}",
        )
        if st.button("💾 Sauvegarder cette simulation", use_container_width=True):
            st.session_state["saved_simulations"][save_label] = {
                "results": st.session_state["results"],
                "dates": st.session_state["dates"],
                "doc_constraints": st.session_state["doc_constraints"],
                "desiderata": st.session_state["desiderata"],
                "conges_days": st.session_state["conges_days"],
                "rtt_days": st.session_state.get("rtt_days", {}),
                "formation_days": st.session_state.get("formation_days", {}),
                "prev_we_data": st.session_state.get("prev_we_data", {}),
                "match_report": st.session_state.get("match_report", {}),
                "allow_overtime": st.session_state.get("allow_overtime", False),
            }
            st.success(f"Simulation '{save_label}' mémorisée !")

        json_bytes = serialize_simulation_to_json(st.session_state).encode("utf-8")
        st.download_button(
            label="⬇️ Télécharger la sauvegarde (.json)",
            data=json_bytes,
            file_name=f"{save_label}.json",
            mime="application/json",
            use_container_width=True,
        )

    saved_names = sorted(list(st.session_state["saved_simulations"].keys()))
    if saved_names:
        st.divider()
        sel_saved = st.selectbox("Simulations mémorisées", saved_names)
        if st.button("📂 Charger la simulation sélectionnée", use_container_width=True):
            loaded = st.session_state["saved_simulations"][sel_saved]
            for k_load, v_load in loaded.items():
                st.session_state[k_load] = v_load
            st.success(f"Simulation '{sel_saved}' chargée !")
            st.rerun()

    st.divider()
    uploaded_json = st.file_uploader("📤 Recharger un fichier (.json)", type=["json"])
    if uploaded_json is not None:
        if st.button("🔄 Restaurer depuis ce fichier JSON", use_container_width=True):
            try:
                restored = deserialize_simulation_from_json(uploaded_json.read().decode("utf-8"))
                for k_res, v_res in restored.items():
                    st.session_state[k_res] = v_res
                st.success("Simulation restaurée avec succès !")
                st.rerun()
            except Exception as err:
                st.error(f"Impossible de lire ce fichier JSON : {err}")

tab1, tab2, tab3 = st.tabs([
    "📅 Génération & Planning Mensuel",
    "📊 Rapports, Alertes & Calendriers",
    "🔍 Audit Désidératas & Contraintes",
])

with tab1:
    st.markdown("### 1. Dépôt des Fichiers")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        f_desid = st.file_uploader("📋 Déposer 'Désidératas' (.xlsx)", type=["xlsx", "xls"])
    with c2:
        f_contr = st.file_uploader("⚙️ Déposer 'Contraintes' (.xlsx)", type=["xlsx", "xls"])
    with c3:
        f_mois = st.file_uploader("🗓️ 'Mois travaillés' (Optionnel)", type=["xlsx", "xls"])
    with c4:
        f_prev_we = st.file_uploader("⏮️ 'Week-end précédent' (Optionnel)", type=["xlsx", "xls"])

    st.divider()
    cb1, cb2 = st.columns(2)
    with cb1:
        btn_gen = st.button("🚀 Générer le planning optimisé", type="primary", use_container_width=True)
    with cb2:
        btn_ot = st.button("⏱️ Activer / Désactiver Heures Supplémentaires", use_container_width=True)

    if "allow_overtime" not in st.session_state:
        st.session_state.allow_overtime = False
    if btn_ot:
        st.session_state.allow_overtime = not st.session_state.allow_overtime

    if (btn_gen or btn_ot) and f_desid and f_contr:
        with st.spinner("Génération en cours : médecins d'étages, désidératas, équité week-ends & dimanches..."):
            doc_constraints_raw = parse_contraintes_excel(f_contr)
            (
                dates,
                desiderata,
                conges_days,
                rtt_days,
                absence_days,
                formation_days,
                doc_constraints_clean,
                match_report,
            ) = parse_desiderata_excel(f_desid, doc_constraints_raw)

            unavail_months = (
                parse_mois_travailles(f_mois, doc_constraints_clean) if f_mois else {}
            )
            prev_we_data = (
                parse_prev_weekend(f_prev_we, doc_constraints_clean) if f_prev_we else {}
            )

            results = solve_planning(
                doc_constraints_clean,
                desiderata,
                dates,
                conges_days,
                absence_days,
                unavail_months,
                prev_we_data=prev_we_data,
                allow_overtime=st.session_state.allow_overtime,
            )

            st.session_state["results"] = results
            st.session_state["dates"] = dates
            st.session_state["doc_constraints"] = doc_constraints_clean
            st.session_state["desiderata"] = desiderata
            st.session_state["conges_days"] = conges_days
            st.session_state["rtt_days"] = rtt_days
            st.session_state["formation_days"] = formation_days
            st.session_state["prev_we_data"] = prev_we_data
            st.session_state["match_report"] = match_report
            st.session_state["show_match_popup"] = True

    if "results" in st.session_state:
        results = st.session_state["results"]
        dates = st.session_state["dates"]
        match_report = st.session_state.get("match_report", {})

        if st.session_state.pop("show_match_popup", False) and match_report:
            show_matching_dialog(match_report)

        ot_status_str = "ACTIVÉES" if st.session_state.allow_overtime else "DÉSACTIVÉES (100% Contrat)"
        st.success(
            f"✅ Planning calculé sur **{len(dates)} jours** pour **{len(st.session_state['doc_constraints'])} médecins confirmés**. "
            f"Heures Supp : **{ot_status_str}**."
        )

        if match_report:
            missing_list = match_report.get("unmatched_docs", [])
            if missing_list:
                st.warning(
                    f"⚠️ Correspondance Contraintes ↔ Désidératas : **{match_report['matched_count']}/{match_report['total']}** "
                    f"retrouvés. Non retrouvés : **{', '.join(missing_list)}**."
                )
            else:
                st.info(
                    f"🔗 Correspondance Contraintes ↔ Désidératas : **100% ({match_report['matched_count']}/{match_report['total']})** des médecins retrouvés."
                )

        st.markdown("### 2. Synthèse & Couverture")
        k1, k2, k3, k4, k5 = st.columns(5)

        smur_n_ass = results["stats_breakdown"]["NUIT_SMUR"]["assigned"]
        smur_n_tot = results["stats_breakdown"]["NUIT_SMUR"]["total"]
        pct_smur_n = (smur_n_ass / smur_n_tot * 100.0) if smur_n_tot > 0 else 0.0

        urg_g_ass = results["stats_breakdown"]["GARDE_URG"]["assigned"]
        urg_g_tot = results["stats_breakdown"]["GARDE_URG"]["total"]
        pct_urg_g = (urg_g_ass / urg_g_tot * 100.0) if urg_g_tot > 0 else 0.0

        pct_glob = (results["assigned_posts"] / results["total_posts"] * 100.0) if results["total_posts"] > 0 else 0.0

        with k1:
            st.markdown(
                f"<div class='metric-card'><div class='metric-title'>Gardes SMUR Nuit</div>"
                f"<div class='metric-value' style='color:#00A896;'>{pct_smur_n:.1f}% ({smur_n_ass}/{smur_n_tot})</div></div>",
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                f"<div class='metric-card'><div class='metric-title'>Gardes Urgences</div>"
                f"<div class='metric-value' style='color:#1A758E;'>{pct_urg_g:.1f}% ({urg_g_ass}/{urg_g_tot})</div></div>",
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                f"<div class='metric-card' style='border-left-color:#8E44AD;'><div class='metric-title'>Combinaisons Réalisées</div>"
                f"<div class='metric-value' style='color:#8E44AD;'>{results['total_combos']}</div></div>",
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                f"<div class='metric-card' style='border-left-color:#28A745;'><div class='metric-title'>Couverture Totale</div>"
                f"<div class='metric-value' style='color:#28A745;'>{pct_glob:.1f}%</div></div>",
                unsafe_allow_html=True,
            )
        cv = RED_ALERT if results["vacant_posts"] > 0 else "#28A745"
        with k5:
            st.markdown(
                f"<div class='metric-card' style='border-left-color:{cv};'><div class='metric-title'>Postes Vacants</div>"
                f"<div class='metric-value' style='color:{cv};'>{results['vacant_posts']}</div></div>",
                unsafe_allow_html=True,
            )

        summary_data = []
        for _, stat in results["stats_breakdown"].items():
            tot, ass = stat["total"], stat["assigned"]
            vac = tot - ass
            p_cat = (ass / tot * 100.0) if tot > 0 else 0.0
            summary_data.append({
                "Service / Période": stat["nom"],
                "Postes Requis": tot,
                "Postes Positionnés": ass,
                "Postes Vacants": f"{vac} NON POURVU(S)" if vac > 0 else "Complet (0)",
                "Taux Couverture": f"{p_cat:.1f}%",
            })
        st.table(pd.DataFrame(summary_data))

        st.divider()
        st.markdown("### 3. Planning Mensuel Global (Postes en haut / Jours à gauche)")
        months_avail_t1 = sorted(list(set((d.year, d.month) for d in dates)))
        months_labels_t1 = [f"{FRENCH_MONTHS[m]} {y}" for y, m in months_avail_t1]

        if "t1_month_idx" not in st.session_state or st.session_state.t1_month_idx >= len(months_labels_t1):
            st.session_state.t1_month_idx = 0
        if "t1_month_sel" not in st.session_state or st.session_state.t1_month_sel not in months_labels_t1:
            st.session_state.t1_month_sel = months_labels_t1[st.session_state.t1_month_idx]
        else:
            st.session_state.t1_month_idx = months_labels_t1.index(st.session_state.t1_month_sel)

        def prev_t1_month():
            idx = months_labels_t1.index(st.session_state.t1_month_sel)
            if idx > 0:
                st.session_state.t1_month_idx = idx - 1
                st.session_state.t1_month_sel = months_labels_t1[idx - 1]

        def next_t1_month():
            idx = months_labels_t1.index(st.session_state.t1_month_sel)
            if idx < len(months_labels_t1) - 1:
                st.session_state.t1_month_idx = idx + 1
                st.session_state.t1_month_sel = months_labels_t1[idx + 1]

        def set_t1_month():
            st.session_state.t1_month_idx = months_labels_t1.index(st.session_state.t1_month_sel)

        col_m1, col_m2, col_m3, col_hl = st.columns([0.6, 2.0, 0.6, 2.8])
        with col_m1:
            st.button("◀", on_click=prev_t1_month, key="t1_prev_m", use_container_width=True)
        with col_m2:
            st.selectbox(
                "Mois affiché",
                months_labels_t1,
                key="t1_month_sel",
                on_change=set_t1_month,
                label_visibility="collapsed",
            )
        with col_m3:
            st.button("▶", on_click=next_t1_month, key="t1_next_m", use_container_width=True)
        with col_hl:
            doc_options = ["- Surligner un médecin (Aucun) -"] + sorted(list(st.session_state["doc_constraints"].keys()))
            selected_hl = st.selectbox(
                "Surligner un médecin",
                doc_options,
                index=0,
                label_visibility="collapsed",
            )
        highlight_doc = None if selected_hl.startswith("-") else selected_hl

        active_t1_idx = months_labels_t1.index(st.session_state.t1_month_sel)
        sel_y1, sel_m1 = months_avail_t1[active_t1_idx]
        month_dates_t1 = [d for d in dates if d.year == sel_y1 and d.month == sel_m1]

        matrix_html = render_monthly_matrix_html(month_dates_t1, results["schedule"], highlight_doc=highlight_doc)
        table_height = min(920, max(360, 110 + len(month_dates_t1) * 32))
        components.html(matrix_html, height=table_height, scrolling=True)

        st.divider()
        st.markdown("### 4. Exportation du Planning")
        grid_rows = []
        for d in dates:
            for s_k in MATRIX_COLUMNS_ORDER:
                cfg = SHIFTS_CONFIG[s_k]
                doc_assigned = results["schedule"][d][s_k]
                grid_rows.append({
                    "Date": d.strftime("%d/%m/%Y"),
                    "Jour": FRENCH_DAYS_FULL[d.weekday()],
                    "Semaine": f"Semaine {d.isocalendar()[1]:02d}",
                    "Service": cfg["service"],
                    "Poste": cfg["nom"],
                    "Horaires": f"{cfg['debut']} - {cfg['fin']}",
                    "Médecin Affecté": doc_assigned if doc_assigned else "NON POURVU",
                    "Statut": "POURVU" if doc_assigned else "VACANT",
                })
        df_grid = pd.DataFrame(grid_rows)
        csv_grid = df_grid.to_csv(index=False).encode("utf-8-sig")

        exp_c1, exp_c2 = st.columns(2)
        with exp_c1:
            st.download_button(
                label="📥 Exporter le Planning Complet (.csv)",
                data=csv_grid,
                file_name=f"Planning_Complet_GHNE_{dates[0].strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with exp_c2:
            if HAS_OPENPYXL:
                xlsx_bytes = build_formatted_excel_report(
                    results,
                    st.session_state["doc_constraints"],
                    dates,
                    st.session_state["conges_days"],
                    st.session_state.get("rtt_days", {}),
                    st.session_state.get("formation_days", {}),
                )
                st.download_button(
                    label="📊 Exporter le Classeur Excel Mis en Forme (.xlsx)",
                    data=xlsx_bytes,
                    file_name=f"Planning_GHNE_{dates[0].strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
    else:
        st.info("ℹ️ Veuillez déposer vos fichiers Excel (**Désidératas** et **Contraintes**) puis cliquer sur **'Générer le planning optimisé'** (ou charger une sauvegarde depuis la barre latérale).")

with tab2:
    if "results" in st.session_state:
        st.markdown("### Tableau de Reporting Interactif par Médecin")
        st.caption("Filtrez par statut d'alerte, par type d'habilitation, ou triez directement en cliquant sur les colonnes.")

        results = st.session_state["results"]
        doc_constraints = st.session_state["doc_constraints"]
        desiderata = st.session_state["desiderata"]
        conges_days = st.session_state["conges_days"]
        rtt_days = st.session_state.get("rtt_days", {})
        formation_days = st.session_state.get("formation_days", {})
        dates_list = st.session_state["dates"]

        rep_rows = []
        for doc in sorted(doc_constraints.keys()):
            info = doc_constraints[doc]
            h_urg = results["doc_urg_hours"][doc]
            h_smur = results["doc_smur_hours"][doc]
            h_tot = results["doc_hours"][doc]
            target_base = results["doc_targets_total"][doc]
            delta = h_tot - target_base
            pct_overtime = (delta / target_base * 100.0) if target_base > 0 else 0.0

            if info.get("is_etage", False):
                status_alert = "🏥 Médecin Étage"
                hab_type = "Médecin Étage"
            else:
                if delta < -8.0:
                    status_alert = "⚠️ Sous-programmé (<0%)"
                elif pct_overtime > 25.0:
                    status_alert = "🔴 Dépassement élevé (>25%)"
                else:
                    status_alert = "✔️ Équilibré / Heures Supp Réglées"

                if info["ratio_smur"] > 0 or info["mix_urg_smur"]:
                    hab_type = "Urgences et Smur"
                else:
                    hab_type = "Urgence"

            r_smur = (h_smur / h_tot * 100.0) if h_tot > 0 else 0.0
            r_urg = (h_urg / h_tot * 100.0) if h_tot > 0 else 0.0
            ratio_str = f"{r_smur:.0f}% SMUR / {r_urg:.0f}% URG"

            dw = results["doc_days_worked"].get(doc, {})
            fte = info["ratio_urg"] + info["ratio_smur"]

            if info.get("is_etage", False) or info.get("nuit_uniquement", False) or not info["accepte_garde"]:
                eq_val = np.nan
            else:
                eq_val = round(
                    (results["doc_gardes"][doc] + 0.5 * results["doc_demi_gardes"][doc]) / max(0.1, fte),
                    2,
                )

            rep_rows.append({
                "Statut / Alerte": status_alert,
                "Médecin": doc,
                "Option Heures Supp": info["temps_add_label"],
                "Habilitation": hab_type,
                "Préférence": info["pref_nuit"].capitalize(),
                "ETP Contractuel": round(fte, 2),
                "ETP SMUR": round(info["ratio_smur"], 2),
                "ETP Urgences": round(info["ratio_urg"], 2),
                "Heures Cibles": round(target_base, 1),
                "Heures Réalisées": round(h_tot, 1),
                "Écart / Heures Supp": f"{delta:+.1f} h ({pct_overtime:+.1f}%)" if not info.get("is_etage") else "-",
                "SMUR (h)": round(h_smur, 1),
                "Urgences (h)": round(h_urg, 1),
                "Ratio Effectif (SMUR / URG)": ratio_str,
                "Gardes SMUR": results["doc_smur_nights"][doc],
                "Gardes Urgences": results["doc_urg_nights"][doc],
                "Demi-Gardes": results["doc_demi_gardes"][doc],
                "Score Équité ETP": eq_val,
                "Week-ends Travaillés": results["doc_weekends"][doc],
                "Samedis": results.get("doc_saturdays", {}).get(doc, 0),
                "Dimanches": results.get("doc_sundays", {}).get(doc, 0),
                "Jours Congés": conges_days.get(doc, 0),
                "Jours RTT": rtt_days.get(doc, 0),
                "Jours Formation": formation_days.get(doc, 0),
                "Lundi": dw.get(0, dw.get("0", 0)),
                "Mardi": dw.get(1, dw.get("1", 0)),
                "Mercredi": dw.get(2, dw.get("2", 0)),
                "Jeudi": dw.get(3, dw.get("3", 0)),
                "Vendredi": dw.get(4, dw.get("4", 0)),
            })

        df_rep = pd.DataFrame(rep_rows)
        col_f1, col_f2, col_f3 = st.columns([2, 2, 2])
        with col_f1:
            search_name = st.text_input("🔍 Rechercher un médecin", "")
        with col_f2:
            filter_status = st.multiselect(
                "Filtre Statut / Alerte",
                [
                    "✔️ Équilibré / Heures Supp Réglées",
                    "⚠️ Sous-programmé (<0%)",
                    "🔴 Dépassement élevé (>25%)",
                    "🏥 Médecin Étage",
                ],
                default=[],
            )
        with col_f3:
            filter_hab = st.selectbox(
                "Type d'habilitation",
                ["Tous les praticiens", "Urgence", "Urgences et Smur", "Médecin Étage"],
            )

        df_filtered = df_rep.copy()
        if search_name:
            df_filtered = df_filtered[
                df_filtered["Médecin"].str.lower().str.contains(search_name.lower())
            ]
        if filter_status:
            df_filtered = df_filtered[df_filtered["Statut / Alerte"].isin(filter_status)]
        if filter_hab != "Tous les praticiens":
            df_filtered = df_filtered[df_filtered["Habilitation"] == filter_hab]

        st.dataframe(
            df_filtered,
            use_container_width=True,
            height=380,
            column_config={
                "Statut / Alerte": st.column_config.TextColumn("Statut / Alerte", width="medium"),
                "Médecin": st.column_config.TextColumn("Médecin", width="small"),
                "Option Heures Supp": st.column_config.TextColumn("Option Heures Supp", width="medium"),
                "Heures Réalisées": st.column_config.NumberColumn("Heures Réalisées", format="%.1f h"),
                "Heures Cibles": st.column_config.NumberColumn("Heures Cibles", format="%.1f h"),
            },
        )

        csv_rep = df_filtered.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="📥 Exporter le Tableau de Reporting Filtré (CSV)",
            data=csv_rep,
            file_name="Reporting_Medecins_GHNE_Filtre.csv",
            mime="text/csv",
        )

        st.divider()
        st.markdown("### 🗓️ Calendrier Mensuel Individuel")
        doc_list_sorted = sorted(list(doc_constraints.keys()))
        months_available = sorted(list(set((d.year, d.month) for d in dates_list)))
        months_labels = [f"{FRENCH_MONTHS[m]} {y}" for y, m in months_available]

        if "current_doc_idx" not in st.session_state or st.session_state.current_doc_idx >= len(doc_list_sorted):
            st.session_state.current_doc_idx = 0
        if "doc_selector" not in st.session_state or st.session_state.doc_selector not in doc_list_sorted:
            st.session_state.doc_selector = doc_list_sorted[st.session_state.current_doc_idx]
        else:
            st.session_state.current_doc_idx = doc_list_sorted.index(st.session_state.doc_selector)

        if "current_month_idx" not in st.session_state or st.session_state.current_month_idx >= len(months_labels):
            st.session_state.current_month_idx = 0
        if "month_selector" not in st.session_state or st.session_state.month_selector not in months_labels:
            st.session_state.month_selector = months_labels[st.session_state.current_month_idx]
        else:
            st.session_state.current_month_idx = months_labels.index(st.session_state.month_selector)

        def prev_doc():
            idx = doc_list_sorted.index(st.session_state.doc_selector)
            if idx > 0:
                st.session_state.current_doc_idx = idx - 1
                st.session_state.doc_selector = doc_list_sorted[idx - 1]

        def next_doc():
            idx = doc_list_sorted.index(st.session_state.doc_selector)
            if idx < len(doc_list_sorted) - 1:
                st.session_state.current_doc_idx = idx + 1
                st.session_state.doc_selector = doc_list_sorted[idx + 1]

        def set_doc():
            st.session_state.current_doc_idx = doc_list_sorted.index(st.session_state.doc_selector)

        def prev_month():
            idx = months_labels.index(st.session_state.month_selector)
            if idx > 0:
                st.session_state.current_month_idx = idx - 1
                st.session_state.month_selector = months_labels[idx - 1]

        def next_month():
            idx = months_labels.index(st.session_state.month_selector)
            if idx < len(months_labels) - 1:
                st.session_state.current_month_idx = idx + 1
                st.session_state.month_selector = months_labels[idx + 1]

        def set_month():
            st.session_state.current_month_idx = months_labels.index(st.session_state.month_selector)

        c1, c2, c3, c4, c5, c6 = st.columns([0.5, 2, 0.5, 0.5, 2, 0.5])
        with c1:
            st.button("◀", on_click=prev_doc, key="p_doc", use_container_width=True)
        with c2:
            st.selectbox(
                "Médecin",
                doc_list_sorted,
                key="doc_selector",
                on_change=set_doc,
                label_visibility="collapsed",
            )
        with c3:
            st.button("▶", on_click=next_doc, key="n_doc", use_container_width=True)
        with c4:
            st.button("◀", on_click=prev_month, key="p_mon", use_container_width=True)
        with c5:
            st.selectbox(
                "Mois",
                months_labels,
                key="month_selector",
                on_change=set_month,
                label_visibility="collapsed",
            )
        with c6:
            st.button("▶", on_click=next_month, key="n_mon", use_container_width=True)

        selected_doc = st.session_state.doc_selector
        active_month_idx = months_labels.index(st.session_state.month_selector)
        selected_year, selected_month = months_available[active_month_idx]
        current_month_title = st.session_state.month_selector

        st.markdown(
            f"<h4 translate='no'>Planning individuel de <b>{selected_doc}</b> — <i>{current_month_title}</i></h4>",
            unsafe_allow_html=True,
        )
        cal_html = render_monthly_calendar(
            selected_doc, selected_year, selected_month, results["schedule"], desiderata, dates_list
        )
        components.html(cal_html, height=640, scrolling=True)

        ind_rows = []
        for d in dates_list:
            if d.year == selected_year and d.month == selected_month:
                shifts = [s for s, assigned in results["schedule"][d].items() if assigned == selected_doc]
                desid = desiderata.get(selected_doc, {}).get(d, {})
                d_labels = []
                if desid.get("AM") == "WANT":
                    d_labels.append("Veut (AM)")
                elif desid.get("AM") == "PREFER":
                    d_labels.append("Préfère (AM)")
                elif desid.get("AM") == "AVOID":
                    d_labels.append("Éviter (AM)")
                elif desid.get("AM") == "BLOCKED" and not (
                    desid.get("is_conges") or desid.get("is_rtt") or desid.get("is_formation") or desid.get("is_maladie")
                ):
                    d_labels.append("Indispo (AM)")

                if desid.get("NU") == "WANT_DEMI":
                    d_labels.append("Veut demi-garde (NU)")
                elif desid.get("NU") == "WANT":
                    d_labels.append("Veut (NU)")
                elif desid.get("NU") == "PREFER":
                    d_labels.append("Préfère (NU)")
                elif desid.get("NU") == "AVOID":
                    d_labels.append("Éviter (NU)")
                elif desid.get("NU") == "DEMI_ONLY":
                    d_labels.append("Demi-garde (NU)")
                elif desid.get("NU") == "BLOCKED" and not (
                    desid.get("is_conges") or desid.get("is_rtt") or desid.get("is_formation") or desid.get("is_maladie")
                ):
                    d_labels.append("Indispo (NU)")

                desid_txt = " | ".join(d_labels) if d_labels else "-"
                status_str = "Repos"
                if shifts:
                    status_str = " / ".join(SHIFTS_CONFIG[s]["nom"] for s in shifts)
                elif desid.get("is_conges"):
                    status_str = "Congés"
                elif desid.get("is_rtt"):
                    status_str = "R.T.T."
                elif desid.get("is_formation"):
                    status_str = "Formation"
                elif desid.get("is_maladie"):
                    status_str = "Arrêt Maladie"
                elif desid.get("AM") == "BLOCKED" and desid.get("NU") == "BLOCKED":
                    status_str = "Indisponible"

                ind_rows.append({
                    "Semaine": f"S {d.isocalendar()[1]:02d}",
                    "Date": d.strftime("%d/%m/%Y"),
                    "Jour": FRENCH_DAYS_FULL[d.weekday()],
                    "Médecin": selected_doc,
                    "Désidérata Saisi": desid_txt,
                    "Affectation / Statut": status_str,
                })

        df_ind = pd.DataFrame(ind_rows)
        csv_ind = df_ind.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label=f"📥 Exporter le Planning Mensuel de {selected_doc} ({current_month_title}) (CSV)",
            data=csv_ind,
            file_name=f"Planning_{selected_doc}_{selected_year}_{selected_month:02d}.csv",
            mime="text/csv",
        )
    else:
        st.info("ℹ️ Veuillez d'abord charger vos fichiers Excel et cliquer sur **'Générer le planning optimisé'** dans le premier onglet.")

with tab3:
    if "results" in st.session_state:
        st.markdown("### 🔍 Audit de Conformité des Désidératas & Contraintes")
        audit = run_audit_checks(
            st.session_state["results"]["schedule"],
            st.session_state["doc_constraints"],
            st.session_state["desiderata"],
            st.session_state["dates"],
            st.session_state.get("prev_we_data", {}),
        )
        ac1, ac2 = st.columns(2)
        with ac1:
            st.metric("Anomalies Strictes Détectées", len(audit["anomalies"]))
        with ac2:
            pct_w = (audit["wants_ok"] / audit["wants_total"] * 100.0) if audit["wants_total"] > 0 else 100.0
            st.metric("Souhaits Forts ('Veut') Exaucés", f"{audit['wants_ok']}/{audit['wants_total']} ({pct_w:.1f}%)")

        if not audit["anomalies"]:
            st.success("✅ Aucune anomalie réglementaire ou violation de contrainte stricte détectée !")
        else:
            for anom in audit["anomalies"]:
                st.error(f"⚠️ {anom}")

        if audit["wants_missed_list"]:
            st.markdown("#### Détail des souhaits forts ('Veut') non pourvus")
            df_missed = pd.DataFrame(audit["wants_missed_list"], columns=["Médecin", "Date", "Souhait"])
            st.dataframe(df_missed, use_container_width=True)
    else:
        st.info("ℹ️ Générez ou chargez un planning pour afficher l'audit de conformité.")