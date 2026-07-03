"""
Outil de valorisation par adresse — version Streamlit (locale)
----------------------------------------------------------------
Installation (une seule fois) :
    pip install streamlit pandas openpyxl requests

Lancement :
    streamlit run app.py

Charge le fichier Excel au format du gabarit (onglet COMPARABLES),
géocode via l'API Adresse du gouvernement (api-adresse.data.gouv.fr,
gratuite, sans clé), filtre par rayon et calcule une valeur
basse/cible/haute pondérée.
"""

import io
import math
import re
import time
from datetime import datetime

import folium
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import live_search

st.set_page_config(page_title="Valorisation — Savills", page_icon="🏢", layout="wide")

# ── Charte graphique Savills ──────────────────────────────────────────────
import base64 as _b64
with open("/mnt/user-data/uploads/1783066337485_image.png", "rb") as _f:
    _logo_b64 = _b64.b64encode(_f.read()).decode()

SAVILLS_CSS = f"""
<style>
/* ── Police Montserrat (plus proche de Gotham accessible librement) ── */
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700&display=swap');

/* ── Reset global ── */
html, body, [class*="css"] {{
  font-family: 'Montserrat', 'Gotham', 'Segoe UI', sans-serif !important;
  color: #25273A;
}}

/* ── Fond général ── */
.stApp {{ background-color: #EEE8E3; }}
[data-testid="stSidebar"] {{ background-color: #25273A !important; }}
[data-testid="stSidebar"] * {{ color: #EEE8E3 !important; }}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {{ color: #FFDF00 !important; font-weight: 700; }}
[data-testid="stSidebar"] label {{ color: #79828C !important; }}
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] select {{ background: #2f3145 !important; color: white !important; border-color: #79828C !important; }}

/* ── Header Savills ── */
.savills-header {{
  display: flex; align-items: center; justify-content: space-between;
  background: #25273A; padding: 12px 28px;
  border-bottom: 3px solid #FFDF00; margin-bottom: 16px;
}}
.savills-header h1 {{
  color: white; font-size: 20px; font-weight: 700;
  margin: 0; letter-spacing: 0.3px;
}}
.savills-header span {{ color: #79828C; font-size: 12px; }}

/* ── Titres ── */
h1 {{ font-size: 26px !important; font-weight: 700 !important; color: #25273A !important; }}
h2 {{ font-size: 22px !important; font-weight: 700 !important; color: #008493 !important; }}
h3 {{ font-size: 18px !important; font-weight: 700 !important; color: #25273A !important; }}

/* ── Métriques ── */
[data-testid="stMetric"] {{
  background: #FFFFFF; border-radius: 8px;
  padding: 12px 16px; border-left: 3px solid #008493;
}}
[data-testid="stMetricLabel"] {{ font-size: 12px !important; color: #79828C !important; font-weight: 600; text-transform: uppercase; }}
[data-testid="stMetricValue"] {{ font-size: 22px !important; font-weight: 700 !important; color: #25273A !important; }}
[data-testid="stMetricDelta"] {{ font-size: 11px !important; }}

/* ── Boutons ── */
.stButton > button {{
  background-color: #008493 !important; color: white !important;
  font-weight: 700 !important; border: none !important; border-radius: 6px !important;
  font-family: 'Montserrat', sans-serif !important; font-size: 14px !important;
  padding: 8px 20px !important;
}}
.stButton > button:hover {{ background-color: #006d7a !important; }}
.stDownloadButton > button {{
  background-color: #FFDF00 !important; color: #25273A !important;
  font-weight: 700 !important; border: none !important; border-radius: 6px !important;
}}
.stDownloadButton > button:hover {{ background-color: #e5c800 !important; }}

/* ── Tabs ── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {{
  background: #FFFFFF; border-radius: 8px 8px 0 0;
  border-bottom: 2px solid #008493;
}}
[data-testid="stTabs"] [data-baseweb="tab"] {{
  font-weight: 600 !important; color: #79828C !important;
  font-family: 'Montserrat', sans-serif !important;
}}
[data-testid="stTabs"] [aria-selected="true"] {{
  color: #008493 !important; border-bottom: 2px solid #008493 !important;
}}
[data-testid="stTabsContent"] > div {{ padding-top: 0 !important; }}

/* ── Tables data_editor ── */
[data-testid="stDataFrame"] th {{
  background-color: #25273A !important; color: white !important;
  font-weight: 700 !important; font-size: 13px !important;
}}
[data-testid="stDataFrame"] td {{
  font-size: 13px !important; color: #25273A !important;
}}

/* ── Cards/surfaces blanches ── */
[data-testid="stExpander"], .stDataFrame {{
  background: #FFFFFF; border-radius: 8px; border: 1px solid #D9D9D9;
}}

/* ── Barres de progression ── */
[data-testid="stProgress"] > div {{ background-color: #008493 !important; }}

/* ── Selectbox / Slider accent ── */
input[type=range] {{ accent-color: #008493; }}

/* ── Logo Savills fixé en haut à droite ── */
.savills-logo-fixed {{
  position: fixed; top: 12px; right: 24px; z-index: 9999;
  background: white; border-radius: 6px; padding: 4px 8px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.15);
}}
</style>

<!-- Logo Savills fixé en haut à droite -->
<div class="savills-logo-fixed">
  <img src="data:image/png;base64,{_logo_b64}" height="36" alt="Savills"/>
</div>

<!-- Header Savills -->
<div class="savills-header">
  <div>
    <h1>🏢 Outil de Valorisation Tertiaire</h1>
    <span>Bureaux · Locaux d'activités · Commerce — Nantes Métropole</span>
  </div>
</div>
"""
st.markdown(SAVILLS_CSS, unsafe_allow_html=True)

BAN_URL = "https://api-adresse.data.gouv.fr/search/"


# ---------------------------------------------------------------- géocodage
@st.cache_data(show_spinner=False)
def geocode(query: str):
    """Retourne un dict {lat, lon, label, city, postcode} ou None. Mis en cache."""
    if not query or not str(query).strip():
        return None
    try:
        r = requests.get(BAN_URL, params={"q": query, "limit": 1}, timeout=8)
        r.raise_for_status()
        feats = r.json().get("features", [])
        if not feats:
            return None
        f = feats[0]
        lon, lat = f["geometry"]["coordinates"]
        props = f.get("properties", {})
        return {
            "lat": lat, "lon": lon,
            "label": props.get("label", query),
            "city": props.get("city", ""),
            "postcode": props.get("postcode", ""),
        }
    except requests.RequestException:
        return None


def haversine_m(lat1, lon1, lat2, lon2):
    R = 6371000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlmb / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def to_num(v):
    try:
        if v is None or v == "":
            return None
        return float(str(v).replace(",", ".").replace(" ", ""))
    except (ValueError, TypeError):
        return None


# ------------------------------------------------- extraction type de bail / preneur
BAIL_RE = re.compile(r"bail\s+(commercial|d[ée]rogatoire|professionnel)\s*(\d/\d/\d)?", re.I)
PRENEUR_RE = re.compile(r"lou[ée]\s+(?:à|a)\s+(?:la\s+soci[ée]t[ée]\s+)?([A-ZÀ-Ü][\w\-' ]{2,40})", re.I)


def extract_type_bail(row_):
    obs = str(row_.get("Observations", "") or "")
    m = BAIL_RE.search(obs)
    if m:
        return f"Bail {m.group(1)}{' ' + m.group(2) if m.group(2) else ''}".strip()
    opv = row_.get("Operation")
    if pd.notna(opv) and opv not in ("", "Non disponible"):
        return str(opv)
    return "Non disponible"


def extract_preneur_etat(row_):
    obs = str(row_.get("Observations", "") or "")
    m = PRENEUR_RE.search(obs)
    if m:
        return m.group(1).strip()
    etat = row_.get("Etat")
    if pd.notna(etat) and etat not in ("", "Non disponible"):
        return str(etat)
    return "Non disponible"


def build_export_table(res_df: pd.DataFrame) -> pd.DataFrame:
    """Format de sortie nettoyé :
    Date | Photo | Adresse complète | Surface (m²) | Loyer HT-HC €/m²/an | Commentaire"""
    rows = []
    for _, r in res_df.iterrows():
        r = r.to_dict()

        # date
        date_val = r.get("Date_transaction") or r.get("Date", "")
        if not date_val or str(date_val).strip() in ("", "nan"):
            date_val = "—"

        # adresse
        adresse = str(r.get("Adresse", "") or "").strip()
        commune = str(r.get("Commune", "") or "").strip()
        cp = str(r.get("Code_postal", "") or "").strip()
        if commune and commune not in adresse:
            adresse_complete = f"{adresse}, {cp} {commune}".strip(", ")
        else:
            adresse_complete = adresse or "—"

        # loyer HT-HC en priorité, sinon facial, sinon prix vente
        loyer_ht_hc = to_num(r.get("Loyer_HT_HC_eur_m2_an") or r.get("Loyer HT-HC €/m²/an"))
        loyer_facial = to_num(r.get("Loyer_facial_eur_m2_an") or r.get("Loyer facial €/m²/an"))
        prix_vente = to_num(r.get("Prix_vente_eur_m2") or r.get("Prix vente €/m²"))
        valeur = loyer_ht_hc or loyer_facial or prix_vente

        # commentaire éditable
        commentaire = str(r.get("Commentaire", "") or "")

        # photo
        photo = str(r.get("Photo_url", "") or "")

        rows.append({
            "Date": date_val,
            "Photo": photo,
            "Adresse complète": adresse_complete,
            "Surface (m²)": to_num(r.get("Surface_m2") or r.get("Surface (m²)")),
            "Prix €/m²": valeur,
            "Commentaire": commentaire,
        })
    return pd.DataFrame(rows)


def _download_photo(url: str, timeout: int = 6) -> bytes | None:
    """Télécharge une photo pour l'intégrer dans l'Excel. Retourne None en cas d'échec."""
    if not url or not url.startswith("http"):
        return None
    try:
        r = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code == 200 and "image" in r.headers.get("Content-Type", ""):
            return r.content
    except Exception:
        pass
    return None


def export_to_excel_bytes(export_df: pd.DataFrame, title: str = "Extraction comparables") -> bytes:
    """Export Excel charte Savills : Date | Photo | Adresse | Surface | Prix | Commentaire"""
    from openpyxl.drawing.image import Image as XLImage
    import tempfile, os

    # couleurs Savills
    C_HEAD  = "25273A"   # fond entête
    C_ACC   = "008493"   # vert Savills
    C_YEL   = "FFDF00"   # jaune Savills
    C_BG    = "EEE8E3"   # fond lignes paires

    wb = Workbook()
    ws = wb.active
    ws.title = "Extraction"

    thin = Side(style="thin", color="D9D9D9")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    COL_WIDTHS = {"Date": 14, "Photo": 22, "Adresse complète": 42,
                  "Surface (m²)": 12, "Prix €/m²": 16, "Commentaire": 42}
    COLS = list(export_df.columns)

    # ── Bandeau titre Savills ───────────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
    t = ws.cell(1, 1, f"  {title}")
    t.font  = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    t.fill  = PatternFill("solid", start_color=C_HEAD)
    t.alignment = LEFT
    ws.row_dimensions[1].height = 28

    # ── En-têtes colonnes ───────────────────────────────────────────────
    for j, h in enumerate(COLS, start=1):
        c = ws.cell(2, j, h)
        c.font  = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
        c.fill  = PatternFill("solid", start_color=C_ACC)
        c.alignment = CENTER; c.border = BORDER
        ws.column_dimensions[get_column_letter(j)].width = COL_WIDTHS.get(h, 16)
    ws.row_dimensions[2].height = 30

    # ── Données ─────────────────────────────────────────────────────────
    photo_col = COLS.index("Photo") + 1 if "Photo" in COLS else None
    ROW_H = 70  # hauteur pour accueillir la photo

    for i, row_ in export_df.iterrows():
        r = i + 3
        is_even = (i % 2 == 0)
        bg = PatternFill("solid", start_color=C_BG) if is_even else None

        for j, col in enumerate(COLS, start=1):
            val = row_[col]
            # colonne Photo : on laisse la cellule vide (l'image sera placée par-dessus)
            cell_val = "" if col == "Photo" else val
            c = ws.cell(r, j, cell_val if not isinstance(cell_val, float) or
                        not math.isnan(cell_val) else "")
            c.border = BORDER
            c.font   = Font(name="Calibri", size=10)
            c.alignment = LEFT if col in ("Adresse complète", "Commentaire") else CENTER
            if bg: c.fill = bg
            if col == "Prix €/m²" and isinstance(val, (int, float)):
                c.number_format = '#,##0 €'
            if col == "Surface (m²)" and isinstance(val, (int, float)):
                c.number_format = "#,##0"

        ws.row_dimensions[r].height = ROW_H

        # ── Intégration photo ─────────────────────────────────────────
        if photo_col:
            photo_url = str(row_.get("Photo", "") or "")
            img_bytes = _download_photo(photo_url)
            if img_bytes:
                try:
                    with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
                        tmp.write(img_bytes); tmp_path = tmp.name
                    xl_img = XLImage(tmp_path)
                    xl_img.width, xl_img.height = 100, 66
                    cell_ref = f"{get_column_letter(photo_col)}{r}"
                    ws.add_image(xl_img, cell_ref)
                    os.unlink(tmp_path)
                except Exception:
                    pass

    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Suppression du titre Streamlit par défaut (remplacé par le header Savills) ──


with st.sidebar:
    st.header("1 — Actif à valoriser")
    address = st.text_input("Adresse (recommandé pour la précision)",
                             placeholder="ex. 5 rue de la Tour, 44200 Nantes")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        manual_commune = st.text_input("Commune", value="", placeholder="ex. Nantes")
    with col_c2:
        manual_cp = st.text_input("Code postal", value="", placeholder="ex. 44000")

    if address:
        target = geocode(address)
        if not target:
            st.error("Adresse introuvable ou réseau indisponible. Vérifie l'orthographe / ajoute la commune.")
    elif manual_commune.strip():
        target = geocode(f"{manual_commune.strip()} {manual_cp.strip()}".strip())
        if target:
            st.caption("ℹ️ Pas d'adresse précise saisie : centrage sur la commune (moins précis).")
        else:
            st.error("Commune introuvable. Vérifie l'orthographe.")
    else:
        target = None

    if target:
        st.success(f"✓ {target['label']}")
        # mini-carte satellite Leaflet/Esri — pas de clé API nécessaire
        lat, lon = target["lat"], target["lon"]
        mini_map_html = f"""
        <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
        <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
        <div id="minimap" style="height:200px;border-radius:8px;overflow:hidden"></div>
        <script>
          var map = L.map('minimap', {{zoomControl:false, attributionControl:false}})
                     .setView([{lat},{lon}], 17);
          L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{{z}}/{{y}}/{{x}}',
            {{maxZoom:19}}).addTo(map);
          L.marker([{lat},{lon}]).addTo(map)
           .bindPopup("<b>{target['label']}</b>").openPopup();
        </script>
        """
        components.html(mini_map_html, height=210)
        st.caption("Vérifie que le point correspond bien au bien recherché.")

    st.header("2 — Critères")
    st.caption("Ces critères pilotent à la fois la recherche en direct et le filtrage des résultats — "
                "un seul réglage, pas de doublon.")
    op = st.selectbox("Opération", ["Location", "Vente", "Tous (Location + Vente)"])
    asset_type = st.selectbox("Type d'actif", ["Bureaux", "Activités", "Commerce", "Tous"])
    surface_target = st.number_input("Surface de l'actif (m²) — optionnel", min_value=0, value=0, step=10)
    radius_km = st.slider("Rayon de recherche (km)", 0.2, 15.0, 1.5, 0.1)

    st.subheader("Pondération du score")
    w_prox = st.slider("Proximité", 0, 100, 45)
    w_sim = st.slider("Similarité (surface)", 0, 100, 35)
    w_rec = st.slider("Récence", 0, 100, 20)

    st.header("3 — Source des comparables")
    mode = st.radio("Mode", ["Charger un fichier Excel", "Rechercher en direct (expérimental)"])

    file = None
    selected_sources = []
    if mode == "Charger un fichier Excel":
        file = st.file_uploader("Fichier Excel (gabarit, onglet COMPARABLES)", type=["xlsx"])
    else:
        if not target:
            st.warning("Renseigne d'abord l'adresse ou la commune ci-dessus : la recherche en "
                        "direct s'appuie dessus.")
        else:
            st.caption(f"📍 Recherche centrée sur **{target['city'] or target['label']}** "
                        f"({target['postcode'] or '?'}).")
        st.caption(f"🔎 Recherchera : **{op}** / **{asset_type}** (réglages de la section 2 ci-dessus).")
        selected_sources = st.multiselect(
            "Sources à interroger",
            options=list(live_search.ALL_SOURCES.keys()),
            default=["BureauxLocaux", "Geolocaux", "ArthurLoyd", "TournyMeyer (JLL)"],
            help="CBRE est marqué « non vérifié » : l'URL de recherche a été extrapolée à partir "
                 "d'un seul exemple observé, pas confirmée pour tous les cas.",
        )
        max_pages = st.slider(
            "Pages à scraper par source",
            min_value=1, max_value=10, value=3,
            help="1 page ≈ 30 annonces (BureauxLocaux). "
                 "3 pages = ~90 annonces. Plus de pages = plus long (~0.4s/page).",
        )
        only_complete = st.checkbox(
            "Ne garder que les annonces complètes (adresse précise + prix exploitable)",
            value=True,
        )
        run_live = st.button("Lancer la recherche en direct", disabled=not target)

if mode == "Charger un fichier Excel":
    if not file:
        st.info("Charge ton fichier Excel dans la barre latérale pour démarrer.")
        st.stop()
    try:
        df = pd.read_excel(file, sheet_name="COMPARABLES", skiprows=2)
    except Exception:
        df = pd.read_excel(file, skiprows=2)
    n_before = len(df)
    df = df[df["Commune"].notna() & df["Surface_m2"].notna()].reset_index(drop=True)
    has_price = df["Loyer_facial_eur_m2_an"].apply(to_num).notna() | df["Prix_vente_eur_m2"].apply(to_num).notna()
    df = df[has_price].reset_index(drop=True)
    if len(df) < n_before:
        st.caption(f"ℹ️ {n_before - len(df)} ligne(s) écartée(s) (commune, surface ou prix manquant).")
    st.write(f"**{len(df)}** référence(s) exploitable(s) chargée(s) depuis le fichier.")
else:
    if "live_results" not in st.session_state:
        st.session_state.live_results = []
    if not target:
        st.info("Renseigne d'abord l'adresse ou la commune de l'actif (section 1) pour activer "
                 "la recherche en direct.")
        st.stop()
    search_commune = target["city"] or manual_commune.strip() or "Nantes"
    search_cp = target["postcode"] or manual_cp.strip() or "44000"
    if run_live:
        ops_to_search = ["Location", "Vente"] if op.startswith("Tous") else [op]
        types_to_search = ["Bureaux", "Activités", "Commerce"] if asset_type == "Tous" else [asset_type]
        n_sources = len(selected_sources) if selected_sources else len(live_search.ALL_SOURCES)
        est_sec = n_sources * max_pages * 0.5
        with st.spinner(f"Recherche en direct — {max_pages} page(s) × {n_sources} source(s) "
                         f"(~{est_sec:.0f}s)…"):
            collected = []
            for o in ops_to_search:
                for t in types_to_search:
                    collected += live_search.search_all_sources(
                        search_commune, search_cp, o, t,
                        sources=selected_sources,
                        max_pages=max_pages,
                    )

        # ── diagnostic précis ────────────────────────────────────────────
        merged = live_search.dedupe_listings(collected)
        n_brut        = len(merged)
        n_avec_prix   = sum(1 for x in merged
                            if to_num(x.get("Loyer_facial_eur_m2_an"))
                            or to_num(x.get("Prix_vente_eur_m2")))
        n_avec_addr   = sum(1 for x in merged if x.get("Adresse_precise"))
        n_geocodable  = sum(1 for x in merged
                            if x.get("Commune") not in ("Non disponible", "", None))
        n_complets    = sum(1 for x in merged if x.get("Complete"))
        filtered      = live_search.filter_complete(merged) if only_complete else merged

        # stockage
        st.session_state.live_results = filtered
        st.session_state.live_diag = {
            "brut": n_brut, "prix": n_avec_prix,
            "addr": n_avec_addr, "geocodable": n_geocodable,
            "complets": n_complets, "filtres": len(filtered),
            "only_complete": only_complete,
        }

        if not filtered:
            diag = st.session_state.live_diag
            st.error("❌ Aucun comparable retenu — voici pourquoi :")
            st.markdown(f"""
| Étape | Résultat |
|---|---|
| Annonces brutes scrappées | **{diag['brut']}** |
| Dont avec un prix extrait | **{diag['prix']}** |
| Dont avec adresse de rue précise | **{diag['addr']}** |
| Dont géocodables (commune identifiée) | **{diag['geocodable']}** |
| Dont complètes (commune + prix) | **{diag['complets']}** |
| Après filtre « complètes seulement » | **{diag['filtres']}** |
""")
            if diag["brut"] == 0:
                st.warning("**Cause probable : JavaScript.** Le site charge ses annonces via JS — "
                            "requests ne peut pas les lire. Décoche toutes les sources sauf "
                            "BureauxLocaux et réessaie.")
            elif diag["prix"] == 0:
                st.warning("**Cause probable : format de prix non reconnu.** "
                            "Décoche « annonces complètes » pour voir les résultats partiels.")
            elif diag["complets"] == 0 and only_complete:
                st.warning(f"**Filtre trop strict.** {diag['geocodable']} annonces sont "
                            f"géocodables mais sans prix associé, ou inversement. "
                            f"**Décoche « Ne garder que les annonces complètes ».**")
            st.stop()
        else:
            n_raw = len(merged)
            msg = f"✅ {len(filtered)} annonce(s) retenue(s) sur {search_commune}"
            if only_complete and n_raw > len(filtered):
                msg += f" ({n_raw - len(filtered)} écartée(s) — adresse ou prix manquant)"
            st.success(msg)
    if not st.session_state.live_results:
        st.info("Lance la recherche en direct dans la barre latérale.")
        st.stop()
    df = pd.DataFrame(st.session_state.live_results)
    st.write(f"**{len(df)}** référence(s) retenue(s) (en direct, à valider).")
    st.dataframe(df, use_container_width=True, hide_index=True)

if not target:
    st.warning("Renseigne et valide une adresse cible pour lancer l'analyse.")
    st.stop()

t_lat, t_lon, t_label = target["lat"], target["lon"], target["label"]
radius_m = radius_km * 1000

# ---------------------------------------------------------------- géocodage des comps
progress = st.progress(0.0, text="Géocodage des comparables…")
lats, lons, dists = [], [], []
for i, row_ in df.iterrows():
    lat = to_num(row_.get("Latitude"))
    lon = to_num(row_.get("Longitude"))
    if lat is None or lon is None:
        addr_val = str(row_.get("Adresse", "") or "")
        # n'inclure l'adresse que si elle contient vraiment une info de rue
        addr_clean = "" if any(x in addr_val.lower() for x in
                               ("non disponible", "non extraite", "adresse précise")) \
                     else addr_val.strip()
        commune = str(row_.get("Commune", "") or "").strip()
        cp      = str(row_.get("Code_postal", "") or "").strip()
        # construire la query en garantissant qu'on a au moins la commune
        parts = [p for p in [addr_clean, cp, commune] if p and p != "Non disponible"]
        q = " ".join(parts)
        g = geocode(q) if q else None
        if g:
            lat, lon = g["lat"], g["lon"]
        time.sleep(0.05)
    # stocker None explicitement (pas NaN) pour les coords manquantes
    lats.append(float(lat) if lat is not None else None)
    lons.append(float(lon) if lon is not None else None)
    dists.append(haversine_m(t_lat, t_lon, lat, lon)
                 if (lat is not None and lon is not None) else None)
    progress.progress((i + 1) / max(len(df), 1), text=f"Géocodage… {i+1}/{len(df)}")
progress.empty()

df["_lat"], df["_lon"], df["_dist"] = lats, lons, dists

# ---------------------------------------------------------------- filtrage + score
def _is_unknown(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "Non disponible")

op_filter = None if op.startswith("Tous") else op
type_filter = None if asset_type == "Tous" else asset_type

def passes(row_):
    if row_["_dist"] is None or row_["_dist"] > radius_m:
        return False
    # une ligne dont l'opération/le type est inconnu n'est PAS exclue : on ne sait pas, donc on
    # ne rejette pas — seul un désaccord confirmé exclut la ligne.
    if op_filter and not _is_unknown(row_.get("Operation")) and row_.get("Operation") != op_filter:
        return False
    if type_filter and not _is_unknown(row_.get("Type_actif")) and row_.get("Type_actif") != type_filter:
        return False
    return True

res = df[df.apply(passes, axis=1)].copy()

if res.empty:
    n_total = len(df)
    n_geocoded = int(df["_lat"].notna().sum())
    n_in_radius = int((df["_dist"].notna() & (df["_dist"] <= radius_m)).sum())
    st.warning("Aucun comparable retenu avec ces critères. Détail du filtrage :")
    st.write(f"- {n_total} référence(s) au total")
    st.write(f"- {n_geocoded} géocodée(s) avec succès")
    st.write(f"- {n_in_radius} dans le rayon de {radius_km} km")
    st.write(f"- 0 après filtre Opération=« {op} » / Type=« {asset_type} »")
    if n_geocoded < n_total:
        st.caption("→ Des comparables n'ont pas pu être géocodés (adresse incomplète). "
                    "Essaie d'élargir le rayon ou de repasser en Opération/Type = Tous.")
    elif n_in_radius == 0:
        st.caption("→ Aucun comparable géocodé n'est dans le rayon choisi. Élargis le rayon.")
    else:
        st.caption("→ Des comparables sont dans le rayon mais ne correspondent pas à "
                    "l'Opération/Type choisis. Essaie « Tous ».")
    st.stop()

now = datetime.now()
def score_row(row_):
    p_prox = max(0.0, 1 - row_["_dist"] / radius_m)
    sc = to_num(row_.get("Surface_m2"))
    if surface_target and sc:
        p_sim = min(sc, surface_target) / max(sc, surface_target)
    else:
        p_sim = 0.5
    p_rec = 0.5
    dt = pd.to_datetime(row_.get("Date_transaction"), errors="coerce")
    if pd.notna(dt):
        months = (now - dt.to_pydatetime()).days / 30.4
        p_rec = max(0.0, 1 - months / 36)
    wsum = (w_prox + w_sim + w_rec) or 1
    return (w_prox * p_prox + w_sim * p_sim + w_rec * p_rec) / wsum

res["_score"] = res.apply(score_row, axis=1)
res = res.sort_values("_score", ascending=False)

if op == "Vente":
    field = "Prix_vente_eur_m2"
elif op == "Location":
    field = "Loyer_facial_eur_m2_an"
else:
    # Tous : on prend le champ le mieux renseigné parmi les résultats retenus
    n_vente = res["Prix_vente_eur_m2"].apply(to_num).notna().sum() if "Prix_vente_eur_m2" in res else 0
    n_location = res["Loyer_facial_eur_m2_an"].apply(to_num).notna().sum() if "Loyer_facial_eur_m2_an" in res else 0
    field = "Prix_vente_eur_m2" if n_vente >= n_location else "Loyer_facial_eur_m2_an"
is_vente_field = field == "Prix_vente_eur_m2"
raw_vals = res[field].apply(to_num)
bounds = ((live_search.PRIX_M2_MIN, live_search.PRIX_M2_MAX) if is_vente_field
          else (live_search.LOYER_M2_AN_MIN, live_search.LOYER_M2_AN_MAX))
plausible_mask = raw_vals.notna() & raw_vals.between(*bounds)
n_implausible = int((raw_vals.notna() & ~plausible_mask).sum())
vals = raw_vals[plausible_mask]
weights = res.loc[vals.index, "_score"]

# ---------------------------------------------------------------- résultats
st.divider()
if len(vals) == 0:
    st.warning(f"Comparables trouvés mais aucun n'a de valeur plausible dans la colonne {field}.")
    st.stop()
if n_implausible:
    st.caption(f"⚠️ {n_implausible} valeur(s) hors fourchette plausible écartée(s) du calcul.")

unit = "€/m²" if is_vente_field else "€/m²/an HT-HC"

COLS_KEEP = {
    "Adresse":                "Adresse",
    "Commune":                "Commune",
    "Type_actif":             "Type",
    "Etat":                   "État",
    "Surface_m2":             "Surface (m²)",
    "Loyer_facial_eur_m2_an": "Loyer facial €/m²/an",
    "Charges_eur_m2_an":      "Charges €/m²/an",
    "Loyer_HT_HC_eur_m2_an":  "Loyer HT-HC €/m²/an",
    "Prix_vente_eur_m2":      "Prix vente €/m²",
    "Nb_parkings":            "Parkings",
    "Date_transaction":       "Date",
    "Source":                 "Source",
    "Lien":                   "Lien",
    "Photo_url":              "Photo",
    "Commentaire":            "Commentaire",
}
COL_CONFIG = {
    "Lien":                   st.column_config.LinkColumn("Lien", display_text="🔗 Annonce"),
    "Photo":                  st.column_config.LinkColumn("Photo", display_text="🖼️ Voir"),
    "Surface (m²)":           st.column_config.NumberColumn(format="%d m²"),
    "Loyer facial €/m²/an":  st.column_config.NumberColumn(format="%.0f €"),
    "Charges €/m²/an":       st.column_config.NumberColumn(format="%.0f €"),
    "Loyer HT-HC €/m²/an":   st.column_config.NumberColumn(format="%.0f €"),
    "Prix vente €/m²":        st.column_config.NumberColumn(format="%.0f €"),
    "Parkings":               st.column_config.NumberColumn(format="%d"),
    "Commentaire":            st.column_config.TextColumn(
        "Commentaire",
        help="Ex : Preneur : XYZ. Bail 3/6/9 en date du 15/01/2025.",
        width="large",
    ),
}

def valo_bloc(df_subset: pd.DataFrame, label: str, color: str, key_suffix: str):
    """Affiche un tableau éditable + les métriques de valorisation pour un sous-groupe."""
    avail = {k: v for k, v in COLS_KEEP.items() if k in df_subset.columns}
    display = df_subset[list(avail.keys())].copy().rename(columns=avail)
    if "Loyer HT-HC €/m²/an" not in display.columns:
        display["Loyer HT-HC €/m²/an"] = None

    # métriques initiales
    v_field = "Prix vente €/m²" if is_vente_field else "Loyer HT-HC €/m²/an"
    if v_field not in display.columns:
        v_field = list(display.columns)[-1]
    raw = display[v_field].apply(to_num)
    plaus = raw[raw.between(*bounds)]
    n = len(plaus)

    if n == 0:
        st.info(f"Aucun comparable {label} avec valeur exploitable dans le rayon.")
        return display

    w = res.loc[df_subset.index, "_score"].reindex(plaus.index).fillna(0.5)
    cible = (plaus * w).sum() / w.sum() if w.sum() > 0 else plaus.mean()
    basse, haute, med = plaus.min(), plaus.max(), plaus.median()
    conf = "🟢" if n >= 8 else ("🟡" if n >= 4 else "🔴")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Loyer bas" if not is_vente_field else "Prix bas",
              f"{basse:,.0f} {unit}".replace(",", " "))
    c2.metric("Loyer cible" if not is_vente_field else "Prix cible",
              f"{cible:,.0f} {unit}".replace(",", " "))
    c3.metric("Loyer haut" if not is_vente_field else "Prix haut",
              f"{haute:,.0f} {unit}".replace(",", " "))
    c4.metric("Comparables", n)
    st.caption(f"Médiane : {med:,.0f} {unit} — Fiabilité : {conf} ({n} comp.)".replace(",", " "))
    if n < 4:
        st.warning("Peu de comparables — élargis le rayon ou ajoute des références.")

    # tableau éditable
    st.caption("Modifiable directement — les corrections sont répercutées dans l'export.")
    edited = st.data_editor(
        display, use_container_width=True, hide_index=True,
        num_rows="fixed", column_config=COL_CONFIG, key=f"editor_{key_suffix}"
    )

    # recalcul si modification
    if v_field in edited.columns:
        new_v = edited[v_field].apply(to_num).dropna()
        new_v = new_v[new_v.between(*bounds)]
        if len(new_v) > 0:
            cible_edit = new_v.mean()
            if abs(cible_edit - cible) > 0.5:
                st.info(f"💡 Après corrections → loyer cible **{cible_edit:,.0f} {unit}**".replace(",", " "))
    return edited


# ── Split Neuf/Restructuré  vs  Seconde main ──────────────────────────────
ETATS_NEUF = {"Neuf", "Restructuré"}
ETATS_SM   = {"Seconde main"}
ETATS_INCONNU = {"Non disponible", None, ""}

def etat_groupe(row):
    e = str(row.get("Etat", "") or "").strip()
    if e in ETATS_NEUF:   return "neuf_restructure"
    if e in ETATS_SM:     return "seconde_main"
    return "inconnu"

res["_groupe"] = res.apply(etat_groupe, axis=1)
res_neuf = res[res["_groupe"] == "neuf_restructure"].copy()
res_sm   = res[res["_groupe"] == "seconde_main"].copy()
res_inc  = res[res["_groupe"] == "inconnu"].copy()

total_inc = len(res_inc)
total_neuf = len(res_neuf)
total_sm = len(res_sm)

st.caption(f"**{len(res)} comparables** : {total_neuf} Neuf/Restructuré · "
            f"{total_sm} Seconde main · {total_inc} état non déterminé")

# ── Onglets principaux ────────────────────────────────────────────────────

def _valid_coord(v) -> bool:
    """Retourne True si v est un float fini (exclut None, NaN, Inf)."""
    try:
        return v is not None and math.isfinite(float(v))
    except (TypeError, ValueError):
        return False


def make_map(target_lat: float, target_lon: float, target_label: str,
             res_df: pd.DataFrame, field: str, bounds: tuple) -> folium.Map:
    """Carte Folium centrée sur l'actif cible avec tous les comparables."""

    # centrage : on ne garde que les coordonnées finies pour éviter les NaN Folium
    valid_lats = [float(v) for v in res_df["_lat"] if _valid_coord(v)]
    valid_lons = [float(v) for v in res_df["_lon"] if _valid_coord(v)]
    all_lats = [target_lat] + valid_lats
    all_lons = [target_lon] + valid_lons
    center = [sum(all_lats) / len(all_lats), sum(all_lons) / len(all_lons)]

    m = folium.Map(location=center, zoom_start=14, tiles=None)

    # ── couches de fond ──────────────────────────────────────────────────
    folium.TileLayer(
        tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        attr="Esri World Imagery",
        name="🛰️ Satellite",
        overlay=False, control=True,
    ).add_to(m)

    folium.TileLayer(
        tiles="https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png",
        attr="CartoDB",
        name="🗺️ Plan",
        overlay=False, control=True,
    ).add_to(m)

    folium.LayerControl(position="topright", collapsed=False).add_to(m)

    # ── marqueur actif cible ─────────────────────────────────────────────
    folium.Marker(
        location=[target_lat, target_lon],
        popup=folium.Popup(f"<b>🏢 ACTIF À VALORISER</b><br>{target_label}", max_width=250),
        tooltip=f"🏢 {target_label}",
        icon=folium.Icon(color="red", icon="home", prefix="fa"),
    ).add_to(m)

    # rayon de recherche
    radius_m_val = radius_km * 1000
    folium.Circle(
        location=[target_lat, target_lon],
        radius=radius_m_val,
        color="#B23A48", fill=True, fill_opacity=0.04,
        tooltip=f"Rayon {radius_km:.1f} km",
    ).add_to(m)

    # ── marqueurs comparables ─────────────────────────────────────────────
    # couleur par état
    COULEUR_ETAT = {
        "Neuf": "green", "Restructuré": "blue",
        "Seconde main": "orange", "Non disponible": "gray",
    }

    # valeur min/max pour dégradé
    vals = res_df[field].apply(to_num).dropna()
    v_min, v_max = (vals.min(), vals.max()) if len(vals) >= 2 else (0, 1)

    for _, row in res_df.iterrows():
        lat, lon = row.get("_lat"), row.get("_lon")
        if not _valid_coord(lat) or not _valid_coord(lon):
            continue
        lat, lon = float(lat), float(lon)

        val = to_num(row.get(field))
        etat = str(row.get("Etat") or "Non disponible").strip()
        adresse = str(row.get("Adresse") or "").strip()
        commune = str(row.get("Commune") or "").strip()
        surface = row.get("Surface_m2")
        lien = str(row.get("Lien") or "")
        dist = row.get("_dist")
        score = row.get("_score", 0)
        source = str(row.get("Source") or "")

        # label loyer/prix
        if val is not None:
            unit_short = "€/m²" if field == "Prix_vente_eur_m2" else "€/m²/an"
            val_str = f"{val:,.0f} {unit_short}".replace(",", "\u202f")
        else:
            val_str = "Prix non extrait"

        # popup HTML riche
        lien_html = (f'<a href="{lien}" target="_blank">🔗 Voir l\'annonce</a>'
                     if lien and lien.startswith("http") else "")
        popup_html = f"""
        <div style="font-family:Arial,sans-serif;font-size:13px;min-width:200px">
          <b>{adresse}{', ' + commune if commune else ''}</b><br>
          <span style="color:#666">{etat} · {int(surface) if surface else '?'} m²</span><br>
          <span style="font-size:15px;font-weight:bold;color:#1F3864">{val_str}</span><br>
          <span style="color:#999;font-size:11px">
            {f'{int(dist)} m' if dist else ''} · score {int(score*100) if score else '?'}/100
            · {source}
          </span><br>{lien_html}
        </div>
        """
        tooltip_str = f"{val_str} · {adresse or commune}"
        couleur = COULEUR_ETAT.get(etat, "gray")

        folium.CircleMarker(
            location=[lat, lon],
            radius=8 if val is not None else 6,
            color="white", weight=1.5,
            fill=True, fill_color=couleur, fill_opacity=0.85,
            popup=folium.Popup(popup_html, max_width=280),
            tooltip=tooltip_str,
        ).add_to(m)

    # ── légende ─────────────────────────────────────────────────────────
    legend_html = """
    <div style="position:fixed;bottom:24px;right:24px;z-index:999;background:white;
         padding:10px 14px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.2);
         font-family:Arial,sans-serif;font-size:12px">
      <b>État des locaux</b><br>
      <span style="color:green">●</span> Neuf<br>
      <span style="color:blue">●</span> Restructuré<br>
      <span style="color:orange">●</span> Seconde main<br>
      <span style="color:gray">●</span> Non déterminé<br>
      <span style="color:red">⌂</span> Actif cible
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))

    return m


tab_analyse, tab_carte, tab_env = st.tabs([
    "📊 Analyse & comparables",
    "🗺️ Carte",
    "🏙️ Environnement",
])

st.markdown("""
<style>
[data-testid="stTabsContent"] > div { padding-top: 0 !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# ONGLET 1 — Analyse
# ══════════════════════════════════════════════════════════════════════════
with tab_analyse:
    all_edited = {}

    st.subheader("🟢 Neuf / Restructuré")
    if res_neuf.empty:
        st.info("Aucun comparable Neuf ou Restructuré dans le rayon avec ces critères.")
    else:
        all_edited["neuf"] = valo_bloc(res_neuf, "Neuf/Restructuré", "#2E7D5B", "neuf")

    st.subheader("🟠 Seconde main")
    if res_sm.empty:
        st.info("Aucun comparable Seconde main dans le rayon avec ces critères.")
    else:
        all_edited["sm"] = valo_bloc(res_sm, "Seconde main", "#B26A3C", "sm")

    if not res_inc.empty:
        with st.expander(f"État non déterminé ({total_inc} comp.) — à qualifier manuellement"):
            st.caption("Qualifie l'état dans la colonne État pour que ces comparables "
                        "remontent dans le bon tableau lors de la prochaine recherche.")
            avail_inc = {k: v for k, v in COLS_KEEP.items() if k in res_inc.columns}
            disp_inc = res_inc[list(avail_inc.keys())].copy().rename(columns=avail_inc)
            all_edited["inc"] = st.data_editor(
                disp_inc, use_container_width=True, hide_index=True,
                num_rows="fixed", column_config=COL_CONFIG, key="editor_inc"
            )

    st.divider()
    st.subheader("Export pour rapport")
    frames = []
    for key in ("neuf", "sm", "inc"):
        if key in all_edited:
            df_e = all_edited[key].copy()
            df_e["_groupe_export"] = {"neuf": "Neuf / Restructuré",
                                       "sm": "Seconde main",
                                       "inc": "État non déterminé"}[key]
            frames.append(df_e)
    all_display = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    if not all_display.empty:
        # ── Fix Bug SM : re-trier par l'État ÉDITÉ par l'utilisateur
        # Si quelqu'un a changé "État non déterminé" → "Seconde main" dans le tableau,
        # cet item doit apparaître dans la section Seconde main de l'export.
        def _groupe_export(row):
            e = str(row.get("État", "") or "").strip()
            if e in ("Neuf", "Restructuré"):  return "Neuf / Restructuré"
            if e == "Seconde main":           return "Seconde main"
            return row.get("_groupe_export", "État non déterminé")
        all_display["_groupe_export"] = all_display.apply(_groupe_export, axis=1)

        # s'assurer que la colonne Commentaire existe (vide par défaut si non renseignée)
        if "Commentaire" not in all_display.columns:
            all_display["Commentaire"] = ""

        edited_for_export = all_display.rename(columns={v: k for k, v in COLS_KEEP.items()})
        export_df = build_export_table(edited_for_export)

        # aperçu sans la colonne Photo (URL brute peu lisible dans l'aperçu)
        preview = export_df.drop(columns=["Photo"], errors="ignore")
        st.dataframe(preview, use_container_width=True, hide_index=True)

        col_a, col_b = st.columns(2)
        with col_a:
            with st.spinner("Génération Excel (téléchargement des photos…)"):
                xlsx_bytes = export_to_excel_bytes(export_df, title=f"Extraction — {t_label}")
            st.download_button("📥 Extraction (Excel)",
                xlsx_bytes, file_name="extraction_comparables.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_b:
            st.download_button("Détail complet (CSV)",
                all_display.to_csv(index=False).encode("utf-8"),
                file_name="comparables_detail.csv", mime="text/csv")


# ══════════════════════════════════════════════════════════════════════════
# ONGLET 2 — Carte pleine page avec comparables auto
# ══════════════════════════════════════════════════════════════════════════
with tab_carte:
    field_for_map = ("Prix_vente_eur_m2" if op == "Vente"
                     else "Loyer_HT_HC_eur_m2_an" if "Loyer_HT_HC_eur_m2_an" in res.columns
                     else "Loyer_facial_eur_m2_an")
    bounds_for_map = ((live_search.PRIX_M2_MIN, live_search.PRIX_M2_MAX) if op == "Vente"
                      else (live_search.LOYER_M2_AN_MIN, live_search.LOYER_M2_AN_MAX))

    folium_map = make_map(t_lat, t_lon, t_label, res, field_for_map, bounds_for_map)

    # injection pour que la carte prenne tout l'espace vertical disponible
    map_html = folium_map._repr_html_()
    full_page = f"""
    <div style="width:100%;height:88vh;margin:0;padding:0;">
      {map_html}
    </div>
    <script>
      setTimeout(function(){{
        var iframes = parent.document.querySelectorAll('iframe');
        iframes.forEach(function(f){{
          if(f.contentDocument && f.contentDocument.querySelector('.folium-map')){{
            f.style.height='88vh';
          }}
        }});
      }}, 400);
    </script>
    """
    components.html(full_page, height=880, scrolling=False)
    st.caption(
        "🔴 Actif · 🟢 Neuf · 🔵 Restructuré · 🟠 Seconde main · ⚫ Inconnu "
        "| Clic → détails | Bouton ↗ : satellite ↔ plan"
    )


# ══════════════════════════════════════════════════════════════════════════
# ONGLET 3 — Environnement (OpenStreetMap / Overpass)
# ══════════════════════════════════════════════════════════════════════════
with tab_env:

    OVERPASS_URL = "https://overpass-api.de/api/interpreter"

    @st.cache_data(ttl=3600, show_spinner=False)
    def overpass_fetch(lat: float, lon: float, radius: int, filters: str) -> list:
        query = f"""[out:json][timeout:20];
(node{filters}(around:{radius},{lat},{lon});
 way{filters}(around:{radius},{lat},{lon}););
out center;"""
        try:
            r = requests.post(OVERPASS_URL, data={"data": query}, timeout=25)
            r.raise_for_status()
            return r.json().get("elements", [])
        except Exception:
            return []

    def dist_m(lat1, lon1, lat2, lon2):
        R = 6371000; p = math.pi / 180
        a = (math.sin((lat2-lat1)*p/2)**2 +
             math.cos(lat1*p)*math.cos(lat2*p)*math.sin((lon2-lon1)*p/2)**2)
        return 2 * R * math.asin(math.sqrt(a))

    def nearest(els, lat, lon):
        best, best_d = None, float("inf")
        for e in els:
            elat = e.get("lat") or (e.get("center") or {}).get("lat")
            elon = e.get("lon") or (e.get("center") or {}).get("lon")
            if elat and elon:
                d = dist_m(lat, lon, elat, elon)
                if d < best_d: best_d, best = d, e
        return best, best_d

    def wmin(m): return round(m / 80)

    st.subheader(f"🏙️ Environnement — {t_label}")
    st.caption("Données OpenStreetMap via Overpass API · Mise en cache 1h · Aucune clé API")

    with st.spinner("Analyse de l'environnement en cours…"):
        tram_els    = overpass_fetch(t_lat, t_lon, 1000, '["railway"="tram_stop"]')
        bus_els     = overpass_fetch(t_lat, t_lon,  800, '["highway"="bus_stop"]')
        train_els   = overpass_fetch(t_lat, t_lon, 3000, '["railway"="station"]')
        bike_els    = overpass_fetch(t_lat, t_lon,  500, '["amenity"="bicycle_rental"]')
        resto_300   = overpass_fetch(t_lat, t_lon,  300, '["amenity"~"restaurant|cafe|fast_food"]')
        resto_500   = overpass_fetch(t_lat, t_lon,  500, '["amenity"~"restaurant|cafe|fast_food"]')
        hotel_500   = overpass_fetch(t_lat, t_lon,  500, '["tourism"="hotel"]')
        bank_300    = overpass_fetch(t_lat, t_lon,  300, '["amenity"~"bank|atm"]')
        sport_1k    = overpass_fetch(t_lat, t_lon, 1000, '["leisure"~"fitness_centre|sports_centre"]')
        creche_1k   = overpass_fetch(t_lat, t_lon, 1000, '["amenity"~"childcare|kindergarten"]')
        supermarche = overpass_fetch(t_lat, t_lon,  500, '["shop"~"supermarket|convenience"]')
        pharmacy    = overpass_fetch(t_lat, t_lon,  500, '["amenity"="pharmacy"]')
        offices_500 = overpass_fetch(t_lat, t_lon,  500, '["office"]')

    # Transports
    st.markdown("### 🚌 Transports")
    nt, dt = nearest(tram_els,  t_lat, t_lon)
    nb, db = nearest(bus_els,   t_lat, t_lon)
    ng, dg = nearest(train_els, t_lat, t_lon)
    nv, dv = nearest(bike_els,  t_lat, t_lon)

    tc1, tc2, tc3, tc4 = st.columns(4)
    tc1.metric("🚋 Tram", f"{int(dt)} m" if nt else "—", f"{wmin(dt)} min" if nt else None)
    tc2.metric("🚌 Bus",  f"{int(db)} m" if nb else "—", f"{wmin(db)} min" if nb else None)
    tc3.metric("🚂 Gare", f"{dg/1000:.1f} km" if ng else "—", f"{wmin(dg)} min" if ng else None)
    tc4.metric("🚲 Vélos", f"{int(dv)} m" if nv else "—", f"{wmin(dv)} min" if nv else None)

    # Services
    st.markdown("### 🍽️ Services à proximité")
    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("Restos/cafés 300 m", len(resto_300))
    sc2.metric("Restos/cafés 500 m", len(resto_500))
    sc3.metric("Hôtels 500 m", len(hotel_500))
    sc4.metric("Banques/ATM 300 m", len(bank_300))
    sc5, sc6, sc7, sc8 = st.columns(4)
    sc5.metric("Salles de sport 1 km", len(sport_1k))
    sc6.metric("Crèches 1 km", len(creche_1k))
    sc7.metric("Supermarchés 500 m", len(supermarche))
    sc8.metric("Pharmacies 500 m", len(pharmacy))

    # Environnement tertiaire
    st.markdown("### 🏢 Environnement tertiaire")
    st.metric("Immeubles de bureaux / locaux 500 m", len(offices_500))

    # Score
    st.markdown("### 🏅 Score synthétique")
    sc_t = min(
        (40 if nt and dt<=300 else 25 if nt and dt<=600 else 0) +
        (20 if nb and db<=200 else 10 if nb and db<=400 else 0) +
        (25 if ng and dg<=800 else 15 if ng and dg<=1500 else 0) +
        (15 if nv and dv<=300 else 0), 100)
    sc_s = min(len(resto_300)*5 + len(hotel_500)*10 + len(bank_300)*5 +
               len(sport_1k)*8 + len(creche_1k)*5, 100)
    sc_b = min(len(offices_500)*5 + len(resto_500)*2, 100)
    sc_g = round(sc_t*0.4 + sc_s*0.35 + sc_b*0.25)
    note = ("⭐⭐⭐⭐⭐ Excellent" if sc_g>=80 else "⭐⭐⭐⭐ Très bon" if sc_g>=65
            else "⭐⭐⭐ Bon" if sc_g>=50 else "⭐⭐ Moyen" if sc_g>=35 else "⭐ Faible")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Transports /100", sc_t)
    c2.metric("Services /100", sc_s)
    c3.metric("Tertiaire /100", sc_b)
    c4.metric("Score global /100", sc_g, note)

    # Carte environnement
    st.markdown("### 🗺️ Carte des commodités")
    em = folium.Map(location=[t_lat, t_lon], zoom_start=15, tiles=None)
    folium.TileLayer(
        "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        attr="Esri", name="🛰️ Satellite", overlay=False, control=True).add_to(em)
    folium.TileLayer(
        "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png",
        attr="CartoDB", name="🗺️ Plan", overlay=False, control=True).add_to(em)
    folium.LayerControl(position="topright", collapsed=False).add_to(em)
    folium.Marker([t_lat, t_lon], tooltip="🏢 Actif",
                  icon=folium.Icon(color="red", icon="home", prefix="fa")).add_to(em)

    LAYERS = [
        ("blue",      tram_els[:15],    "🚋 Tram"),
        ("cadetblue", bus_els[:30],     "🚌 Bus"),
        ("orange",    resto_500[:40],   "🍽️ Resto"),
        ("purple",    hotel_500[:10],   "🏨 Hôtel"),
        ("green",     sport_1k[:10],    "🏋️ Sport"),
        ("darkgreen", bank_300[:10],    "🏦 Banque"),
    ]
    for color, els, label in LAYERS:
        for e in els:
            elat = e.get("lat") or (e.get("center") or {}).get("lat")
            elon = e.get("lon") or (e.get("center") or {}).get("lon")
            name = e.get("tags", {}).get("name", label)
            if elat and elon:
                folium.CircleMarker([elat, elon], radius=7,
                    color="white", weight=1,
                    fill=True, fill_color=color, fill_opacity=0.85,
                    tooltip=name).add_to(em)

    components.html(em._repr_html_(), height=520, scrolling=False)
    st.caption("🔴 Actif · 🔵 Tram · 🟦 Bus · 🟠 Restos · 🟣 Hôtels · 🟢 Sport · 💚 Banques")

