"""
Module de collecte en direct — BureauxLocaux & Geolocaux
----------------------------------------------------------------
ATTENTION — lire avant usage :
- Ce module n'a pas pu être testé contre les sites réels (sandbox sans accès
  à ces domaines). La fonction diagnose() te dira immédiatement si l'extraction
  fonctionne ou si le site sert du contenu rendu en JavaScript (auquel cas
  requests/BeautifulSoup ne suffit pas — il faudrait Playwright).
- L'extraction est basée sur des regex sur le texte visible, pas sur des
  sélecteurs CSS précis (je n'ai pas pu inspecter le DOM réel). Fragile par
  nature : si ça casse, c'est probablement la mise en page du site qui a changé.
- Usage à tes propres risques vis-à-vis des CGU des sites concernés.
"""

import re
import time
import unicodedata
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Cache-Control": "max-age=0",
}

SURFACE_RE = re.compile(r"(\d[\d\s]{0,5})\s?m\s?2|(\d[\d\s]{0,5})\s?m²", re.I)

# ── Loyer : ordre de priorité décroissante ──────────────────────────────────
# Séparateur entre HT et HC : tiret, slash, espace ou rien → "HT-HC", "HT/HC", "HTHC", "HT HC"
# €  optionnel car le tableau des étages affiche "350 HT/HC/m²/an" sans le symbole €.
# 1. Format explicite "340 € HT-HC/m²/an" ou "350 HT/HC/m²/an" — le plus fiable
LOYER_HTHC_M2_AN_RE = re.compile(
    r"(\d[\d\s]{1,5})\s?€?\s?HT\s?[-/ ]?\s?HC\s?/\s?m²\s?/\s?an", re.I)
# 2. Format "180 € HT /m²/an" (HT seul, sans HC)
LOYER_ANNUEL_M2_RE = re.compile(
    r"(\d[\d\s]{1,5})\s?€?\s?(?:HT)?\s?/\s?m²\s?/\s?an", re.I)
# 3. Format mensuel explicitement HT-HC "9 152 € HT-HC/mois" → calculé en /m²/an
LOYER_HTHC_MENSUEL_RE = re.compile(
    r"(\d[\d\s]{1,7})\s?€?\s?HT\s?[-/ ]?\s?HC\s?/\s?mois", re.I)
# 4. Format mensuel générique "4 500 €/mois" (peut être HT ou TTC)
LOYER_MENSUEL_RE = re.compile(
    r"(\d[\d\s]{1,7})\s?€?\s?(?:HT|TTC)?\s?(?:/|par)\s?mois", re.I)

PRIX_TOTAL_RE = re.compile(r"prix\s*:?\s*(\d[\d\s]{3,9})\s?€", re.I)

# ── Charges ──────────────────────────────────────────────────────────────────
CHARGES_M2_RE = re.compile(
    r"charges?\s*:?\s*(?:d'environ\s*)?(\d[\d\s]{0,4})\s?€\s?(?:HT)?\s?/\s?m²\s?/\s?an", re.I)
# "Charges locatives : 6 783 € HT/an" (format BureauxLocaux fiche détail)
CHARGES_LOCATIVES_AN_RE = re.compile(
    r"charges?\s+locatives?\s*:?\s*(\d[\d\s]{0,6})\s?€\s?(?:HT)?\s?/?\s?an", re.I)
CHARGES_MENSUEL_RE = re.compile(
    r"(?:provisions?\s+pour\s+)?charges?\s*:?\s*(\d[\d\s]{0,6})\s?€?\s?(?:HT|TTC)?\s?(?:/|par)\s?mois", re.I)
CHARGES_ANNUEL_RE = re.compile(
    r"(?:provisions?\s+pour\s+)?charges?\s*:?\s*(\d[\d\s]{0,6})\s?€?\s?(?:HT|TTC)?\s?(?:/|par)\s?an", re.I)

# "charges incluses" / "tout compris" → loyer affiché = déjà HT/HC
ALL_INCLUSIVE_RE = re.compile(
    r"charges?\s+(?:incluses?|comprises?)|tout\s+(?:est\s+)?compris|loyer\s+(?:tout\s+)?charges?\s+comprises?",
    re.I,
)
# Charges mentionnées sans montant → candidat au deep-fetch de la fiche complète
CHARGES_MENTIONED_NO_AMOUNT_RE = re.compile(r"\b(?:provisions?\s+pour\s+)?charges?\b", re.I)


# Bornes de plausibilité (marché tertiaire France, large volontairement) : toute valeur
# extraite en dehors est considérée comme une erreur d'extraction, pas une vraie donnée.
LOYER_M2_AN_MIN, LOYER_M2_AN_MAX = 30, 600       # €/m²/an HT/HC
PRIX_M2_MIN, PRIX_M2_MAX = 200, 8000             # €/m² (vente)

# Communes de Nantes Métropole + périphérie immédiate (périmètre de la prompt d'origine).
# Triées par longueur décroissante pour que le regex matche le nom le plus précis en premier
# (ex. "Saint-Sébastien-sur-Loire" avant "Sébastien").
COMMUNES_PERIMETRE = [
    "Saint-Sébastien-sur-Loire", "La Chapelle-sur-Erdre", "Saint-Etienne-de-Montluc",
    "Vigneux-de-Bretagne", "Grandchamp-des-Fontaines", "Les Sorinières", "Sainte-Luce-sur-Loire",
    "Thouaré-sur-Loire", "Saint-Herblain", "Haute-Goulaine", "Basse-Goulaine", "Bouguenais",
    "Carquefou", "Treillières", "Couëron", "Sautron", "Orvault", "Vertou", "Indre", "Rezé",
    "Nantes",
]
_commune_alt = "|".join(re.escape(c) for c in COMMUNES_PERIMETRE)
COMMUNE_RE = re.compile(r"\b(" + _commune_alt + r")\b")

VOIE_TYPES = (r"Rue|Boulevard|Bd|Avenue|Av|All[ée]e|Quai|Place|Mail|Chemin|Route|Impasse|"
              r"Cours|Square|Esplanade|Cit[ée]|Passage|Rond-?point|Voie")

# Adresse — tentative 1 (la plus fiable) : numéro + type de voie + nom, code postal, commune connue
ADDRESS_RE = re.compile(
    r"(\d+[\d\-\s]*\s+(?:" + VOIE_TYPES + r")\s+[^,\d]{2,55}?)\s*,?\s*(\d{5})\s+(" + _commune_alt + r")\b",
    re.I,
)
# Adresse — tentative 2 (repli) : pas de numéro de rue, mais type de voie + nom + CP + commune
ADDRESS_NO_NUM_RE = re.compile(
    r"((?:" + VOIE_TYPES + r")\s+[^,\d]{2,55}?)\s*,?\s*(\d{5})\s+(" + _commune_alt + r")\b",
    re.I,
)
# Adresse — tentative 3 (dernier repli) : juste CP + commune connue, sans nom de voie
CP_COMMUNE_RE = re.compile(r"\b(\d{5})\s+(" + _commune_alt + r")\b", re.I)

OP_TYPE_RE = re.compile(
    r"\b(Vente|Location)\s+(Bureaux?|Entrep[oô]ts?|Locaux?\s+d'activit[ée]s?|"
    r"Locaux?\s+commerciaux?|Commerces?)\b", re.I,
)

ETAT_KEYWORDS = [
    # ── NEUF ──────────────────────────────────────────────────────────────────
    # Formulations directes
    (re.compile(r"\bimmeuble\s+neuf\b", re.I),                  "Neuf"),
    (re.compile(r"\bsurfaces?\s+neuves?\b", re.I),              "Neuf"),
    (re.compile(r"\bbureaux?\s+neufs?\b", re.I),                "Neuf"),
    (re.compile(r"\blocaux?\s+neufs?\b", re.I),                 "Neuf"),
    (re.compile(r"\bbâtiment\s+neuf\b", re.I),                  "Neuf"),
    (re.compile(r"\bprogramme\s+neuf\b", re.I),                 "Neuf"),
    (re.compile(r"\bconstruction\s+neuve?\b", re.I),            "Neuf"),
    (re.compile(r"\blivr[ée]\s+(?:en\s+)?(?:20[1-9]\d|neuf)\b", re.I), "Neuf"),
    (re.compile(r"\blivraison\s+(?:20[1-9]\d|prévue)\b", re.I),"Neuf"),
    (re.compile(r"\b(?:R[+]\d+\s+)?neuf\b", re.I),             "Neuf"),
    # ── RESTRUCTURÉ ───────────────────────────────────────────────────────────
    (re.compile(r"enti[èe]rement\s+r[ée]nov[ée]s?", re.I),     "Restructuré"),
    (re.compile(r"enti[èe]rement\s+r[ée]habilit[ée]s?", re.I), "Restructuré"),
    (re.compile(r"enti[èe]rement\s+r[ée]am[ée]nag[ée]s?", re.I),"Restructuré"),
    (re.compile(r"r[ée]novation\s+compl[èe]te?", re.I),        "Restructuré"),
    (re.compile(r"enti[èe]rement\s+r[ée]structur[ée]s?", re.I),"Restructuré"),
    (re.compile(r"plateau\s+r[ée]nov[ée]", re.I),              "Restructuré"),
    (re.compile(r"locaux?\s+r[ée]nov[ée]s?", re.I),            "Restructuré"),
    (re.compile(r"bureaux?\s+r[ée]nov[ée]s?", re.I),           "Restructuré"),
    (re.compile(r"r[ée]habilit[ée]s?", re.I),                  "Restructuré"),
    (re.compile(r"r[ée]structur[ée]s?", re.I),                 "Restructuré"),
    (re.compile(r"r[ée]nov[ée]s?\s+(?:en\s+)?20[1-9]\d", re.I),"Restructuré"),
    (re.compile(r"travaux?\s+(?:de\s+)?r[ée]novation", re.I),  "Restructuré"),
    (re.compile(r"r[ée]nov[ée]s?\s+(?:et\s+)?(?:am[ée]nag[ée]s?|[ée]quip[ée]s?)", re.I), "Restructuré"),
    (re.compile(r"ancien\s+appartement\s+(?:bourg[eo]+is\s+)?(?:enti[èe]rement\s+)?r[ée]nov[ée]", re.I), "Restructuré"),
    (re.compile(r"\br[ée]nov[ée]s?\b", re.I),                  "Restructuré"),  # générique, en dernier
    # ── SECONDE MAIN ──────────────────────────────────────────────────────────
    (re.compile(r"\b[àa]\s+r[ée]nover\b", re.I),               "Seconde main"),
    (re.compile(r"travaux?\s+[àa]\s+pr[ée]voir", re.I),        "Seconde main"),
    (re.compile(r"rafra[iî]chissement\s+[àa]\s+pr[ée]voir", re.I), "Seconde main"),
    (re.compile(r"[ée]tat\s+d'usage", re.I),                   "Seconde main"),
    (re.compile(r"bon\s+[ée]tat\s+g[ée]n[ée]ral", re.I),      "Seconde main"),
    (re.compile(r"en\s+bon\s+[ée]tat", re.I),                  "Seconde main"),
    (re.compile(r"existant\s+cloisonn[ée]", re.I),             "Seconde main"),
]

TYPE_ACTIF_MAP = {
    "bureau": "Bureaux", "bureaux": "Bureaux",
    "entrepot": "Activités", "entrepôt": "Activités", "entrepots": "Activités", "entrepôts": "Activités",
    "commerce": "Commerce", "commerces": "Commerce",
}

PARKING_RE = re.compile(r"(\d{1,3})\s*places?\s*(?:de\s*)?(?:parking|stationnement)", re.I)
ASCENSEUR_RE = re.compile(r"\bascenseur\b", re.I)
ASCENSEUR_NEG_RE = re.compile(r"sans\s+ascenseur", re.I)
CLIM_RE = re.compile(r"\bclimatis", re.I)
FIBRE_RE = re.compile(r"\bfibre\b", re.I)
ERP_RE = re.compile(r"\bERP\b")
PMR_RE = re.compile(r"\bPMR\b|accessible.{0,15}handicap", re.I)
CERTIF_RE = re.compile(r"\b(HQE|BREEAM|LEED)\b")
DIVISIBLE_RE = re.compile(r"divisible.{0,20}(?:à partir de|dès)\s*(\d[\d\s]{0,5})\s?m", re.I)


def slugify(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-zA-Z0-9]+", "-", s.lower()).strip("-")
    return s


def fetch(url: str, timeout: int = 12) -> str | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout)
        r.raise_for_status()
        return r.text
    except requests.RequestException as e:
        print(f"[live_search] échec requête {url} : {e}")
        return None


def looks_like_js_shell(html: str) -> bool:
    """Heuristique : si très peu de texte visible par rapport à la taille de la page,
    le contenu est probablement injecté en JavaScript après coup."""
    soup = BeautifulSoup(html, "html.parser")
    visible_text = soup.get_text(strip=True)
    return len(visible_text) < 1500


def _clean_int(s: str) -> float | None:
    """Convertit une chaîne numérique en float.
    Gère tous les séparateurs de milliers français :
    espace normale, \xa0 (insécable), \u202f (fine insécable), etc."""
    if not s:
        return None
    try:
        import unicodedata
        # supprime tout caractère Unicode classé comme espace (catégorie Z*)
        cleaned = "".join(c for c in str(s) if not unicodedata.category(c).startswith("Z"))
        cleaned = cleaned.replace(",", ".")
        return float(cleaned)
    except (ValueError, TypeError):
        return None


def _extract_href(tag, base_url: str, search_url: str) -> str | None:
    """Cherche le lien le plus probable vers la fiche détaillée de l'annonce.
    Rejette : ancres vides, pagination, liens identiques à l'URL de recherche elle-même
    (= lien de catégorie, pas une fiche précise), liens trop courts pour être une fiche."""
    from urllib.parse import urljoin, urlparse
    search_path = urlparse(search_url).path.rstrip("/")
    candidates = []
    for a in tag.find_all("a", href=True):
        href = a["href"]
        if not href or href in ("#", "/") or href.startswith("javascript:") or href.startswith("mailto:"):
            continue
        if re.search(r"page/\d+|\?page=", href, re.I):
            continue
        full = urljoin(base_url, href)
        path = urlparse(full).path.rstrip("/")
        if path == search_path:
            continue  # lien de catégorie identique à la recherche : pas une fiche précise
        candidates.append(full)
    if not candidates:
        return None
    # préfère le lien le plus spécifique (contient un identifiant numérique, ou le plus long chemin)
    candidates.sort(key=lambda u: (bool(re.search(r"\d{3,}", u)), len(u)), reverse=True)
    return candidates[0]


def _is_generic_link(href: str | None, search_url: str) -> bool:
    return href is None or href.rstrip("/") == search_url.rstrip("/")


def _extract_photo(tag, base_url: str) -> str | None:
    """Extrait l'URL de la première photo de l'annonce.
    Cherche : data-src (lazy-load), src, srcset, ou background-image CSS."""
    from urllib.parse import urljoin
    # 1. balises img : priorité data-src (lazy-loading), puis src
    for img in tag.find_all("img"):
        for attr in ("data-src", "data-lazy-src", "data-original", "src"):
            src = img.get(attr, "")
            if not src:
                continue
            if any(x in src.lower() for x in ["logo", "placeholder", "blank", "pixel", "spacer"]):
                continue
            if any(ext in src.lower() for ext in [".jpg", ".jpeg", ".png", ".webp", ".gif"]):
                return urljoin(base_url, src) if not src.startswith("http") else src
        # srcset : prendre la première URL
        srcset = img.get("srcset", "")
        if srcset:
            first_url = srcset.split(",")[0].strip().split(" ")[0]
            if first_url.startswith("http") and "logo" not in first_url.lower():
                return first_url
    # 2. divs avec background-image en style inline
    for div in tag.find_all(style=True):
        style = div.get("style", "")
        bg_m = re.search(r"background-image\s*:\s*url\(['\"]?([^'\")\s]+)['\"]?\)", style, re.I)
        if bg_m:
            url = bg_m.group(1)
            if url.startswith("http") and "logo" not in url.lower():
                return url
    return None


def _split_listing_blocks(soup: BeautifulSoup, base_url: str) -> list[tuple[str, str | None, str | None]]:
    """Retourne une liste de (texte, lien_direct, photo_url) — un par annonce détectée."""
    candidates = soup.find_all(["article", "li", "div"],
                                class_=re.compile(r"annonce|listing|card|result|offre", re.I))
    raw = []
    for c in candidates:
        txt = c.get_text(" ", strip=True)
        if len(txt) > 60 and SURFACE_RE.search(txt):
            raw.append((txt, _extract_href(c, base_url, base_url), _extract_photo(c, base_url)))

    if not raw:
        full_text = soup.get_text(" ", strip=True)
        parts = re.split(r"(?=Voir l'annonce|En savoir plus|Disponibilité\s?:)", full_text)
        raw = [(p, None, None) for p in parts if SURFACE_RE.search(p)]

    raw.sort(key=lambda x: len(x[0]))
    kept: list[tuple[str, str | None, str | None]] = []
    for txt, href, photo in raw:
        if any(txt in other_txt for other_txt, _, __ in kept):
            continue
        kept.append((txt, href, photo))
    return kept


def _detect_type_actif(text: str, fallback_type_bien: str) -> str:
    m = re.search(r"\b(bureaux?|entrep[oô]ts?|locaux?\s+d'activit[ée]s?|locaux?\s+commerciaux?|commerces?)\b",
                   text, re.I)
    if m:
        key = m.group(1).lower()
        if "bureau" in key:
            return "Bureaux"
        if "entrep" in key or "activit" in key:
            return "Activités"
        if "commerc" in key:
            return "Commerce"
    return TYPE_ACTIF_MAP.get(fallback_type_bien.lower(), "Non disponible")


def _detect_etat(text: str) -> str:
    for pattern, label in ETAT_KEYWORDS:
        if pattern.search(text):
            return label
    # Convention : si l'annonce ne mentionne pas explicitement "neuf" ou "rénové",
    # le bien est de seconde main (les biens neufs s'annoncent toujours comme tels).
    return "Seconde main"


def parse_listing_block(block: str, source_name: str, source_url: str,
                         fallback_commune: str = "", fallback_cp: str = "",
                         fallback_type_bien: str = "") -> dict | None:
    surf_m = SURFACE_RE.search(block)
    if not surf_m:
        return None
    surface = _clean_int(surf_m.group(1) or surf_m.group(2))
    if not surface:
        return None

    prix_total = PRIX_TOTAL_RE.search(block)

    # ── Loyer : ordre de priorité strict ──────────────────────────────────────
    loyer_annuel_m2 = None
    loyer_is_hthc = False

    # 1. "167 € HT-HC/m²/an" — format le plus direct, déjà HT-HC
    m = LOYER_HTHC_M2_AN_RE.search(block)
    if m:
        v = _clean_int(m.group(1))
        if v and LOYER_M2_AN_MIN <= v <= LOYER_M2_AN_MAX:
            loyer_annuel_m2 = v
            loyer_is_hthc = True

    # 2. "180 € HT /m²/an" — annuel par m², non HT-HC (charges à déduire séparément)
    if loyer_annuel_m2 is None:
        m = LOYER_ANNUEL_M2_RE.search(block)
        if m:
            v = _clean_int(m.group(1))
            if v and LOYER_M2_AN_MIN <= v <= LOYER_M2_AN_MAX:
                loyer_annuel_m2 = v

    # 3. "4 500 € HT-HC/mois" — mensuel HT-HC → /m²/an par calcul
    if loyer_annuel_m2 is None:
        m = LOYER_HTHC_MENSUEL_RE.search(block)
        if m and surface:
            v = _clean_int(m.group(1))
            if v:
                calc = round(v * 12 / surface, 1)
                if LOYER_M2_AN_MIN <= calc <= LOYER_M2_AN_MAX:
                    loyer_annuel_m2 = calc
                    loyer_is_hthc = True

    # 4. "4 500 €/mois" — mensuel générique → /m²/an par calcul
    # Sur BureauxLocaux, le prix affiché en liste est toujours HT-HC (c'est la convention du site)
    if loyer_annuel_m2 is None:
        m = LOYER_MENSUEL_RE.search(block)
        if m and surface:
            v = _clean_int(m.group(1))
            if v:
                calc = round(v * 12 / surface, 1)
                if LOYER_M2_AN_MIN <= calc <= LOYER_M2_AN_MAX:
                    loyer_annuel_m2 = calc
                    loyer_is_hthc = True  # prix affiché en annonce = HT-HC par convention

    # ── Charges : ordre de priorité ───────────────────────────────────────────
    charges_m2_an = None
    charges_known = False

    # "Charges locatives : 6 783 € HT/an" (format fiche BureauxLocaux)
    m = CHARGES_LOCATIVES_AN_RE.search(block)
    if m and surface:
        v = _clean_int(m.group(1))
        if v:
            c = round(v / surface, 1)
            if 0 < c <= 150:
                charges_m2_an = c
                charges_known = True

    if not charges_known:
        m = CHARGES_M2_RE.search(block)
        if m:
            v = _clean_int(m.group(1))
            if v and 0 < v <= 150:
                charges_m2_an = v
                charges_known = True

    if not charges_known:
        m = CHARGES_MENSUEL_RE.search(block)
        if m and surface:
            v = _clean_int(m.group(1))
            if v:
                c = round(v * 12 / surface, 1)
                if 0 < c <= 150:
                    charges_m2_an = c
                    charges_known = True

    if not charges_known:
        m = CHARGES_ANNUEL_RE.search(block)
        if m and surface:
            v = _clean_int(m.group(1))
            if v:
                c = round(v / surface, 1)
                if 0 < c <= 150:
                    charges_m2_an = c
                    charges_known = True

    all_inclusive = bool(ALL_INCLUSIVE_RE.search(block))
    charges_mentioned_no_amount = (not charges_known and not all_inclusive
                                    and bool(CHARGES_MENTIONED_NO_AMOUNT_RE.search(block)))

    # ── Loyer HT-HC final ─────────────────────────────────────────────────────
    if loyer_is_hthc:
        loyer_ht_hc_m2_an = loyer_annuel_m2
    elif all_inclusive and loyer_annuel_m2 is not None:
        loyer_ht_hc_m2_an = loyer_annuel_m2
        charges_m2_an = 0.0
    elif loyer_annuel_m2 is not None and charges_m2_an is not None:
        loyer_ht_hc_m2_an = round(loyer_annuel_m2 + charges_m2_an, 1)
    else:
        loyer_ht_hc_m2_an = None

    prix_total_val = _clean_int(prix_total.group(1)) if prix_total else None
    prix_m2 = round(prix_total_val / surface, 1) if (prix_total_val and surface) else None
    if prix_m2 is not None and not (PRIX_M2_MIN <= prix_m2 <= PRIX_M2_MAX):
        prix_m2 = None
        prix_total_val = None

    op_type_m = OP_TYPE_RE.search(block)
    operation = "Non disponible"
    if op_type_m:
        operation = "Vente" if op_type_m.group(1).lower() == "vente" else "Location"
    elif prix_total_val and not loyer_annuel_m2:
        operation = "Vente"
    elif loyer_annuel_m2:
        operation = "Location"

    type_actif = _detect_type_actif(block, fallback_type_bien)

    # adresse : 3 tentatives en cascade, du plus précis au plus large — on ne lâche
    # "Non disponible" qu'en tout dernier recours.
    addr_m = ADDRESS_RE.search(block)
    if addr_m:
        adresse = addr_m.group(1).strip().rstrip(",")
        code_postal, commune = addr_m.group(2), addr_m.group(3)
        adresse_precise = True
    else:
        addr_m2 = ADDRESS_NO_NUM_RE.search(block)
        if addr_m2:
            adresse = addr_m2.group(1).strip().rstrip(",")
            code_postal, commune = addr_m2.group(2), addr_m2.group(3)
            adresse_precise = True
        else:
            cp_com_m = CP_COMMUNE_RE.search(block)
            if cp_com_m:
                code_postal, commune = cp_com_m.group(1), cp_com_m.group(2)
                adresse = f"{commune} (adresse précise non extraite)"
            else:
                com_m = COMMUNE_RE.search(block)
                commune = com_m.group(1) if com_m else (fallback_commune or "Non disponible")
                code_postal = fallback_cp or "Non disponible"
                adresse = f"{commune} (adresse précise non extraite)" if com_m else "Non disponible"
            adresse_precise = False

    etat = _detect_etat(block)

    nb_parkings = None
    pk_m = PARKING_RE.search(block)
    if pk_m:
        nb_parkings = _clean_int(pk_m.group(1))

    ascenseur = "Non disponible"
    if ASCENSEUR_NEG_RE.search(block):
        ascenseur = "Non"
    elif ASCENSEUR_RE.search(block):
        ascenseur = "Oui"

    climatisation = "Oui" if CLIM_RE.search(block) else "Non disponible"
    fibre = "Oui" if FIBRE_RE.search(block) else "Non disponible"
    erp = "Oui" if ERP_RE.search(block) else "Non disponible"
    pmr = "Oui" if PMR_RE.search(block) else "Non disponible"
    certif_m = CERTIF_RE.search(block)
    certif_env = certif_m.group(1).upper() if certif_m else "Non disponible"
    div_m = DIVISIBLE_RE.search(block)
    surface_div_min = _clean_int(div_m.group(1)) if div_m else None

    has_price = loyer_annuel_m2 is not None or prix_m2 is not None
    # "géocodable" = on a au moins une commune identifiable → on peut placer le point sur la carte
    # (même sans numéro de rue, le centroïde de la commune suffit pour le calcul de distance)
    has_location = commune not in ("Non disponible", "", None)
    is_complete = has_location and has_price

    return {
        "Type_actif": type_actif,
        "Operation": operation,
        "Etat": etat,
        "Adresse": adresse,
        "Commune": commune,
        "Code_postal": code_postal,
        "Surface_m2": surface,
        "Surface_div_min_m2": surface_div_min,
        "Nb_parkings": nb_parkings,
        "Ascenseur": ascenseur,
        "Climatisation": climatisation,
        "Fibre": fibre,
        "ERP": erp,
        "PMR": pmr,
        "Certif_env": certif_env,
        "Loyer_facial_eur_m2_an": loyer_annuel_m2,
        "Charges_eur_m2_an": charges_m2_an,
        "Loyer_HT_HC_eur_m2_an": loyer_ht_hc_m2_an,
        "Prix_vente_total_eur": prix_total_val,
        "Prix_vente_eur_m2": prix_m2,
        "Date_transaction": "Offre en cours",
        "Source": source_name,
        "Lien": source_url,
        "Fiabilite": 2,
        "Adresse_precise": adresse_precise,
        "Complete": is_complete,
        "Needs_detail_fetch": charges_mentioned_no_amount,
        "Observations": block[:500],
    }


def filter_complete(items: list[dict]) -> list[dict]:
    """Ne garde que les annonces avec adresse précise ET au moins un prix exploitable
    (loyer ou prix de vente plausible)."""
    return [it for it in items if it.get("Complete")]


def _extract_photo_from_detail(soup: BeautifulSoup, base_url: str = "") -> str | None:
    """Extrait l'URL de la photo principale depuis une fiche détaillée.
    Ordre de priorité :
    1. og:image (méta HTML5, présente sur quasi tous les sites immo)
    2. twitter:image
    3. JSON-LD schema.org (image déclarée dans les données structurées)
    4. Première <img> de taille significative dans la page
    """
    import json

    # 1. og:image
    tag = soup.find("meta", property="og:image") or \
          soup.find("meta", attrs={"name": "og:image"})
    if tag and tag.get("content"):
        url = tag["content"].strip()
        if url.startswith("http") and "logo" not in url.lower():
            return url

    # 2. twitter:image
    tag = soup.find("meta", attrs={"name": "twitter:image"})
    if tag and tag.get("content"):
        url = tag["content"].strip()
        if url.startswith("http") and "logo" not in url.lower():
            return url

    # 3. JSON-LD
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
            if isinstance(data, dict):
                for key in ("image", "photo"):
                    val = data.get(key)
                    if isinstance(val, str) and val.startswith("http"):
                        return val
                    if isinstance(val, list) and val and isinstance(val[0], str):
                        return val[0]
                    if isinstance(val, dict) and val.get("url", "").startswith("http"):
                        return val["url"]
        except Exception:
            pass

    # 4. Première grande image dans la page
    from urllib.parse import urljoin
    for img in soup.find_all("img"):
        for attr in ("data-src", "data-lazy-src", "src"):
            src = img.get(attr, "")
            if not src or "logo" in src.lower() or "placeholder" in src.lower():
                continue
            # exclure les icônes (souvent < 50px)
            try:
                if int(img.get("width", "999")) < 100:
                    continue
            except (ValueError, TypeError):
                pass
            full = urljoin(base_url, src) if not src.startswith("http") else src
            if any(ext in full.lower() for ext in [".jpg", ".jpeg", ".png", ".webp"]):
                return full
    return None


def enrich_listing(item: dict, timeout: int = 12) -> dict:
    """Va chercher la fiche complète de l'annonce pour compléter :
    - la photo principale (og:image)
    - les charges si inconnues
    - l'adresse précise si manquante
    Ne modifie jamais une valeur déjà présente."""
    lien = item.get("Lien", "")
    if not lien or not lien.startswith("http"):
        return item
    if not item.get("Lien_direct", True):
        return item  # lien générique de recherche, pas une fiche individuelle

    html = fetch(lien, timeout=timeout)
    if not html:
        return item
    soup = BeautifulSoup(html, "html.parser")
    full_text = soup.get_text(" ", strip=True)

    updated = dict(item)
    surface = item.get("Surface_m2")

    # ── Photo principale ─────────────────────────────────────────────────
    if not updated.get("Photo_url"):
        photo = _extract_photo_from_detail(soup, lien)
        if photo:
            updated["Photo_url"] = photo

    # ── Charges / loyer HT-HC ───────────────────────────────────────────
    if updated.get("Charges_eur_m2_an") is None and updated.get("Loyer_HT_HC_eur_m2_an") is None:
        if ALL_INCLUSIVE_RE.search(full_text) and updated.get("Loyer_facial_eur_m2_an") is not None:
            updated["Loyer_HT_HC_eur_m2_an"] = updated["Loyer_facial_eur_m2_an"]
            updated["Charges_eur_m2_an"] = 0.0
        else:
            for pattern, divisor in [
                (CHARGES_LOCATIVES_AN_RE, surface),
                (CHARGES_M2_RE, None),
                (CHARGES_MENSUEL_RE, surface / 12 if surface else None),
                (CHARGES_ANNUEL_RE, surface),
            ]:
                m = pattern.search(full_text)
                if m:
                    raw = _clean_int(m.group(1))
                    if raw and divisor:
                        c = round(raw / divisor, 1)
                    elif raw:
                        c = raw
                    else:
                        continue
                    if 0 < c <= 150:
                        updated["Charges_eur_m2_an"] = c
                        if updated.get("Loyer_facial_eur_m2_an") is not None:
                            updated["Loyer_HT_HC_eur_m2_an"] = round(
                                updated["Loyer_facial_eur_m2_an"] + c, 1)
                        break

    # ── Adresse précise ─────────────────────────────────────────────────
    if not updated.get("Adresse_precise"):
        addr_m = ADDRESS_RE.search(full_text) or ADDRESS_NO_NUM_RE.search(full_text)
        if addr_m:
            updated["Adresse"] = addr_m.group(1).strip().rstrip(",")
            updated["Code_postal"] = addr_m.group(2)
            updated["Commune"] = addr_m.group(3)
            updated["Adresse_precise"] = True
            updated["Complete"] = True

    updated["Needs_detail_fetch"] = False
    return updated


def _fetch_one_photo(it: dict) -> tuple:
    """Fetch la photo d'une seule fiche. Retourne (id(it), dict_mis_à_jour)."""
    updated = dict(it)
    try:
        html = fetch(it["Lien"], timeout=10)
        if html:
            soup = BeautifulSoup(html, "html.parser")
            photo = _extract_photo_from_detail(soup, it["Lien"])
            if photo:
                updated["Photo_url"] = photo
    except Exception:
        pass
    return id(it), updated


def fetch_photos_batch(items: list[dict], max_items: int = 40,
                       progress_cb=None, parallel: bool = False,
                       workers: int = 5) -> list[dict]:
    """Fetche les fiches individuelles pour récupérer les photos manquantes.
    Limité à max_items requêtes pour ne pas surcharger les serveurs.

    parallel=False (Sûr)  : séquentiel, 150ms entre requêtes, aucun risque.
    parallel=True  (Rapide): jusqu'à `workers` requêtes simultanées (~3-4× plus
                             rapide), léger risque de rate-limit sur les serveurs.
    """
    to_fetch = [it for it in items
                if not it.get("Photo_url")
                and it.get("Lien", "").startswith("http")
                and it.get("Lien_direct", True)][:max_items]

    result_map = {}

    if parallel and len(to_fetch) > 1:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        done = 0
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futures = {ex.submit(_fetch_one_photo, it): it for it in to_fetch}
            for fut in as_completed(futures):
                try:
                    key, updated = fut.result()
                    result_map[key] = updated
                except Exception:
                    pass
                done += 1
                if progress_cb:
                    progress_cb(done, len(to_fetch))
    else:
        for i, it in enumerate(to_fetch):
            key, updated = _fetch_one_photo(it)
            result_map[key] = updated
            if progress_cb:
                progress_cb(i + 1, len(to_fetch))
            time.sleep(0.15)  # courtoisie : 150ms entre requêtes

    return [result_map.get(id(it), it) for it in items]


def enrich_listings(items: list[dict], max_items: int = 30, progress_cb=None) -> list[dict]:
    """Enrichit charges + adresse + photo pour les annonces qui en ont besoin."""
    to_enrich = [it for it in items
                 if it.get("Needs_detail_fetch")
                 or not it.get("Adresse_precise")
                 or not it.get("Photo_url")][:max_items]
    enrich_map = {}
    for i, it in enumerate(to_enrich):
        enrich_map[id(it)] = enrich_listing(it)
        if progress_cb:
            progress_cb(i + 1, len(to_enrich))
        time.sleep(0.15)
    return [enrich_map.get(id(it), it) for it in items]


def dedupe_listings(items: list[dict]) -> list[dict]:
    """Supprime les doublons. Clé : adresse précise + commune + surface si disponible,
    sinon empreinte du texte brut (annonces sans adresse précise)."""
    seen = set()
    out = []
    for it in items:
        if it.get("Adresse_precise"):
            key = ("addr", it.get("Commune", "").strip().lower(),
                   (it.get("Adresse") or "").strip().lower(), it.get("Surface_m2"))
        else:
            key = ("txt", it.get("Observations", "")[:120].strip().lower(), it.get("Surface_m2"))
        if key in seen:
            continue
        seen.add(key)
        out.append(it)
    return out


def _page_url(base_url: str, page: int, source_name: str) -> str:
    """Construit l'URL de la page N selon le schéma propre à chaque source.
    Schémas confirmés par inspection directe des sites (juillet 2026) :
    - BureauxLocaux : /page/N à la fin de l'URL de base
    - Geolocaux     : ?page=N (query param)
    - Arthur Loyd   : ?page=N
    - TournyMeyer   : ?page=N
    - CBRE          : non vérifié — on tente ?page=N
    """
    if page <= 1:
        return base_url
    if "bureauxlocaux.com" in base_url:
        return base_url.rstrip("/") + f"/page/{page}"
    return base_url.rstrip("/") + f"?page={page}"


def _run_search(url: str, source_name: str, commune: str, code_postal: str,
                type_bien: str, max_pages: int = 1,
                progress_cb=None) -> list[dict]:
    """Fetch une ou plusieurs pages de résultats.
    - max_pages=1  → comportement identique à l'ancien code
    - max_pages>1  → pagine automatiquement jusqu'à max_pages ou jusqu'à ce qu'une
                     page revienne vide (fin du catalogue).
    progress_cb(page, total_pages) optionnel, pour la barre de progression Streamlit.
    """
    all_items: list[dict] = []

    for page in range(1, max_pages + 1):
        page_url = _page_url(url, page, source_name)
        html = fetch(page_url)

        if html is None:
            break   # erreur réseau → on arrête

        if looks_like_js_shell(html):
            print(f"[live_search] {source_name} p.{page} : rendu JavaScript — "
                  f"extraction impossible. ({page_url})")
            break

        soup = BeautifulSoup(html, "html.parser")
        blocks = _split_listing_blocks(soup, page_url)

        if not blocks:
            break   # page vide → on a atteint la fin du catalogue

        page_items: list[dict] = []
        for b, href, photo in blocks:
            item = parse_listing_block(b, source_name, href or page_url,
                                        fallback_commune=commune, fallback_cp=code_postal,
                                        fallback_type_bien=type_bien)
            if item:
                item["Lien_direct"] = href is not None
                item["Photo_url"] = photo or ""
                if href is None:
                    item["Observations"] = ("⚠ Lien direct non trouvé. "
                                             + item["Observations"])
                page_items.append(item)

        if not page_items:
            break   # aucune annonce parseable → fin effective

        all_items.extend(page_items)

        if progress_cb:
            progress_cb(page, max_pages)

        if page < max_pages:
            time.sleep(0.4)   # courtoisie : 400 ms entre pages

    return dedupe_listings(all_items)


def search_bureauxlocaux(commune: str, code_postal: str, operation: str = "location",
                          type_bien: str = "bureaux", max_pages: int = 1) -> list[dict]:
    op = "location" if operation.lower().startswith("loc") else "vente"
    type_map = {"bureaux": "bureaux", "activités": "entrepots", "activites": "entrepots",
                "commerce": "locaux-commerciaux"}
    tb = type_map.get(type_bien.lower(), "bureaux")
    slug = f"{slugify(commune)}-{code_postal}"
    url = f"https://www.bureauxlocaux.com/immobilier-d-entreprise/annonces/{slug}/{op}-{tb}"
    return _run_search(url, "BureauxLocaux.com", commune, code_postal, type_bien,
                       max_pages=max_pages)


def search_geolocaux(commune: str, code_postal: str, operation: str = "location",
                      type_bien: str = "bureau", max_pages: int = 1) -> list[dict]:
    op = "location" if operation.lower().startswith("loc") else "vente"
    type_map = {"bureaux": "bureau", "activités": "local-activite", "activites": "local-activite",
                "commerce": "local-commercial"}
    tb = type_map.get(type_bien.lower(), "bureau")
    slug = f"{slugify(commune)}-{code_postal}"
    url = f"https://www.geolocaux.com/{op}/{tb}/{slug}/"
    return _run_search(url, "Geolocaux.com", commune, code_postal, type_bien,
                       max_pages=max_pages)


def search_arthurloyd(commune: str, code_postal: str, operation: str = "location",
                       type_bien: str = "bureaux", region: str = "pays-de-la-loire",
                       max_pages: int = 1) -> list[dict]:
    op = "location" if operation.lower().startswith("loc") else "vente"
    type_map = {
        "bureaux": f"bureau-{op}",
        "activités": f"locaux-activite-entrepots-{op}",
        "activites": f"locaux-activite-entrepots-{op}",
        "commerce": f"locaux-commerciaux-{op}",
    }
    tb = type_map.get(type_bien.lower(), f"bureau-{op}")
    slug = slugify(commune)
    url = f"https://www.arthur-loyd.com/{tb}/{region}/{slug}"
    return _run_search(url, "ArthurLoyd.com", commune, code_postal, type_bien,
                       max_pages=max_pages)


def search_tournymeyer(commune: str, code_postal: str, operation: str = "location",
                        type_bien: str = "bureaux", departement: str = "loire-atlantique",
                        max_pages: int = 1) -> list[dict]:
    op = "location" if operation.lower().startswith("loc") else "ventes"
    type_map = {"bureaux": "bureaux", "activités": "entrepots", "activites": "entrepots",
                "commerce": "commerces"}
    tb = type_map.get(type_bien.lower(), "bureaux")
    slug = slugify(commune)
    url = f"https://www.tournymeyer.fr/offres/{op}/{tb}/{departement}/{slug}/"
    return _run_search(url, "TournyMeyer.fr (JLL)", commune, code_postal, type_bien,
                       max_pages=max_pages)


def search_cbre(commune: str, code_postal: str, operation: str = "location",
                 type_bien: str = "bureaux", region: str = "pays-de-la-loire",
                 departement: str = "loire-atlantique", max_pages: int = 1) -> list[dict]:
    op = "location" if operation.lower().startswith("loc") else "vente"
    type_map = {"bureaux": "bureaux", "activités": "entrepots", "activites": "entrepots",
                "commerce": "commerces"}
    tb = type_map.get(type_bien.lower(), "bureaux")
    slug = slugify(commune)
    url = f"https://immobilier.cbre.fr/{op}-{tb}/{region}/{departement}/{slug}.aspx"
    return _run_search(url, "CBRE.fr (non vérifié)", commune, code_postal, type_bien,
                       max_pages=max_pages)


ALL_SOURCES = {
    "BureauxLocaux": search_bureauxlocaux,
    "Geolocaux": search_geolocaux,
    "ArthurLoyd": search_arthurloyd,
    "TournyMeyer (JLL)": search_tournymeyer,
    "CBRE (non vérifié)": search_cbre,
}


def search_all_sources(commune: str, code_postal: str, operation: str = "location",
                        type_bien: str = "bureaux", sources: list[str] | None = None,
                        max_pages: int = 1) -> list[dict]:
    """Interroge plusieurs sources et fusionne avec dédoublonnage global."""
    names = sources or list(ALL_SOURCES.keys())
    collected = []
    for name in names:
        fn = ALL_SOURCES.get(name)
        if fn:
            collected += fn(commune, code_postal, operation, type_bien,
                            max_pages=max_pages)
    return dedupe_listings(collected)


def diagnose(commune: str = "Nantes", code_postal: str = "44000") -> None:
    """Lance un test sur toutes les sources et explique clairement ce qui se passe.
    À exécuter en premier : `python live_search.py`"""
    print(f"Test de connexion en direct pour {commune} ({code_postal})…\n")
    for name, fn in ALL_SOURCES.items():
        print(f"— {name} —")
        results = fn(commune, code_postal)
        if results:
            n_complete = sum(1 for r in results if r.get("Complete"))
            print(f"  ✓ {len(results)} annonce(s) extraite(s), dont {n_complete} complète(s) "
                  f"(adresse précise + prix). Exemple :")
            print(f"    {results[0]}")
        else:
            print(f"  ✗ Aucune donnée extraite. Causes possibles :")
            print(f"    1) Le site rend son contenu en JavaScript (le plus probable) →")
            print(f"       il faudrait Playwright/Selenium au lieu de requests.")
            print(f"    2) L'URL générée ne correspond pas au vrai format du site.")
            print(f"    3) Le site bloque les requêtes automatisées (anti-bot).")
        print()
        time.sleep(1)
    print("⚠ CBRE : URL extrapolée à partir d'un seul exemple observé, non garantie pour "
          "tous les types de biens. Vérifie le résultat ci-dessus avant de t'y fier.")


if __name__ == "__main__":
    diagnose()
