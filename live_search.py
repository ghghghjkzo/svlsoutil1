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
    "Accept-Language": "fr-FR,fr;q=0.9",
}

SURFACE_RE = re.compile(r"(\d[\d\s]{0,5})\s?m\s?2|(\d[\d\s]{0,5})\s?m²", re.I)
LOYER_ANNUEL_M2_RE = re.compile(r"(\d[\d\s]{1,5})\s?€\s?(HT)?\s?/\s?m²\s?/\s?an", re.I)
LOYER_MENSUEL_RE = re.compile(r"(\d[\d\s]{1,7})\s?€?\s?(HT|TTC)?\s?(/|par)\s?mois", re.I)
PRIX_TOTAL_RE = re.compile(r"prix\s*:?\s*(\d[\d\s]{3,9})\s?€", re.I)

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

# Adresse : numéro + type de voie + nom, suivi d'un code postal puis d'une commune connue
ADDRESS_RE = re.compile(
    r"(\d+[\d\-\s]*\s+(?:Rue|Boulevard|Bd|Avenue|Av|All[ée]e|Quai|Place|Mail|Chemin|Route|Impasse)"
    r"\s+[^,\d]{2,55}?)\s*,?\s*(\d{5})\s+(" + _commune_alt + r")\b",
    re.I,
)

OP_TYPE_RE = re.compile(
    r"\b(Vente|Location)\s+(Bureaux?|Entrep[oô]ts?|Locaux?\s+d'activit[ée]s?|"
    r"Locaux?\s+commerciaux?|Commerces?)\b", re.I,
)

ETAT_KEYWORDS = [
    (re.compile(r"\bneuf\b", re.I), "Neuf"),
    (re.compile(r"enti[èe]rement r[ée]nov[ée]s?", re.I), "Restructuré"),
    (re.compile(r"r[ée]nov[ée]s?", re.I), "Restructuré"),
    (re.compile(r"\b[àa]\s+r[ée]nover\b", re.I), "Seconde main"),
    (re.compile(r"[ée]tat d'usage", re.I), "Seconde main"),
]

TYPE_ACTIF_MAP = {
    "bureau": "Bureaux", "bureaux": "Bureaux",
    "entrepot": "Activités", "entrepôt": "Activités", "entrepots": "Activités", "entrepôts": "Activités",
    "commerce": "Commerce", "commerces": "Commerce",
}


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
    if not s:
        return None
    try:
        return float(s.replace(" ", "").replace("\xa0", ""))
    except ValueError:
        return None


def _split_listing_blocks(soup: BeautifulSoup) -> list[str]:
    # tentative 1 : balises sémantiques probables
    candidates = soup.find_all(["article", "li", "div"],
                                class_=re.compile(r"annonce|listing|card|result|offre", re.I))
    blocks = [c.get_text(" ", strip=True) for c in candidates if len(c.get_text(strip=True)) > 60]
    if blocks:
        return blocks
    # tentative 2 (repli) : découpage du texte brut sur des marqueurs récurrents observés
    full_text = soup.get_text(" ", strip=True)
    parts = re.split(r"(?=Voir l'annonce|En savoir plus|Disponibilité\s?:)", full_text)
    return [p for p in parts if SURFACE_RE.search(p)]


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
    return "Non disponible"


def parse_listing_block(block: str, source_name: str, source_url: str,
                         fallback_commune: str = "", fallback_cp: str = "",
                         fallback_type_bien: str = "") -> dict | None:
    surf_m = SURFACE_RE.search(block)
    if not surf_m:
        return None
    surface = _clean_int(surf_m.group(1) or surf_m.group(2))
    if not surface:
        return None

    loyer_m2 = LOYER_ANNUEL_M2_RE.search(block)
    loyer_mensuel = LOYER_MENSUEL_RE.search(block)
    prix_total = PRIX_TOTAL_RE.search(block)

    loyer_annuel_m2 = _clean_int(loyer_m2.group(1)) if loyer_m2 else None
    if loyer_annuel_m2 is None and loyer_mensuel and surface:
        mensuel = _clean_int(loyer_mensuel.group(1))
        if mensuel:
            loyer_annuel_m2 = round(mensuel * 12 / surface, 1)

    prix_total_val = _clean_int(prix_total.group(1)) if prix_total else None
    prix_m2 = round(prix_total_val / surface, 1) if (prix_total_val and surface) else None

    # opération + type d'actif depuis l'amorce du bloc (ex. "Vente Bureaux 322 m²...")
    op_type_m = OP_TYPE_RE.search(block)
    operation = "Non disponible"
    if op_type_m:
        operation = "Vente" if op_type_m.group(1).lower() == "vente" else "Location"
    elif prix_total_val and not loyer_annuel_m2:
        operation = "Vente"
    elif loyer_annuel_m2 or loyer_mensuel:
        operation = "Location"

    type_actif = _detect_type_actif(block, fallback_type_bien)

    # adresse réelle + commune réelle (PAS les paramètres de recherche saisis par l'utilisateur)
    addr_m = ADDRESS_RE.search(block)
    if addr_m:
        adresse = addr_m.group(1).strip().rstrip(",")
        code_postal = addr_m.group(2)
        commune = addr_m.group(3)
    else:
        # repli : on a trouvé une commune connue dans le texte sans adresse précise
        com_m = COMMUNE_RE.search(block)
        adresse = "Non disponible (adresse précise non extraite, voir lien)"
        commune = com_m.group(1) if com_m else (fallback_commune or "Non disponible")
        code_postal = fallback_cp or "Non disponible"

    etat = _detect_etat(block)

    return {
        "Type_actif": type_actif,
        "Operation": operation,
        "Etat": etat,
        "Adresse": adresse,
        "Commune": commune,
        "Code_postal": code_postal,
        "Surface_m2": surface,
        "Loyer_facial_eur_m2_an": loyer_annuel_m2,
        "Prix_vente_total_eur": prix_total_val,
        "Prix_vente_eur_m2": prix_m2,
        "Date_transaction": "Non disponible (annonce active, pas une transaction signée)",
        "Source": source_name,
        "Lien": source_url,
        "Fiabilite": 2,
        "Observations": block[:280],
    }


def dedupe_listings(items: list[dict]) -> list[dict]:
    """Supprime les doublons (même adresse + même surface), garde la première occurrence."""
    seen = set()
    out = []
    for it in items:
        key = (it.get("Adresse", "").strip().lower(), it.get("Surface_m2"))
        if key in seen:
            continue
        seen.add(key)
        out.append(it)
    return out


def search_bureauxlocaux(commune: str, code_postal: str, operation: str = "location",
                          type_bien: str = "bureaux") -> list[dict]:
    op = "location" if operation.lower().startswith("loc") else "vente"
    type_map = {"bureaux": "bureaux", "activités": "entrepots", "activites": "entrepots",
                "commerce": "locaux-commerciaux"}
    tb = type_map.get(type_bien.lower(), "bureaux")
    slug = f"{slugify(commune)}-{code_postal}"
    url = f"https://www.bureauxlocaux.com/immobilier-d-entreprise/annonces/{slug}/{op}-{tb}"

    html = fetch(url)
    if html is None:
        return []
    if looks_like_js_shell(html):
        print("[live_search] BureauxLocaux : page probablement rendue en JavaScript — "
              "extraction impossible avec requests seul. Voir diagnose().")
        return []

    soup = BeautifulSoup(html, "html.parser")
    blocks = _split_listing_blocks(soup)
    out = []
    for b in blocks:
        item = parse_listing_block(b, "BureauxLocaux.com", url,
                                    fallback_commune=commune, fallback_cp=code_postal,
                                    fallback_type_bien=type_bien)
        if item:
            out.append(item)
    return dedupe_listings(out)


def search_geolocaux(commune: str, code_postal: str, operation: str = "location",
                      type_bien: str = "bureau") -> list[dict]:
    op = "location" if operation.lower().startswith("loc") else "vente"
    type_map = {"bureaux": "bureau", "activités": "local-activite", "activites": "local-activite",
                "commerce": "local-commercial"}
    tb = type_map.get(type_bien.lower(), "bureau")
    slug = f"{slugify(commune)}-{code_postal}"
    url = f"https://www.geolocaux.com/{op}/{tb}/{slug}/"

    html = fetch(url)
    if html is None:
        return []
    if looks_like_js_shell(html):
        print("[live_search] Geolocaux : page probablement rendue en JavaScript — "
              "extraction impossible avec requests seul. Voir diagnose().")
        return []

    soup = BeautifulSoup(html, "html.parser")
    blocks = _split_listing_blocks(soup)
    out = []
    for b in blocks:
        item = parse_listing_block(b, "Geolocaux.com", url,
                                    fallback_commune=commune, fallback_cp=code_postal,
                                    fallback_type_bien=type_bien)
        if item:
            out.append(item)
    return dedupe_listings(out)


def diagnose(commune: str = "Nantes", code_postal: str = "44000") -> None:
    """Lance un test sur les deux sources et explique clairement ce qui se passe.
    À exécuter en premier : `python live_search.py`"""
    print(f"Test de connexion en direct pour {commune} ({code_postal})…\n")
    for name, fn, url_hint in [
        ("BureauxLocaux", search_bureauxlocaux,
         "https://www.bureauxlocaux.com/immobilier-d-entreprise/annonces/<slug>/location-bureaux"),
        ("Geolocaux", search_geolocaux,
         "https://www.geolocaux.com/location/bureau/<slug>/"),
    ]:
        print(f"— {name} —")
        results = fn(commune, code_postal)
        if results:
            print(f"  ✓ {len(results)} annonce(s) extraite(s). Exemple :")
            print(f"    {results[0]}")
        else:
            print(f"  ✗ Aucune donnée extraite. Causes possibles :")
            print(f"    1) Le site rend son contenu en JavaScript (le plus probable) →")
            print(f"       il faudrait Playwright/Selenium au lieu de requests.")
            print(f"    2) L'URL générée ne correspond pas au vrai format du site.")
            print(f"       Vérifie manuellement dans ton navigateur : {url_hint}")
            print(f"    3) Le site bloque les requêtes automatisées (anti-bot).")
        print()
        time.sleep(1)


if __name__ == "__main__":
    diagnose()
