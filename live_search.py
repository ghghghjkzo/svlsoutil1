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


def parse_listing_block(block: str, source_name: str, source_url: str) -> dict | None:
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

    return {
        "Type_actif": "Non disponible",
        "Operation": "Location" if (loyer_annuel_m2 or loyer_mensuel) else "Non disponible",
        "Etat": "Non disponible",
        "Adresse": "Non disponible (voir lien)",
        "Surface_m2": surface,
        "Loyer_facial_eur_m2_an": loyer_annuel_m2,
        "Prix_vente_total_eur": _clean_int(prix_total.group(1)) if prix_total else None,
        "Date_transaction": "Non disponible (annonce active, pas une transaction signée)",
        "Source": source_name,
        "Lien": source_url,
        "Fiabilite": 2,
        "Observations": block[:280],
    }


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
        item = parse_listing_block(b, "BureauxLocaux.com", url)
        if item:
            item["Commune"] = commune
            item["Code_postal"] = code_postal
            out.append(item)
    return out


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
        item = parse_listing_block(b, "Geolocaux.com", url)
        if item:
            item["Commune"] = commune
            item["Code_postal"] = code_postal
            out.append(item)
    return out


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
