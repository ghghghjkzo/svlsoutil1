"""Reconnaissance multi-villes : teste quels sites immo tertiaires sont
scrapables ET donnent des adresses, pour plusieurs villes. Ne modifie rien.

Usage:  python3 test_sources.py            (teste Paris, Nantes, Bordeaux, Lyon)
        python3 test_sources.py paris      (teste juste Paris)
"""
import requests, re, sys

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
           "Accept-Language": "fr-FR,fr;q=0.9"}

# Villes de test : (slug, code postal, dept, region)
VILLES = {
    "paris":    ("paris",    "75000", "paris",            "ile-de-france"),
    "nantes":   ("nantes",   "44000", "loire-atlantique", "pays-de-la-loire"),
    "bordeaux": ("bordeaux", "33000", "gironde",          "nouvelle-aquitaine"),
    "lyon":     ("lyon",     "69000", "rhone",            "auvergne-rhone-alpes"),
}

# Fonctions qui construisent l'URL de listing pour une ville donnée.
# {slug}=nom ville, {cp}=code postal, {dept}=département, {region}=région
def build_urls(slug, cp, dept, region):
    return {
        "BureauxLocaux": f"https://www.bureauxlocaux.com/annonces/location/bureaux/{slug}-{cp}",
        "Geolocaux":     f"https://www.geolocaux.com/location/bureaux/{slug}-{cp[:2]}/",
        "Bureaux&Co":    f"https://www.bureauxandco.com/location/bureaux/{slug}",
        "JLL":           f"https://immobilier.jll.fr/location/bureaux/{slug}-{cp[:2]}",
        "KnightFrank":   f"https://www.knightfrank.fr/bureaux/location/{slug}",
        "ArthurLoyd":    f"https://www.arthur-loyd.com/bureau-location/{region}/{slug}",
        "Sinety":        f"https://www.sinety.com/location-bureaux-{slug}",
        "Alexbroc":      f"https://www.alexbrocante-immo.fr/{slug}",  # exemple, à ajuster
    }

ADDR_HINTS = re.compile(r'\b\d{1,3}\s+(rue|avenue|av\.|bd|boulevard|place|quai|cours|all[ée]e)\b', re.I)

villes_a_tester = sys.argv[1:] if len(sys.argv) > 1 else list(VILLES.keys())

for ville_key in villes_a_tester:
    if ville_key not in VILLES:
        print(f"Ville inconnue: {ville_key}"); continue
    slug, cp, dept, region = VILLES[ville_key]
    print(f"\n{'='*60}\n  {ville_key.upper()} ({cp})\n{'='*60}")
    for name, url in build_urls(slug, cp, dept, region).items():
        try:
            r = requests.get(url, timeout=12, headers=HEADERS)
            html = r.text
            size = len(html)
            n_addr = len(ADDR_HINTS.findall(html))
            js_heavy = "OUI" if (size < 40000 or "__NEXT_DATA__" in html or "window.__" in html) else "non"
            flag = "✓" if (r.status_code == 200 and n_addr > 3 and js_heavy == "non") else " "
            print(f"{flag} {name:16}[{r.status_code}] {size//1000:4}Ko | adr~{n_addr:3} | JS:{js_heavy}")
        except Exception as e:
            print(f"  {name:16}ERREUR {type(e).__name__}")
