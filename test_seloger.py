"""Test ciblé des URLs réelles fournies : SeLoger BC, SeLoger, JLL."""
import requests, re

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
           "Accept-Language": "fr-FR,fr;q=0.9",
           "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"}

URLS = {
    "SeLoger-BC":  "https://www.seloger-bureaux-commerces.com/location/bureau/ile-de-france/paris",
    "SeLoger":     "https://www.seloger.com/recherche/location/locaux-professionnels/ile-de-france/paris-75000/",
    "JLL":         "https://immobilier.jll.fr/location-bureaux/location-bureaux-paris",
}

ADDR = re.compile(r'\b\d{1,3}\s+(rue|avenue|av\.|bd|boulevard|place|quai|cours|all[ée]e)\b', re.I)
PRIX = re.compile(r'\d[\d\s]{2,}\s*(€|EUR|euros)', re.I)
SURF = re.compile(r'\d{2,}\s*m²', re.I)

for name, url in URLS.items():
    try:
        r = requests.get(url, timeout=15, headers=HEADERS)
        html = r.text
        size = len(html)
        n_addr = len(ADDR.findall(html))
        n_prix = len(PRIX.findall(html))
        n_surf = len(SURF.findall(html))
        js = "OUI" if ("__NEXT_DATA__" in html or "window.__" in html or size < 40000) else "non"
        blocked = "CLOUDFLARE" if ("cf-browser" in html or "challenge" in html.lower() or "captcha" in html.lower()) else ""
        print(f"{name:14}[{r.status_code}] {size//1000:4}Ko | adr~{n_addr} prix~{n_prix} surf~{n_surf} | JS:{js} {blocked}")
    except Exception as e:
        print(f"{name:14}ERREUR {type(e).__name__}: {str(e)[:50]}")
