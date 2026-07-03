"""
Module apis.py — Intégrations foncières + Targomo
══════════════════════════════════════════════════
APIs intégrées :

1. DVF (data.gouv.fr)       — transactions notariales, gratuit, sans clé
2. DVF+ Cerema (apidf)       — transactions enrichies, endpoint open
3. Targomo                   — isochrones multi-modes, clé gratuite
4. PLU / GPU                 — zones d'urbanisme PLU, gratuit (WMS + API)
5. Cadastre IGN              — info parcellaire (surface, section)
6. Géorisques                — risques naturels et technologiques
7. ADS / SITADEL             — permis de construire récents

Tous les appels API incluent un timeout et un fallback silencieux.
"""

from __future__ import annotations

import io
import json
import math
import time
from functools import lru_cache

import pandas as pd
import requests

# ══════════════════════════════════════════════════════════════════════════════
# HEADERS standard — toujours envoyer un User-Agent propre
# ══════════════════════════════════════════════════════════════════════════════
_H = {"User-Agent": "Savills-Valuation-Tool/1.0 (contact@savills.fr)"}


def _get(url: str, params: dict | None = None, timeout: int = 12,
         headers: dict | None = None) -> dict | None:
    """Requête GET avec gestion d'erreur silencieuse."""
    try:
        r = requests.get(url, params=params, timeout=timeout,
                         headers={**_H, **(headers or {})})
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"[API] {url[:60]} — {type(e).__name__}: {e}")
        return None


def _post(url: str, payload: dict, timeout: int = 20,
          headers: dict | None = None) -> dict | None:
    """Requête POST avec gestion d'erreur silencieuse."""
    try:
        r = requests.post(url, json=payload, timeout=timeout,
                          headers={**_H, "Content-Type": "application/json",
                                   **(headers or {})})
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"[API] POST {url[:60]} — {type(e).__name__}: {e}")
        return None


# ══════════════════════════════════════════════════════════════════════════════
# GEOCODAGE — BAN (Base Adresse Nationale)
# Déjà utilisé dans app.py, mais exposé ici pour enrichir les comparables
# ══════════════════════════════════════════════════════════════════════════════

def geocode_address(address: str) -> dict | None:
    """
    Géocode une adresse française via l'API BAN (gratuite, sans clé).
    Retourne {lat, lon, label, city, postcode} ou None.
    """
    data = _get("https://api-adresse.data.gouv.fr/search/",
                params={"q": address, "limit": 1})
    if not data:
        return None
    feats = data.get("features", [])
    if not feats:
        return None
    f     = feats[0]
    props = f.get("properties", {})
    lon, lat = f["geometry"]["coordinates"]
    return {
        "lat":      lat,
        "lon":      lon,
        "label":    props.get("label", address),
        "city":     props.get("city", ""),
        "postcode": props.get("postcode", ""),
        "score":    props.get("score", 0),
    }


# ══════════════════════════════════════════════════════════════════════════════
# PLU / URBANISME — Géoportail de l'Urbanisme (GPU)
# Endpoint : geoportail-urbanisme.gouv.fr + data.geopf.fr WMS
# ══════════════════════════════════════════════════════════════════════════════

# Types de zones PLU avec descriptions lisibles
PLU_ZONE_LABELS = {
    "U":   "Zone Urbaine — constructibilité immédiate",
    "AU":  "Zone À Urbaniser — future urbanisation",
    "N":   "Zone Naturelle — protection stricte, pas de construction",
    "A":   "Zone Agricole — activité agricole uniquement",
    "Nh":  "Zone Naturelle avec habitat existant",
    "UCa": "Zone Urbaine Centre-A (spécifique commune)",
    "UE":  "Zone Urbaine d'activités Économiques",
    "1AU": "Zone À Urbaniser Phase 1",
    "2AU": "Zone À Urbaniser Phase 2",
}

PLU_ZONE_COLORS = {
    "U":  "#FFA500",   # orange — zone urbaine
    "AU": "#FF6666",   # rouge clair — à urbaniser
    "N":  "#228B22",   # vert — naturel
    "A":  "#90EE90",   # vert clair — agricole
}


@lru_cache(maxsize=512)
def fetch_plu_zone(lat: float, lon: float) -> dict | None:
    """
    Récupère la zone PLU pour un point lat/lon.

    Source 1 : Géoportail de l'Urbanisme (GPU) — API officielle
    Source 2 : IGN APIcarto GPU (fallback)

    Returns:
        dict avec zone_type (U/AU/N/A), zone_libelle, zone_description
        ou None si indisponible.
    """
    geom = json.dumps({"type": "Point", "coordinates": [lon, lat]})

    # ── Source 1 : GPU (Géoportail Urbanisme) ───────────────────────────
    data = _get(
        "https://www.geoportail-urbanisme.gouv.fr/api/search/partitions",
        params={"lat": lat, "lon": lon, "type": "zone_urba"},
        timeout=10,
    )
    if data:
        items = data.get("results") or data.get("features") or []
        if items:
            item  = items[0]
            props = item.get("properties", item)
            zone  = props.get("typezone", "") or props.get("type_zone", "")
            lib   = props.get("libelle", "") or props.get("lib_zone", "")
            return {
                "zone_type":   zone,
                "zone_libelle": lib,
                "zone_description": PLU_ZONE_LABELS.get(zone, lib or zone),
                "partition":   props.get("partition", ""),
                "libelong":    props.get("libelong", ""),
                "source":      "GPU",
            }

    # ── Source 2 : IGN APIcarto GPU ─────────────────────────────────────
    data2 = _get(
        "https://apicarto.ign.fr/api/gpu/zone-urba",
        params={"geometry": geom, "_limit": 1},
        timeout=10,
    )
    if data2:
        feats = data2.get("features", [])
        if feats:
            props = feats[0].get("properties", {})
            zone  = props.get("typezone", "")
            lib   = props.get("libelle", "")
            return {
                "zone_type":    zone,
                "zone_libelle": lib,
                "zone_description": PLU_ZONE_LABELS.get(zone, lib or zone),
                "partition":    props.get("partition", ""),
                "libelong":     props.get("libelong", ""),
                "source":       "IGN",
            }

    return None


def get_plu_wms_layer() -> dict:
    """
    Retourne la configuration du layer WMS PLU pour Folium.
    Données : IGN / Géoportail de l'Urbanisme via data.geopf.fr
    Aucune clé requise pour la visualisation WMS.
    """
    return {
        "url":     "https://data.geopf.fr/wms-r/wms?",
        "layers":  "URBANISMES_REGLEMENTES.ZONE_URBA",
        "fmt":     "image/png",
        "name":    "🏙️ Zones PLU",
        "transparent": True,
        "version": "1.3.0",
        "opacity": 0.35,
    }


# ══════════════════════════════════════════════════════════════════════════════
# CADASTRE — IGN APIcarto
# Surface parcellaire, section, référence cadastrale
# ══════════════════════════════════════════════════════════════════════════════

@lru_cache(maxsize=512)
def fetch_cadastre_info(lat: float, lon: float) -> dict | None:
    """
    Récupère les informations cadastrales d'une parcelle.

    Source : IGN APIcarto cadastre (gratuit)
    Endpoint : https://apicarto.ign.fr/api/cadastre/parcelle

    Returns:
        dict avec section, numero, contenance (m²), commune, code_dep, code_com
        ou None si indisponible.
    """
    data = _get(
        "https://apicarto.ign.fr/api/cadastre/parcelle",
        params={"lon": lon, "lat": lat, "_limit": 1},
        timeout=10,
    )
    if not data:
        return None
    feats = data.get("features", [])
    if not feats:
        return None
    props = feats[0].get("properties", {})
    return {
        "section":     props.get("section", ""),
        "numero":      props.get("numero", ""),
        "contenance":  props.get("contenance"),    # surface en m²
        "commune":     props.get("commune", ""),
        "code_dep":    props.get("code_dep", ""),
        "code_com":    props.get("code_com", ""),
        "ref_cadastrale": f"{props.get('code_dep','')}{props.get('code_com','')}"
                          f"{'0'+props.get('section','') if len(props.get('section',''))==1 else props.get('section','')}"
                          f"{props.get('numero','').zfill(4)}",
        "source": "IGN Cadastre",
    }


# ══════════════════════════════════════════════════════════════════════════════
# GÉORISQUES — Risques naturels et technologiques
# Source : georisques.gouv.fr (gratuit, sans clé)
# ══════════════════════════════════════════════════════════════════════════════

@lru_cache(maxsize=256)
def fetch_georisques(lat: float, lon: float) -> dict | None:
    """
    Récupère les indicateurs de risques pour un point géographique.

    Source : Géorisques API v1 (BRGM/DGPR)
    Endpoint : https://georisques.gouv.fr/api/v1/

    Risques retournés : argile, sismicité, radon, inondation, ICPE, BASIAS

    Returns:
        dict avec les niveaux de risque ou None si indisponible.
    """
    # Risque argile (retrait-gonflement)
    argile = _get(
        "https://georisques.gouv.fr/api/v1/argiles",
        params={"latlon": f"{lon},{lat}", "rayon": 1},
        timeout=10,
    )

    # Risque sismique
    seisme = _get(
        "https://georisques.gouv.fr/api/v1/zonage_sismique",
        params={"latlon": f"{lon},{lat}", "rayon": 1},
        timeout=10,
    )

    # ICPE (installations classées) à proximité
    icpe = _get(
        "https://georisques.gouv.fr/api/v1/installations_classees",
        params={"latlon": f"{lon},{lat}", "rayon": 500},
        timeout=10,
    )

    # Risque radon
    radon = _get(
        "https://georisques.gouv.fr/api/v1/radon",
        params={"latlon": f"{lon},{lat}", "rayon": 1},
        timeout=10,
    )

    result = {}

    if argile and argile.get("data"):
        niv = argile["data"][0].get("niveauAlea", "")
        result["argile"] = {"niveau": niv, "label": f"Argile — {niv or 'NC'}"}

    if seisme and seisme.get("data"):
        zone = seisme["data"][0].get("zoneDesc", "")
        result["sismicite"] = {"zone": zone, "label": f"Sismicité zone {zone}"}

    if icpe and icpe.get("data"):
        n = len(icpe["data"])
        result["icpe"] = {"nb": n, "label": f"{n} ICPE dans 500 m" if n else "Aucune ICPE proche"}

    if radon and radon.get("data"):
        cat = radon["data"][0].get("categorieRadon", "")
        result["radon"] = {"categorie": cat, "label": f"Radon catégorie {cat}"}

    return result if result else None


# ══════════════════════════════════════════════════════════════════════════════
# ADS / SITADEL — Permis de construire (autorisations d'urbanisme)
# Source : data.gouv.fr (SITADEL)
# ══════════════════════════════════════════════════════════════════════════════

def fetch_recent_permits(commune_insee: str, years: int = 3) -> pd.DataFrame:
    """
    Récupère les permis de construire récents pour une commune via SITADEL.

    Source : data.gouv.fr — Fichiers SITADEL par département
    URL pattern : https://files.data.gouv.fr/sitadel/pc-{dept}-{year}.csv
    Note : Les fichiers SITADEL sont par département et année.

    Returns:
        DataFrame avec les permis (date, adresse, type, surface) ou DataFrame vide.
    """
    from datetime import datetime
    import io as _io
    
    current_year = datetime.now().year
    dept = commune_insee[:2]
    rows = []

    for year in range(current_year - years, current_year + 1):
        url = f"https://files.data.gouv.fr/sitadel/pc-{dept}-{year}.csv"
        try:
            r = requests.get(url, timeout=20, headers=_H)
            if r.status_code == 200:
                df_raw = pd.read_csv(_io.StringIO(r.text), sep=";",
                                     dtype=str, low_memory=False,
                                     encoding="utf-8", on_bad_lines="skip")
                # Filtrer par commune
                code_col = next((c for c in df_raw.columns
                                 if "commune" in c.lower() or "code_com" in c.lower()), None)
                if code_col:
                    df_raw = df_raw[df_raw[code_col].astype(str).str.startswith(commune_insee[:5])]
                rows.append(df_raw)
        except Exception:
            pass

    if not rows:
        return pd.DataFrame()

    return pd.concat(rows, ignore_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# DVF & DVF+ — Transactions foncières
# ══════════════════════════════════════════════════════════════════════════════

COMMUNE_INSEE = {
    "Nantes": "44109", "Saint-Herblain": "44162", "Rezé": "44143",
    "Orvault": "44112", "Saint-Sébastien-sur-Loire": "44190",
    "Carquefou": "44022", "Vertou": "44215", "La Chapelle-sur-Erdre": "44026",
    "Bouguenais": "44018", "Couëron": "44047", "Sainte-Luce-sur-Loire": "44172",
    "Thouaré-sur-Loire": "44201", "Basse-Goulaine": "44007",
    "Haute-Goulaine": "44071", "Sautron": "44191", "Treillières": "44205",
    "Les Sorinières": "44193", "Indre": "44074",
}

LOCAL_TERTIAIRE = {
    "Local industriel. commercial ou assimilé",
    "Appartement",
}

DVF_BASE_URL = "https://files.data.gouv.fr/geo-dvf/latest/csv"


@lru_cache(maxsize=32)
def download_dvf_commune(code_insee: str) -> pd.DataFrame | None:
    dept = code_insee[:2]
    url  = f"{DVF_BASE_URL}/{dept}/communes/{code_insee}.csv"
    try:
        r = requests.get(url, timeout=30, headers=_H)
        r.raise_for_status()
        return pd.read_csv(io.StringIO(r.text), sep=",", dtype=str, low_memory=False)
    except Exception as e:
        print(f"[DVF] {url[:60]} — {e}")
        return None


def _to_float(v) -> float | None:
    try:
        return float(str(v).replace(",", ".").replace(" ", ""))
    except (ValueError, TypeError):
        return None


def get_dvf_transactions(
    lat: float, lon: float, radius_m: float = 2000,
    commune: str = "Nantes", types_locaux: set | None = None,
    annees: int = 5,
) -> pd.DataFrame:
    if types_locaux is None:
        types_locaux = LOCAL_TERTIAIRE
    code = COMMUNE_INSEE.get(commune, "44109")
    raw  = download_dvf_commune(code)
    if raw is None or raw.empty:
        return pd.DataFrame()

    needed = ["date_mutation","nature_mutation","valeur_fonciere","type_local",
              "surface_reelle_bati","latitude","longitude",
              "adresse_numero","adresse_nom_voie","nom_commune"]
    df = raw[[c for c in needed if c in raw.columns]].copy()
    if "nature_mutation" in df.columns:
        df = df[df["nature_mutation"].str.contains("Vente", na=False)]
    if "type_local" in df.columns and types_locaux:
        df = df[df["type_local"].isin(types_locaux)]
    df["_lat"] = df["latitude"].apply(_to_float)
    df["_lon"] = df["longitude"].apply(_to_float)
    df = df.dropna(subset=["_lat","_lon"])

    if "date_mutation" in df.columns:
        from datetime import datetime, timedelta
        cutoff = (datetime.now()-timedelta(days=annees*365)).strftime("%Y-%m-%d")
        df = df[df["date_mutation"] >= cutoff]
    if df.empty:
        return pd.DataFrame()

    def dist(r):
        R, p = 6371000, math.pi/180
        a = (math.sin((r["_lat"]-lat)*p/2)**2
             + math.cos(lat*p)*math.cos(r["_lat"]*p)*math.sin((r["_lon"]-lon)*p/2)**2)
        return 2*R*math.asin(math.sqrt(a))

    df["distance_m"] = df.apply(dist, axis=1)
    df = df[df["distance_m"] <= radius_m]
    if df.empty:
        return pd.DataFrame()

    df["prix_total"] = df["valeur_fonciere"].apply(_to_float)
    df["surface_m2"] = df["surface_reelle_bati"].apply(_to_float)
    df["prix_m2"]    = df.apply(lambda r: round(r["prix_total"]/r["surface_m2"])
                                if r["prix_total"] and r["surface_m2"] and r["surface_m2"]>0
                                else None, axis=1)
    df["adresse"]    = (df.get("adresse_numero",pd.Series(dtype=str)).fillna("")
                        + " " + df.get("adresse_nom_voie",pd.Series(dtype=str)).fillna("")
                        ).str.strip()
    df["commune_dvf"] = df.get("nom_commune",pd.Series(dtype=str)).fillna(commune)

    result = df[["date_mutation","adresse","commune_dvf","type_local",
                 "surface_m2","prix_total","prix_m2","_lat","_lon","distance_m"]].copy()
    return result.sort_values("distance_m").rename(columns={
        "date_mutation":"Date","adresse":"Adresse","commune_dvf":"Commune",
        "type_local":"Type","surface_m2":"Surface (m²)","prix_total":"Prix total (€)",
        "prix_m2":"Prix/m²","distance_m":"Distance (m)",
    })


def dvf_summary(df: pd.DataFrame) -> dict:
    vals = df["Prix/m²"].dropna()
    if vals.empty: return {}
    return {"n":len(df),"median":vals.median(),"mean":vals.mean(),
            "min":vals.min(),"max":vals.max(),"q1":vals.quantile(0.25),"q3":vals.quantile(0.75)}


# ══════════════════════════════════════════════════════════════════════════════
# TARGOMO — Isochrones multi-modes
# Clé gratuite : targomo.com/developers
# ══════════════════════════════════════════════════════════════════════════════

TARGOMO_ENDPOINT = "https://api.targomo.com/westcentraleurope/v1/polygon"

TARGOMO_COLORS = {
    5:  "#27AE60",
    10: "#F39C12",
    15: "#E74C3C",
    20: "#8E44AD",
    30: "#2C3E50",
}

TRAVEL_MODES = {
    "walk":    "À pied",
    "bike":    "Vélo",
    "car":     "Voiture",
    "transit": "Transports en commun",
}


def fetch_isochrones(lat: float, lon: float, api_key: str,
                     times_minutes: list[int] | None = None,
                     mode: str = "walk") -> dict | None:
    """
    Calcule des polygones isochrones via l'API Targomo.

    Args:
        lat, lon      : coordonnées d'origine
        api_key       : clé API Targomo (gratuite sur targomo.com/developers)
        times_minutes : durées en minutes [5, 10, 15] par défaut
        mode          : walk | bike | car | transit

    Documentation : https://docs.targomo.com/core/
    """
    if times_minutes is None:
        times_minutes = [5, 10, 15]

    payload = {
        "sources": [{"id": "origin", "lat": lat, "lng": lon}],
        "polygon": {
            "values":           sorted(times_minutes, reverse=True),
            "intersectionMode": "union",
            "serializer":       "geojson",
            "decimalPrecision": 5,
        },
        "edgeWeight":    "time",
        "travelType":    mode,
        "maxEdgeWeight": max(times_minutes) * 60,
    }

    result = _post(
        TARGOMO_ENDPOINT,
        payload=payload,
        headers={"X-Api-Key": api_key},
        timeout=25,
    )
    if result is None:
        return {"error": "Requête Targomo échouée (timeout ou erreur réseau)."}

    # Targomo renvoie le GeoJSON dans data.features ou features selon la version
    if "data" in result:
        return result
    if "features" in result:
        return {"data": result}
    if "message" in result:
        return {"error": result["message"]}
    return {"error": str(result)[:200]}


def add_isochrones_to_map(folium_map, geojson_result: dict,
                          times_minutes: list[int], mode: str = "walk") -> None:
    """Ajoute les polygones isochrones sur une carte Folium."""
    import folium

    if not geojson_result or "error" in geojson_result:
        return

    features = (geojson_result.get("data", {}).get("features", [])
                or geojson_result.get("features", []))
    mode_label = TRAVEL_MODES.get(mode, mode)

    for feat in features:
        props    = feat.get("properties", {})
        time_sec = props.get("time") or props.get("value") or 0
        time_min = round(time_sec / 60)
        color    = TARGOMO_COLORS.get(
            min(TARGOMO_COLORS, key=lambda t: abs(t - time_min)), "#888")
        folium.GeoJson(
            feat,
            style_function=lambda _, c=color: {
                "fillColor": c, "color": c,
                "weight": 1.5, "fillOpacity": 0.12, "opacity": 0.5,
            },
            tooltip=f"{mode_label} — {time_min} min",
        ).add_to(folium_map)

    colors_html_parts = []
    for t in sorted(times_minutes):
        c = TARGOMO_COLORS.get(t, "#888")
        colors_html_parts.append(
            "<span style='color:" + c + "'>\u25cf</span> " + str(t) + " min<br>"
        )
    colors_html = "".join(colors_html_parts)
    import folium.element
    folium_map.get_root().html.add_child(folium.element.Element(f"""
    <div style="position:fixed;bottom:80px;right:24px;z-index:999;
         background:white;padding:10px 14px;border-radius:8px;
         box-shadow:0 2px 8px rgba(0,0,0,.2);font-family:Arial;font-size:12px">
      <b>Isochrones — {mode_label}</b><br>{colors_html}
    </div>"""))


# ══════════════════════════════════════════════════════════════════════════════
# ENRICHISSEMENT COMPARABLE — combine toutes les APIs pour un seul point
# ══════════════════════════════════════════════════════════════════════════════

def enrich_comparable_with_geodata(item: dict) -> dict:
    """
    Enrichit un comparable avec les données géographiques disponibles :
    - Géocodage si lat/lon manquants (BAN)
    - Zone PLU (GPU)
    - Info cadastrale (IGN)
    - Risques (Géorisques)

    Appelé après le scraping, avant l'affichage dans l'app.
    """
    updated = dict(item)

    # 1. Géocodage si nécessaire
    lat = _to_float(item.get("Latitude") or item.get("_lat"))
    lon = _to_float(item.get("Longitude") or item.get("_lon"))

    if lat is None or lon is None:
        adresse = str(item.get("Adresse", "") or "")
        commune = str(item.get("Commune", "") or "")
        cp      = str(item.get("Code_postal", "") or "")
        if not any(x in adresse.lower() for x in ("non disponible", "non extraite")):
            q = f"{adresse} {cp} {commune}".strip()
        else:
            q = f"{cp} {commune}".strip()
        if q:
            geo = geocode_address(q)
            if geo:
                lat, lon = geo["lat"], geo["lon"]
                updated["_lat"] = lat
                updated["_lon"] = lon

    if lat is None or lon is None:
        return updated

    # 2. Zone PLU
    plu = fetch_plu_zone(lat, lon)
    if plu:
        updated["PLU_zone"]      = plu.get("zone_type", "")
        updated["PLU_libelle"]   = plu.get("zone_libelle", "")
        updated["PLU_desc"]      = plu.get("zone_description", "")
        updated["PLU_libelong"]  = plu.get("libelong", "")

    # 3. Cadastre
    cadastre = fetch_cadastre_info(lat, lon)
    if cadastre:
        updated["Cadastre_ref"]        = cadastre.get("ref_cadastrale", "")
        updated["Cadastre_section"]    = cadastre.get("section", "")
        updated["Cadastre_parcelle"]   = cadastre.get("numero", "")
        updated["Cadastre_contenance"] = cadastre.get("contenance")   # m²

    return updated


def enrich_comparables_batch(items: list[dict], max_items: int = 30,
                              progress_cb=None) -> list[dict]:
    """Enrichit un lot de comparables avec les données géo/PLU/cadastre."""
    out = []
    for i, item in enumerate(items):
        enriched = enrich_comparable_with_geodata(item)
        out.append(enriched)
        if progress_cb:
            progress_cb(i + 1, min(len(items), max_items))
        if i >= max_items - 1:
            out.extend(items[max_items:])
            break
        time.sleep(0.08)   # courtoisie API
    return out
