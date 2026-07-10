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

# Cache manuel : on ne veut mémoriser QUE les succès. @lru_cache mettait en
# cache les None d'un échec réseau transitoire, ce qui « gelait » ensuite la
# commune sur un résultat vide jusqu'au redémarrage du serveur.
_DVF_CACHE: dict[str, pd.DataFrame] = {}


def download_dvf_commune(code_insee: str) -> pd.DataFrame | None:
    if code_insee in _DVF_CACHE:
        return _DVF_CACHE[code_insee]
    dept = code_insee[:2]
    url  = f"{DVF_BASE_URL}/{dept}/communes/{code_insee}.csv"
    # Certains CDN data.gouv rejettent les User-Agent « bot ». On tente d'abord
    # l'en-tête projet, puis un repli navigateur si la 1re requête est refusée.
    for _hdr in (_H, {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}):
        try:
            r = requests.get(url, timeout=30, headers=_hdr)
            r.raise_for_status()
            df = pd.read_csv(io.StringIO(r.text), sep=",", dtype=str, low_memory=False)
            _DVF_CACHE[code_insee] = df           # on ne cache que le succès
            return df
        except Exception as e:
            print(f"[DVF] {url[:60]} ({_hdr.get('User-Agent','')[:15]}…) — {e}")
            continue
    return None


def _to_float(v) -> float | None:
    try:
        return float(str(v).replace(",", ".").replace(" ", ""))
    except (ValueError, TypeError):
        return None


@lru_cache(maxsize=256)
def resolve_insee(commune: str, lat: float | None = None, lon: float | None = None) -> str | None:
    """Résout le code INSEE d'une commune.

    1) table locale COMMUNE_INSEE (rapide, Nantes métropole) ;
    2) sinon interrogation BAN : d'abord par coordonnées (reverse), le plus
       fiable ; à défaut par nom de commune. La BAN renvoie `citycode` = INSEE.

    Corrige le bug où toute commune hors table (Rennes, Bordeaux…) retombait
    silencieusement sur Nantes (44109) → DVF interrogeait la mauvaise ville.
    """
    if commune in COMMUNE_INSEE:
        return COMMUNE_INSEE[commune]
    # reverse géo par coordonnées si dispo (plus robuste que le nom)
    if lat is not None and lon is not None:
        data = _get("https://api-adresse.data.gouv.fr/reverse/",
                    params={"lat": lat, "lon": lon, "limit": 1})
        feats = (data or {}).get("features", [])
        if feats:
            code = feats[0].get("properties", {}).get("citycode")
            if code:
                return code
    # repli : recherche par nom
    if commune:
        data = _get("https://api-adresse.data.gouv.fr/search/",
                    params={"q": commune, "type": "municipality", "limit": 1})
        feats = (data or {}).get("features", [])
        if feats:
            code = feats[0].get("properties", {}).get("citycode")
            if code:
                return code
    return None


def get_dvf_transactions(
    lat: float, lon: float, radius_m: float = 2000,
    commune: str = "Nantes", types_locaux: set | None = None,
    annees: int = 5,
) -> pd.DataFrame:
    if types_locaux is None:
        types_locaux = LOCAL_TERTIAIRE
    code = resolve_insee(commune, lat, lon) or COMMUNE_INSEE.get(commune, "44109")
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


def get_dvf_commune_complete(
    commune: str = "Nantes", types_locaux: set | None = None,
    annees: int = 5, ref_lat: float | None = None, ref_lon: float | None = None,
) -> pd.DataFrame:
    """
    Retourne TOUTES les transactions DVF d'une commune (pas de filtre par rayon),
    pour une cartographie exhaustive. Réutilise le fichier déjà mis en cache par
    download_dvf_commune — aucun appel réseau supplémentaire par rapport à
    get_dvf_transactions.

    Args:
        commune       : nom de la commune
        types_locaux  : types de locaux à inclure (défaut = tertiaire)
        annees        : nombre d'années à remonter
        ref_lat/ref_lon : point de référence optionnel pour calculer une
                          distance informative (n'exclut aucune ligne)

    Returns:
        DataFrame — mêmes colonnes que get_dvf_transactions, toute la commune.
    """
    if types_locaux is None:
        types_locaux = LOCAL_TERTIAIRE
    code = resolve_insee(commune, ref_lat, ref_lon) or COMMUNE_INSEE.get(commune, "44109")
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

    df["prix_total"] = df["valeur_fonciere"].apply(_to_float)
    df["surface_m2"] = df["surface_reelle_bati"].apply(_to_float)
    df["prix_m2"]    = df.apply(lambda r: round(r["prix_total"]/r["surface_m2"])
                                if r["prix_total"] and r["surface_m2"] and r["surface_m2"]>0
                                else None, axis=1)
    df["adresse"]    = (df.get("adresse_numero",pd.Series(dtype=str)).fillna("")
                        + " " + df.get("adresse_nom_voie",pd.Series(dtype=str)).fillna("")
                        ).str.strip()
    df["commune_dvf"] = df.get("nom_commune",pd.Series(dtype=str)).fillna(commune)

    # distance informative uniquement — n'exclut rien, sert juste au tri/affichage
    if ref_lat is not None and ref_lon is not None:
        def dist(r):
            R, p = 6371000, math.pi/180
            a = (math.sin((r["_lat"]-ref_lat)*p/2)**2
                 + math.cos(ref_lat*p)*math.cos(r["_lat"]*p)*math.sin((r["_lon"]-ref_lon)*p/2)**2)
            return 2*R*math.asin(math.sqrt(a))
        df["distance_m"] = df.apply(dist, axis=1)
    else:
        df["distance_m"] = None

    result = df[["date_mutation","adresse","commune_dvf","type_local",
                 "surface_m2","prix_total","prix_m2","_lat","_lon","distance_m"]].copy()
    if result["distance_m"].notna().any():
        result = result.sort_values("distance_m")
    return result.rename(columns={
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
# SOGEFI DVF+ — API payante, données CEREMA enrichies (clé requise)
# Contrat technique confirmé le 08/07/2026 via la doc Swagger fournie par l'utilisateur :
#   POST https://api.sogefi-sig.com/{CLÉ}/dvfplus/v1.0/sogefi/mutation/search
#     query  : filter={"champ[opérateur]": valeur}, skipGeometry, buffer
#     body   : {"geojson": {"type": "Polygon", "coordinates": [[[lon,lat], ...]]}}
#     limite : polygone ≤ 10 km², longueur ≤ 10 km
#   GET  https://api.sogefi-sig.com/{CLÉ}/dvfplus/v1.0/sogefi/infos
#     → confirmé fonctionnel, donnée source CEREMA, mise à jour semestrielle.
#
# Les noms de champs du filtre (valeurfonc, datemut) correspondent au schéma
# officiel CEREMA DVF+ — on s'appuie donc sur ce dictionnaire de champs connu
# pour parser les `properties` de chaque feature retournée. L'exemple de doc
# Swagger ne montrait qu'un placeholder ("foo":"bar"), donc le mapping ci-dessous
# est notre meilleure hypothèse informée : à valider/ajuster au premier vrai
# appel (voir raw_properties renvoyée pour debug).
# ══════════════════════════════════════════════════════════════════════════════

SOGEFI_BASE = "https://api.sogefi-sig.com"

# Dictionnaire de champs officiel CEREMA DVF+ (noms stables, documentés par le Cerema)
SOGEFI_FIELD_CANDIDATES = {
    "date":       ["datemut", "date_mutation"],
    "valeur":     ["valeurfonc", "valeur_fonciere"],
    "nature":     ["libnatmut", "nature_mutation"],
    "surface_bati": ["sbati", "surface_reelle_bati"],
    "surface_terrain": ["sterr", "surface_terrain"],
    "code_dep":   ["coddep", "code_departement"],
    "code_com":   ["l_codinsee", "code_commune", "coddep"],
    "nb_locaux":  ["nbdispo", "nblocmut"],
    "id_parcelle":["l_idpar", "idpar"],
    # Adresse : pas confirmée dans le schéma DVF+ (structuré par parcelle, pas
    # forcément par adresse postale) — candidats les plus probables, à valider
    # via l'encart Debug de l'app au premier vrai appel.
    "adresse":    ["adresse", "l_adresse", "voie", "libvoie", "nomvoie",
                   "adresse_nom_voie", "l_voie"],
    "commune":    ["commune", "libcom", "nomcom", "l_libcom"],
}


def _sogefi_get_field(props: dict, key: str):
    """Cherche une valeur dans `properties` en essayant plusieurs noms de
    champs candidats (le schéma exact n'a pas pu être confirmé à l'avance)."""
    for candidate in SOGEFI_FIELD_CANDIDATES.get(key, []):
        if candidate in props and props[candidate] not in (None, ""):
            return props[candidate]
    return None


def _bbox_polygon(lat: float, lon: float, radius_m: float) -> dict:
    """Construit un polygone carré GeoJSON autour d'un point, dimensionné pour
    rester sous la limite de 10 km² / 10 km imposée par l'API SOGEFI.
    1° de latitude ≈ 111 320 m ; 1° de longitude ≈ 111 320 * cos(lat) m."""
    radius_m = min(radius_m, 1500)  # 1500m de rayon → carré ~9 km², sous la limite de 10 km²
    dlat = radius_m / 111_320
    dlon = radius_m / (111_320 * math.cos(math.radians(lat)) or 1)
    return {
        "type": "Polygon",
        "coordinates": [[
            [lon - dlon, lat - dlat],
            [lon - dlon, lat + dlat],
            [lon + dlon, lat + dlat],
            [lon + dlon, lat - dlat],
            [lon - dlon, lat - dlat],
        ]],
    }


def fetch_sogefi_mutations(
    lat: float, lon: float, api_key: str,
    radius_m: float = 800, date_from: str | None = None,
    valeur_min: float | None = None,
    timeout: int = 25,
) -> dict:
    """
    Interroge l'API SOGEFI DVF+ pour les mutations foncières dans un rayon
    donné autour d'un point. Nécessite une clé API payante SOGEFI.

    Args:
        lat, lon   : centre de la recherche
        api_key    : clé API SOGEFI (intégrée dans l'URL par leur convention)
        radius_m   : rayon en mètres (plafonné à 1500m pour respecter la
                     limite de surface de l'API)
        date_from  : filtre optionnel "AAAA-MM-JJ" — mutations depuis cette date
        valeur_min : filtre optionnel — valeur foncière minimale en €

    Returns:
        {"ok": True, "features": [...]} en cas de succès
        {"ok": False, "error": "..."} en cas d'échec (clé invalide, quota, etc.)
    """
    if not api_key:
        return {"ok": False, "error": "Clé API SOGEFI manquante."}

    url = f"{SOGEFI_BASE}/{api_key}/dvfplus/v1.0/sogefi/mutation/search"

    filter_obj = {}
    if date_from:
        filter_obj["datemut[gte]"] = date_from
    if valeur_min:
        filter_obj["valeurfonc[gte]"] = valeur_min

    params = {"skipGeometry": "true"}
    if filter_obj:
        params["filter"] = json.dumps(filter_obj)

    body = {"geojson": _bbox_polygon(lat, lon, radius_m)}

    try:
        r = requests.post(url, params=params, json=body, timeout=timeout,
                          headers={**_H, "Content-Type": "application/json"})
        if r.status_code == 401 or r.status_code == 403:
            return {"ok": False, "error": "Clé API SOGEFI invalide ou expirée (401/403)."}
        if r.status_code == 429:
            return {"ok": False, "error": "Quota de requêtes SOGEFI atteint (429)."}
        if r.status_code == 400:
            return {"ok": False, "error": f"Requête invalide (400) — polygone probablement "
                                          f"trop grand : {r.text[:200]}"}
        r.raise_for_status()
        data = r.json()
        return {"ok": True, "features": data.get("features", [])}
    except requests.exceptions.Timeout:
        return {"ok": False, "error": "Timeout — l'API SOGEFI n'a pas répondu à temps."}
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}


def sogefi_features_to_df(features: list[dict], center_lat: float, center_lon: float) -> pd.DataFrame:
    """Normalise les features GeoJSON SOGEFI DVF+ dans le même format de
    tableau que le DVF gratuit, pour un affichage cohérent dans l'onglet."""
    rows = []
    for feat in features:
        props = feat.get("properties", {}) or {}
        geom = feat.get("geometry", {}) or {}

        date_val   = _sogefi_get_field(props, "date")
        valeur     = _sogefi_get_field(props, "valeur")
        surface    = _sogefi_get_field(props, "surface_bati")
        nature     = _sogefi_get_field(props, "nature")
        adresse    = _sogefi_get_field(props, "adresse")
        commune    = _sogefi_get_field(props, "commune")

        prix_m2 = None
        try:
            if valeur and surface and float(surface) > 0:
                prix_m2 = round(float(valeur) / float(surface))
        except (ValueError, TypeError):
            pass

        # Centroïde approximatif pour la distance (si géométrie présente)
        lat_pt, lon_pt = None, None
        coords = geom.get("coordinates")
        if coords:
            try:
                # gère Polygon/MultiPolygon en prenant le premier point du premier anneau
                pt = coords
                while isinstance(pt, list) and pt and isinstance(pt[0], list):
                    pt = pt[0]
                if isinstance(pt, list) and len(pt) == 2:
                    lon_pt, lat_pt = pt[0], pt[1]
            except Exception:
                pass

        dist = None
        if lat_pt is not None and lon_pt is not None:
            R = 6371000
            p = math.pi / 180
            a = (math.sin((lat_pt - center_lat) * p / 2) ** 2
                 + math.cos(center_lat * p) * math.cos(lat_pt * p)
                 * math.sin((lon_pt - center_lon) * p / 2) ** 2)
            dist = 2 * R * math.asin(math.sqrt(a))

        rows.append({
            "Date":            date_val or "—",
            "Adresse":         adresse or "—",
            "Commune":         commune or "—",
            "Nature":          nature or "—",
            "Surface (m²)":    surface,
            "Prix total (€)":  valeur,
            "Prix/m²":         prix_m2,
            "_lat":            lat_pt,
            "_lon":            lon_pt,
            "Distance (m)":    dist,
            "_raw_properties": props,   # conservé pour debug/vérification du mapping
        })

    df = pd.DataFrame(rows)
    if not df.empty and "Distance (m)" in df.columns:
        df = df.sort_values("Distance (m)", na_position="last")
    return df


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

    # ⚠️ Targomo attend TOUTES les durées en SECONDES — aussi bien dans
    # polygon.values que dans maxEdgeWeight. Passer des minutes ici produisait
    # des polygones minuscules (15 s au lieu de 15 min) → « aucun résultat ».
    times_seconds = sorted((int(t) * 60 for t in times_minutes), reverse=True)

    payload = {
        "sources": [{"id": "origin", "lat": lat, "lng": lon}],
        "polygon": {
            "values":           times_seconds,
            "intersectionMode": "union",
            "serializer":       "geojson",
            "decimalPrecision": 5,
        },
        "edgeWeight":    "time",
        "travelType":    mode,
        "maxEdgeWeight": max(times_seconds),
    }

    # La clé API Targomo passe en paramètre d'URL (?key=…), pas en en-tête.
    # L'en-tête X-Api-Key est conservé en complément par sécurité (ignoré si inutile).
    result = _post(
        f"{TARGOMO_ENDPOINT}?key={api_key}",
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
    if "error" in result:
        return {"error": str(result["error"])[:200]}
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
                              progress_cb=None, parallel: bool = False,
                              workers: int = 4) -> list[dict]:
    """Enrichit un lot de comparables avec les données géo/PLU/cadastre.

    parallel=False (Sûr)   : séquentiel, 80ms entre appels, aucun risque.
    parallel=True  (Rapide): jusqu'à `workers` enrichissements simultanés,
                             plus rapide mais sollicite davantage les APIs
                             gratuites (PLU/GPU, cadastre) — léger risque 429.
    """
    to_process = items[:max_items]
    tail = items[max_items:]

    if parallel and len(to_process) > 1:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        results = [None] * len(to_process)
        done = 0
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futures = {ex.submit(enrich_comparable_with_geodata, item): idx
                       for idx, item in enumerate(to_process)}
            for fut in as_completed(futures):
                idx = futures[fut]
                try:
                    results[idx] = fut.result()
                except Exception:
                    results[idx] = to_process[idx]  # garde l'original si échec
                done += 1
                if progress_cb:
                    progress_cb(done, len(to_process))
        out = [r if r is not None else to_process[i] for i, r in enumerate(results)]
        out.extend(tail)
        return out

    # Mode séquentiel (Sûr)
    out = []
    for i, item in enumerate(to_process):
        out.append(enrich_comparable_with_geodata(item))
        if progress_cb:
            progress_cb(i + 1, len(to_process))
        time.sleep(0.08)   # courtoisie API
    out.extend(tail)
    return out
