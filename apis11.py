"""
Module apis.py — intégrations DVF et Targomo
─────────────────────────────────────────────
DVF  : Demandes de Valeurs Foncières (data.gouv.fr, gratuit, sans clé)
       → transactions réelles de vente immobilière notarisées en France

Targomo : API d'isochrones / temps de trajet (clé gratuite sur targomo.com)
          → polygone de zone accessible depuis un point en N minutes

"""

from __future__ import annotations

import io
import math
import re
import time
from functools import lru_cache

import pandas as pd
import requests

# ─── DVF ──────────────────────────────────────────────────────────────────────

# Correspondance commune Nantes Métropole → code INSEE à 5 chiffres
# Source : INSEE / data.gouv.fr (codes stables)
COMMUNE_INSEE = {
    "Nantes":                   "44109",
    "Saint-Herblain":           "44162",
    "Rezé":                     "44143",
    "Orvault":                  "44112",
    "Saint-Sébastien-sur-Loire":"44190",
    "Carquefou":                "44022",
    "Vertou":                   "44215",
    "La Chapelle-sur-Erdre":    "44026",
    "Bouguenais":               "44018",
    "Couëron":                  "44047",
    "Sainte-Luce-sur-Loire":    "44172",
    "Thouaré-sur-Loire":        "44201",
    "Basse-Goulaine":           "44007",
    "Haute-Goulaine":           "44071",
    "Sautron":                  "44191",
    "Treillières":              "44205",
    "Les Sorinières":           "44193",
    "Indre":                    "44074",
    "Bouaye":                   "44017",
    "Saint-Aignan-de-Grand-Lieu":"44151",
}

# Types de locaux DVF pertinents pour l'immobilier tertiaire
LOCAL_TERTIAIRE = {
    "Local industriel. commercial ou assimilé",
    "Appartement",   # parfois utilisé pour des bureaux
}

DVF_BASE_URL = "https://files.data.gouv.fr/geo-dvf/latest/csv"


def _get_dept(code_insee: str) -> str:
    """Extrait le code département depuis le code INSEE commune."""
    return code_insee[:3] if code_insee[:2] in ("97",) else code_insee[:2]


@lru_cache(maxsize=32)
def download_dvf_commune(code_insee: str) -> pd.DataFrame | None:
    """
    Télécharge le fichier CSV DVF d'une commune depuis data.gouv.fr.
    Mis en cache en mémoire pour la session (pas de re-téléchargement).

    Format : https://files.data.gouv.fr/geo-dvf/latest/csv/{dept}/communes/{code}.csv

    Returns:
        DataFrame brut ou None si indisponible.
    """
    dept = _get_dept(code_insee)
    url  = f"{DVF_BASE_URL}/{dept}/communes/{code_insee}.csv"
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        df = pd.read_csv(
            io.StringIO(r.text),
            sep=",",
            dtype=str,
            low_memory=False,
        )
        return df
    except Exception as e:
        print(f"[DVF] Échec téléchargement {url}: {e}")
        return None


def _to_float(val) -> float | None:
    """Convertit une valeur DVF en float (virgule → point)."""
    try:
        return float(str(val).replace(",", ".").replace(" ", ""))
    except (ValueError, TypeError):
        return None


def get_dvf_transactions(
    lat: float,
    lon: float,
    radius_m: float = 2000,
    commune: str = "Nantes",
    types_locaux: set | None = None,
    annees: int = 5,
) -> pd.DataFrame:
    """
    Retourne les transactions DVF dans un rayon autour d'un point.

    Args:
        lat, lon    : coordonnées du centre
        radius_m    : rayon de recherche en mètres
        commune     : nom de la commune principale (pour téléchargement ciblé)
        types_locaux: ensemble de types à inclure (défaut = tertiaire)
        annees      : nombre d'années à inclure en remontant depuis aujourd'hui

    Returns:
        DataFrame avec colonnes :
        date_mutation, adresse, type_local, surface_m2, prix_total,
        prix_m2, latitude, longitude, distance_m
    """
    if types_locaux is None:
        types_locaux = LOCAL_TERTIAIRE

    code = COMMUNE_INSEE.get(commune)
    if not code:
        # Cherche dans les communes voisines
        code = "44109"  # Nantes par défaut

    raw = download_dvf_commune(code)
    if raw is None or raw.empty:
        return pd.DataFrame()

    # Colonnes nécessaires
    needed = {
        "date_mutation", "nature_mutation",
        "valeur_fonciere", "type_local", "surface_reelle_bati",
        "latitude", "longitude",
        "adresse_numero", "adresse_nom_voie",
        "code_commune", "nom_commune",
    }
    present = [c for c in needed if c in raw.columns]
    df = raw[present].copy()

    # Filtrer : ventes uniquement
    if "nature_mutation" in df.columns:
        df = df[df["nature_mutation"].str.contains("Vente", na=False)]

    # Filtrer : types tertiaires
    if "type_local" in df.columns and types_locaux:
        df = df[df["type_local"].isin(types_locaux)]

    # Convertir lat/lon
    df["_lat"] = df["latitude"].apply(_to_float)
    df["_lon"] = df["longitude"].apply(_to_float)
    df = df.dropna(subset=["_lat", "_lon"])

    # Filtrer par date
    if "date_mutation" in df.columns:
        from datetime import datetime, timedelta
        cutoff = (datetime.now() - timedelta(days=annees * 365)).strftime("%Y-%m-%d")
        df = df[df["date_mutation"] >= cutoff]

    if df.empty:
        return pd.DataFrame()

    # Calcul distance
    def dist(row):
        R = 6371000
        p = math.pi / 180
        a = (math.sin((row["_lat"] - lat) * p / 2) ** 2
             + math.cos(lat * p) * math.cos(row["_lat"] * p)
             * math.sin((row["_lon"] - lon) * p / 2) ** 2)
        return 2 * R * math.asin(math.sqrt(a))

    df["distance_m"] = df.apply(dist, axis=1)
    df = df[df["distance_m"] <= radius_m]

    if df.empty:
        return pd.DataFrame()

    # Calculs prix
    df["prix_total"] = df["valeur_fonciere"].apply(_to_float)
    df["surface_m2"] = df["surface_reelle_bati"].apply(_to_float)

    def prix_m2(row):
        if row["prix_total"] and row["surface_m2"] and row["surface_m2"] > 0:
            return round(row["prix_total"] / row["surface_m2"], 0)
        return None

    df["prix_m2"] = df.apply(prix_m2, axis=1)

    # Adresse
    df["adresse"] = (
        df.get("adresse_numero", pd.Series(dtype=str)).fillna("").astype(str)
        + " "
        + df.get("adresse_nom_voie", pd.Series(dtype=str)).fillna("").astype(str)
    ).str.strip()

    # Colonne commune lisible
    df["commune_dvf"] = df.get("nom_commune", pd.Series(dtype=str)).fillna(commune)

    # Sélection finale
    result_cols = [
        "date_mutation", "adresse", "commune_dvf",
        "type_local", "surface_m2", "prix_total", "prix_m2",
        "_lat", "_lon", "distance_m",
    ]
    result = df[[c for c in result_cols if c in df.columns]].copy()
    result = result.sort_values("distance_m")

    # Renommage propre
    result = result.rename(columns={
        "date_mutation": "Date",
        "adresse":       "Adresse",
        "commune_dvf":   "Commune",
        "type_local":    "Type",
        "surface_m2":    "Surface (m²)",
        "prix_total":    "Prix total (€)",
        "prix_m2":       "Prix/m²",
        "distance_m":    "Distance (m)",
    })

    return result


def dvf_summary(df: pd.DataFrame) -> dict:
    """Calcule les statistiques de synthèse DVF."""
    vals = df["Prix/m²"].dropna()
    if vals.empty:
        return {}
    return {
        "n": len(df),
        "median": vals.median(),
        "mean": vals.mean(),
        "min": vals.min(),
        "max": vals.max(),
        "q1": vals.quantile(0.25),
        "q3": vals.quantile(0.75),
    }


# ─── TARGOMO ──────────────────────────────────────────────────────────────────

TARGOMO_ENDPOINT = "https://api.targomo.com/westcentraleurope/v1/polygon"

TARGOMO_COLORS = {
    5:  "#27AE60",   # vert — 5 min
    10: "#F39C12",   # orange — 10 min
    15: "#E74C3C",   # rouge — 15 min
    20: "#8E44AD",   # violet — 20 min
    30: "#2C3E50",   # gris foncé — 30 min
}

TRAVEL_MODES = {
    "walk":    "À pied",
    "bike":    "Vélo",
    "car":     "Voiture",
    "transit": "Transports en commun",
}


def fetch_isochrones(
    lat: float,
    lon: float,
    api_key: str,
    times_minutes: list[int] | None = None,
    mode: str = "walk",
) -> dict | None:
    """
    Appelle l'API Targomo pour obtenir des polygones isochrones.

    Args:
        lat, lon      : coordonnées du point d'origine
        api_key       : clé API Targomo (gratuite sur targomo.com/developers)
        times_minutes : liste de temps en minutes [5, 10, 15] par défaut
        mode          : "walk" | "bike" | "car" | "transit"

    Returns:
        GeoJSON FeatureCollection ou None si erreur.

    Documentation : https://docs.targomo.com/core/
    """
    if times_minutes is None:
        times_minutes = [5, 10, 15]

    payload = {
        "sources": [
            {
                "id":  "origin",
                "lat": lat,
                "lng": lon,
            }
        ],
        "polygon": {
            "values":           sorted(times_minutes, reverse=True),   # du plus grand au plus petit
            "intersectionMode": "union",
            "serializer":       "geojson",
            "decimalPrecision": 5,
        },
        "edgeWeight":       "time",
        "travelType":       mode,
        "maxEdgeWeight":    max(times_minutes) * 60,   # en secondes
    }

    headers = {
        "Content-Type": "application/json",
        "X-Api-Key":    api_key,
    }

    try:
        r = requests.post(
            TARGOMO_ENDPOINT,
            json=payload,
            headers=headers,
            timeout=20,
        )
        if r.status_code == 401:
            return {"error": "Clé API Targomo invalide ou expirée."}
        if r.status_code == 429:
            return {"error": "Limite de requêtes Targomo atteinte. Réessaie dans quelques secondes."}
        r.raise_for_status()
        return r.json()
    except requests.RequestException as e:
        return {"error": str(e)}


def add_isochrones_to_map(
    folium_map,
    geojson_result: dict,
    times_minutes: list[int],
    mode: str = "walk",
) -> None:
    """
    Ajoute les polygones isochrones sur une carte Folium existante.

    Args:
        folium_map    : objet folium.Map
        geojson_result: résultat de fetch_isochrones (GeoJSON)
        times_minutes : liste ordonnée des temps (même ordre que la requête)
        mode          : mode de transport (pour le label)
    """
    import folium

    if not geojson_result or "error" in geojson_result:
        return

    features = geojson_result.get("data", {}).get("features", []) \
               or geojson_result.get("features", [])

    mode_label = TRAVEL_MODES.get(mode, mode)

    for feat in features:
        props = feat.get("properties", {})
        # Targomo renvoie le temps en secondes dans "time" ou "value"
        time_sec = props.get("time") or props.get("value") or 0
        time_min = round(time_sec / 60)

        # trouver la couleur correspondante
        color = TARGOMO_COLORS.get(
            min(TARGOMO_COLORS.keys(), key=lambda t: abs(t - time_min)),
            "#888888"
        )

        tooltip = f"{mode_label} — {time_min} min"

        folium.GeoJson(
            feat,
            style_function=lambda _, c=color: {
                "fillColor":   c,
                "color":       c,
                "weight":      1.5,
                "fillOpacity": 0.15,
                "opacity":     0.6,
            },
            tooltip=tooltip,
        ).add_to(folium_map)

    # Légende isochrones
    colors_html = "".join(
        f'<span style="color:{TARGOMO_COLORS.get(t,\"#888\")};font-size:14px">●</span> '
        f'{t} min<br>'
        for t in sorted(times_minutes)
    )
    legend = f"""
    <div style="position:fixed;bottom:80px;right:24px;z-index:999;
         background:white;padding:10px 14px;border-radius:8px;
         box-shadow:0 2px 8px rgba(0,0,0,.2);font-family:Arial;font-size:12px">
      <b>Isochrones — {mode_label}</b><br>
      {colors_html}
    </div>
    """
    import folium.element
    folium_map.get_root().html.add_child(folium.element.Element(legend))
