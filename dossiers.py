"""
Module dossiers.py — Persistance des dossiers de valorisation (Phase 2)
─────────────────────────────────────────────────────────────────────
Un "dossier" = l'état de travail complet sur un actif :
adresse, critères de recherche, comparables trouvés, résultats DVF, etc.
Stocké dans la table `dossiers` (colonne `data` en JSONB) créée par
migration_phase1.sql.

Toutes les fonctions retournent None en cas d'échec réseau/DB plutôt que
de lever une exception — l'app doit rester utilisable même si la sauvegarde
échoue ponctuellement (perte de service dégradée, pas de crash).
"""

from __future__ import annotations

import json
import math
from datetime import datetime, timezone

import pandas as pd

import auth


# ─────────────────────────────────────────────────────────────────────────
def _client():
    return auth.get_supabase_client()


def _clean_for_json(obj):
    """Nettoie récursivement les valeurs non-JSON-sérialisables
    (NaN, Infinity, numpy/pandas types) avant stockage en JSONB."""
    if isinstance(obj, dict):
        return {k: _clean_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_clean_for_json(v) for v in obj]
    if isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return None
        return obj
    return obj


# ─────────────────────────────────────────────────────────────────────────
def create_dossier(user_id: str, nom_client: str = "", adresse: str = "",
                   type_bien: str = "Bureaux") -> dict | None:
    """Crée un nouveau dossier vide et le retourne (avec son id)."""
    sb = _client()
    if not sb:
        return None
    try:
        res = sb.table("dossiers").insert({
            "user_id":    user_id,
            "nom_client": nom_client or "Sans nom",
            "adresse":    adresse,
            "type_bien":  type_bien,
            "statut":     "brouillon",
            "data":       {},
        }).execute()
        return res.data[0] if res.data else None
    except Exception as e:
        print(f"[dossiers] create_dossier échec: {e}")
        return None


def list_dossiers(user_id: str) -> list[dict]:
    """Liste tous les dossiers d'un utilisateur, du plus récent au plus ancien."""
    sb = _client()
    if not sb:
        return []
    try:
        res = (sb.table("dossiers")
               .select("id, nom_client, adresse, type_bien, statut, created_at, updated_at")
               .eq("user_id", user_id)
               .order("updated_at", desc=True)
               .execute())
        return res.data or []
    except Exception as e:
        print(f"[dossiers] list_dossiers échec: {e}")
        return []


def load_dossier(dossier_id: str) -> dict | None:
    """Charge un dossier complet (avec son contenu JSONB data)."""
    sb = _client()
    if not sb:
        return None
    try:
        res = sb.table("dossiers").select("*").eq("id", dossier_id).single().execute()
        return res.data
    except Exception as e:
        print(f"[dossiers] load_dossier échec: {e}")
        return None


def save_dossier_data(dossier_id: str, data: dict, nom_client: str | None = None,
                      adresse: str | None = None, statut: str | None = None) -> bool:
    """Sauvegarde (autosave) le contenu d'un dossier. Retourne True si succès."""
    sb = _client()
    if not sb:
        return False
    payload = {
        "data":       _clean_for_json(data),
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    if nom_client is not None:
        payload["nom_client"] = nom_client
    if adresse is not None:
        payload["adresse"] = adresse
    if statut is not None:
        payload["statut"] = statut
    try:
        sb.table("dossiers").update(payload).eq("id", dossier_id).execute()
        return True
    except Exception as e:
        print(f"[dossiers] save_dossier_data échec: {e}")
        return False


def delete_dossier(dossier_id: str) -> bool:
    sb = _client()
    if not sb:
        return False
    try:
        sb.table("dossiers").delete().eq("id", dossier_id).execute()
        return True
    except Exception as e:
        print(f"[dossiers] delete_dossier échec: {e}")
        return False


def duplicate_dossier(dossier_id: str, user_id: str) -> dict | None:
    """Duplique un dossier existant (utile pour comparer des variantes)."""
    original = load_dossier(dossier_id)
    if not original:
        return None
    sb = _client()
    if not sb:
        return None
    try:
        res = sb.table("dossiers").insert({
            "user_id":    user_id,
            "nom_client": (original.get("nom_client") or "") + " (copie)",
            "adresse":    original.get("adresse", ""),
            "type_bien":  original.get("type_bien", "Bureaux"),
            "statut":     "brouillon",
            "data":       original.get("data", {}),
        }).execute()
        return res.data[0] if res.data else None
    except Exception as e:
        print(f"[dossiers] duplicate_dossier échec: {e}")
        return None


# ─────────────────────────────────────────────────────────────────────────
# Sérialisation de l'état applicatif ↔ dict JSONB
# ─────────────────────────────────────────────────────────────────────────

def snapshot_session_state(st_session_state) -> dict:
    """Capture l'état pertinent de la session Streamlit pour sauvegarde.
    Ne capture QUE les données sérialisables et utiles à restaurer —
    pas les objets Streamlit/Folium/clients API."""
    snap = {}

    if "live_results" in st_session_state:
        snap["live_results"] = _clean_for_json(st_session_state["live_results"])

    if "dvf_results" in st_session_state:
        dvf = st_session_state["dvf_results"]
        if isinstance(dvf, pd.DataFrame) and not dvf.empty:
            snap["dvf_results"] = _clean_for_json(dvf.to_dict("records"))

    # Champs de saisie utilisateur qu'on veut restaurer
    for key in ("address_value", "manual_commune_value", "manual_cp_value",
                "op_value", "asset_type_value", "radius_km_value"):
        if key in st_session_state:
            snap[key] = st_session_state[key]

    snap["_saved_at"] = datetime.now(timezone.utc).isoformat()
    return snap


def restore_session_state(snapshot: dict, st_session_state) -> None:
    """Réinjecte un snapshot sauvegardé dans st.session_state au chargement
    d'un dossier existant."""
    if not snapshot:
        return
    if "live_results" in snapshot:
        st_session_state["live_results"] = snapshot["live_results"]
    if "dvf_results" in snapshot:
        st_session_state["dvf_results"] = pd.DataFrame(snapshot["dvf_results"])
    for key in ("address_value", "manual_commune_value", "manual_cp_value",
                "op_value", "asset_type_value", "radius_km_value"):
        if key in snapshot:
            st_session_state[key] = snapshot[key]
