"""
Module auth.py — Authentification Supabase pour l'app Streamlit
──────────────────────────────────────────────────────────────
Gère : inscription, connexion, déconnexion, récupération mot de passe.
La session utilisateur vit dans st.session_state pendant la durée de
la session navigateur (rechargement d'onglet = reconnexion nécessaire,
comportement standard sans gestion de cookies persistants).

Configuration requise (variables d'environnement, JAMAIS en dur dans le code) :
    SUPABASE_URL       — Project URL (Settings → API)
    SUPABASE_ANON_KEY  — anon public key (Settings → API)

Sur Render : Dashboard → ton service → Environment → Add Environment Variable
En local   : export SUPABASE_URL=... / export SUPABASE_ANON_KEY=...
"""

from __future__ import annotations

import os

import streamlit as st

try:
    from supabase import create_client, Client
except ImportError:
    create_client = None
    Client = None


# ─────────────────────────────────────────────────────────────────────────
def get_supabase_client():
    """Retourne un client Supabase mis en cache, ou None si non configuré."""
    if create_client is None:
        return None
    url = os.environ.get("SUPABASE_URL", "").strip()
    key = os.environ.get("SUPABASE_ANON_KEY", "").strip()
    if not url or not key:
        return None
    if "_supabase_client" not in st.session_state:
        st.session_state._supabase_client = create_client(url, key)
    return st.session_state._supabase_client


def is_configured() -> bool:
    """True si les variables d'environnement Supabase sont présentes."""
    return bool(
        os.environ.get("SUPABASE_URL", "").strip()
        and os.environ.get("SUPABASE_ANON_KEY", "").strip()
        and create_client is not None
    )


def current_user() -> dict | None:
    """Retourne {id, email} si connecté, sinon None."""
    return st.session_state.get("auth_user")


def is_logged_in() -> bool:
    return current_user() is not None


def logout():
    sb = get_supabase_client()
    if sb:
        try:
            sb.auth.sign_out()
        except Exception:
            pass
    for k in ("auth_user", "_supabase_client"):
        st.session_state.pop(k, None)
    st.rerun()


# ─────────────────────────────────────────────────────────────────────────
def _do_login(email: str, password: str) -> str | None:
    """Tente une connexion. Retourne un message d'erreur ou None si succès."""
    sb = get_supabase_client()
    if not sb:
        return "Supabase non configuré (variables d'environnement manquantes)."
    try:
        res = sb.auth.sign_in_with_password({"email": email, "password": password})
        if res.user:
            st.session_state.auth_user = {"id": res.user.id, "email": res.user.email}
            return None
        return "Identifiants incorrects."
    except Exception as e:
        msg = str(e)
        if "Invalid login credentials" in msg:
            return "Email ou mot de passe incorrect."
        return f"Erreur de connexion : {msg}"


def _do_signup(email: str, password: str, full_name: str) -> str | None:
    """Tente une inscription. Retourne un message d'erreur, ou 'CONFIRM_EMAIL' si succès."""
    sb = get_supabase_client()
    if not sb:
        return "Supabase non configuré (variables d'environnement manquantes)."
    if len(password) < 6:
        return "Le mot de passe doit contenir au moins 6 caractères."
    try:
        res = sb.auth.sign_up({
            "email": email,
            "password": password,
            "options": {"data": {"full_name": full_name}},
        })
        if res.user:
            return "CONFIRM_EMAIL"
        return "Erreur inconnue lors de l'inscription."
    except Exception as e:
        msg = str(e)
        if "already registered" in msg.lower():
            return "Cet email est déjà utilisé. Connecte-toi plutôt."
        return f"Erreur d'inscription : {msg}"


def _do_reset(email: str) -> str:
    sb = get_supabase_client()
    if not sb:
        return "Supabase non configuré."
    try:
        sb.auth.reset_password_for_email(email)
        return "OK"
    except Exception as e:
        return f"Erreur : {e}"


# ─────────────────────────────────────────────────────────────────────────
def require_login():
    """
    Point d'entrée principal. À appeler tout en haut de app.py, juste après
    st.set_page_config(). Affiche un écran de connexion/inscription et
    interrompt l'exécution (st.stop()) tant que l'utilisateur n'est pas
    authentifié. Si déjà connecté, ne fait rien et laisse le script continuer.
    """
    if is_logged_in():
        return

    if not is_configured():
        st.error(
            "⚠️ Authentification non configurée. Ajoute les variables d'environnement "
            "`SUPABASE_URL` et `SUPABASE_ANON_KEY` dans les réglages Render "
            "(Settings → Environment) puis redéploie."
        )
        st.stop()

    st.markdown(
        "<div style='max-width:420px;margin:60px auto 20px;padding:0 16px'>"
        "<h2 style='text-align:center;color:#25273A;margin-bottom:4px'>🏢 Savills</h2>"
        "<p style='text-align:center;color:#79828C;font-size:13px;margin-bottom:8px'>"
        "Outil de valorisation tertiaire</p></div>",
        unsafe_allow_html=True,
    )

    tab_login, tab_signup, tab_reset = st.tabs(["Connexion", "Créer un compte", "Mot de passe oublié"])

    with tab_login:
        with st.form("login_form"):
            email = st.text_input("Email", key="login_email")
            password = st.text_input("Mot de passe", type="password", key="login_pwd")
            submitted = st.form_submit_button("Se connecter", use_container_width=True)
        if submitted:
            if not email or not password:
                st.warning("Renseigne email et mot de passe.")
            else:
                err = _do_login(email, password)
                if err:
                    st.error(err)
                else:
                    st.rerun()

    with tab_signup:
        with st.form("signup_form"):
            full_name = st.text_input("Nom complet", key="signup_name")
            email_s = st.text_input("Email", key="signup_email")
            pwd_s = st.text_input("Mot de passe (6 caractères min.)", type="password", key="signup_pwd")
            submitted_s = st.form_submit_button("Créer mon compte", use_container_width=True)
        if submitted_s:
            if not email_s or not pwd_s:
                st.warning("Renseigne au moins email et mot de passe.")
            else:
                err = _do_signup(email_s, pwd_s, full_name)
                if err == "CONFIRM_EMAIL":
                    st.success(
                        "✅ Compte créé ! Vérifie ta boîte mail pour confirmer ton adresse, "
                        "puis reviens te connecter ici."
                    )
                elif err:
                    st.error(err)

    with tab_reset:
        with st.form("reset_form"):
            email_r = st.text_input("Ton email", key="reset_email")
            submitted_r = st.form_submit_button("Envoyer le lien de réinitialisation", use_container_width=True)
        if submitted_r:
            if not email_r:
                st.warning("Renseigne ton email.")
            else:
                res = _do_reset(email_r)
                if res == "OK":
                    st.success("📧 Email envoyé si ce compte existe. Vérifie ta boîte mail.")
                else:
                    st.error(res)

    st.stop()


def render_user_badge():
    """Affiche l'utilisateur connecté + bouton déconnexion dans la sidebar.
    À appeler dans la sidebar de app.py, après require_login()."""
    user = current_user()
    if not user:
        return
    with st.sidebar:
        st.markdown(
            f"<div style='background:#2f3145;border-radius:8px;padding:8px 12px;"
            f"margin-bottom:12px;font-size:12px;color:#EEE8E3'>"
            f"👤 {user['email']}</div>",
            unsafe_allow_html=True,
        )
        if st.button("🚪 Déconnexion", use_container_width=True):
            logout()
