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

    # ═══════════════════════════════════════════════════════════════════
    # Écran de connexion — style Apple (glassmorphism, easing spring, orbes ambiants)
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    :root {
        --apple-ease: cubic-bezier(0.16, 1, 0.3, 1);
        --savills-teal: #008493;
        --savills-yellow: #FFDF00;
        --ink: #1D1D1F;
        --ink-soft: #6E6E73;
    }

    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, 'Inter', 'SF Pro Display',
                     'Segoe UI', Roboto, sans-serif !important;
    }

    .stApp {
        background: radial-gradient(ellipse 80% 50% at 50% -10%, #eef7f8 0%, #FAFAFA 55%, #F2F2F5 100%);
    }

    /* ── Cacher le chrome Streamlit par défaut sur l'écran de login ── */
    [data-testid="stSidebar"] { display: none; }
    #MainMenu, footer, header { visibility: hidden; }

    /* ── Orbes flottants ambiants (signature de la page) ── */
    .orb {
        position: fixed;
        border-radius: 50%;
        filter: blur(60px);
        opacity: 0.35;
        z-index: 0;
        pointer-events: none;
        animation: drift 18s ease-in-out infinite;
    }
    .orb-1 {
        width: 420px; height: 420px;
        background: radial-gradient(circle, var(--savills-teal), transparent 70%);
        top: -8%; left: -6%;
        animation-delay: 0s;
    }
    .orb-2 {
        width: 360px; height: 360px;
        background: radial-gradient(circle, var(--savills-yellow), transparent 70%);
        bottom: -10%; right: -8%;
        animation-delay: -6s;
    }
    .orb-3 {
        width: 260px; height: 260px;
        background: radial-gradient(circle, var(--savills-teal), transparent 70%);
        top: 40%; right: 10%;
        opacity: 0.18;
        animation-delay: -12s;
    }
    @keyframes drift {
        0%, 100% { transform: translate(0, 0) scale(1); }
        33%      { transform: translate(30px, -40px) scale(1.06); }
        66%      { transform: translate(-25px, 25px) scale(0.96); }
    }
    @media (prefers-reduced-motion: reduce) {
        .orb { animation: none; }
    }

    /* ── Carte de login (glassmorphism) ── */
    .login-card-wrap {
        position: relative;
        z-index: 1;
        max-width: 440px;
        margin: 5vh auto 0;
        animation: cardIn 0.7s var(--apple-ease) both;
    }
    @keyframes cardIn {
        from { opacity: 0; transform: translateY(24px) scale(0.98); }
        to   { opacity: 1; transform: translateY(0) scale(1); }
    }

    .login-glass {
        background: rgba(255, 255, 255, 0.72);
        backdrop-filter: blur(24px) saturate(180%);
        -webkit-backdrop-filter: blur(24px) saturate(180%);
        border: 1px solid rgba(255, 255, 255, 0.6);
        border-radius: 24px;
        padding: 40px 40px 8px;
        box-shadow:
            0 1px 2px rgba(0,0,0,0.04),
            0 20px 60px -12px rgba(0, 60, 70, 0.18);
    }

    .login-logo-ring {
        width: 64px; height: 64px;
        margin: 0 auto 20px;
        border-radius: 18px;
        background: linear-gradient(145deg, var(--savills-teal), #00666f);
        display: flex; align-items: center; justify-content: center;
        font-size: 28px;
        box-shadow: 0 8px 24px -6px rgba(0, 132, 147, 0.5);
        animation: logoPop 0.8s var(--apple-ease) 0.15s both;
    }
    @keyframes logoPop {
        from { opacity: 0; transform: scale(0.6) rotate(-8deg); }
        to   { opacity: 1; transform: scale(1) rotate(0); }
    }

    .login-title {
        text-align: center;
        font-size: 24px;
        font-weight: 700;
        color: var(--ink);
        letter-spacing: -0.02em;
        margin: 0 0 4px;
    }
    .login-subtitle {
        text-align: center;
        font-size: 14px;
        color: var(--ink-soft);
        margin: 0 0 28px;
        font-weight: 400;
    }

    /* ── Segmented control pour les tabs (remplace l'aspect Streamlit natif) ── */
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        background: rgba(120, 120, 128, 0.12) !important;
        border-radius: 12px !important;
        padding: 4px !important;
        gap: 2px !important;
        border: none !important;
        box-shadow: none !important;
    }
    [data-testid="stTabs"] button[data-baseweb="tab"] {
        border-radius: 9px !important;
        font-size: 13px !important;
        font-weight: 600 !important;
        color: var(--ink-soft) !important;
        transition: all 0.35s var(--apple-ease) !important;
        border: none !important;
    }
    [data-testid="stTabs"] button[aria-selected="true"] {
        background: white !important;
        color: var(--ink) !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.12) !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab-highlight"] { display: none !important; }
    [data-testid="stTabsContent"] > div { padding: 20px 0 0 !important; }

    /* ── Champs de saisie ── */
    .login-glass input {
        border-radius: 12px !important;
        border: 1px solid rgba(0,0,0,0.08) !important;
        background: rgba(255,255,255,0.8) !important;
        padding: 11px 14px !important;
        font-size: 14px !important;
        transition: all 0.25s var(--apple-ease) !important;
    }
    .login-glass input:focus {
        border-color: var(--savills-teal) !important;
        box-shadow: 0 0 0 4px rgba(0, 132, 147, 0.12) !important;
        outline: none !important;
    }

    /* ── Bouton principal ── */
    .login-glass .stFormSubmitButton button {
        background: linear-gradient(145deg, var(--savills-teal), #00747f) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        padding: 11px 0 !important;
        margin-top: 6px !important;
        transition: transform 0.25s var(--apple-ease), box-shadow 0.25s var(--apple-ease) !important;
        box-shadow: 0 4px 14px -4px rgba(0, 132, 147, 0.45) !important;
    }
    .login-glass .stFormSubmitButton button:hover {
        transform: translateY(-1px) scale(1.01) !important;
        box-shadow: 0 8px 20px -4px rgba(0, 132, 147, 0.55) !important;
    }
    .login-glass .stFormSubmitButton button:active {
        transform: translateY(0) scale(0.99) !important;
    }

    .login-footer-space { height: 24px; }
    </style>

    <div class="orb orb-1"></div>
    <div class="orb orb-2"></div>
    <div class="orb orb-3"></div>

    <div class="login-card-wrap">
      <div class="login-glass">
        <div class="login-logo-ring">🏢</div>
        <p class="login-title">Savills</p>
        <p class="login-subtitle">Outil de valorisation tertiaire</p>
    """, unsafe_allow_html=True)

    tab_login, tab_signup, tab_reset = st.tabs(["Connexion", "Créer un compte", "Mot de passe oublié"])

    with tab_login:
        with st.form("login_form"):
            email = st.text_input("Email", key="login_email", placeholder="toi@savills.fr")
            password = st.text_input("Mot de passe", type="password", key="login_pwd", placeholder="••••••••")
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
            full_name = st.text_input("Nom complet", key="signup_name", placeholder="Jean Dupont")
            email_s = st.text_input("Email", key="signup_email", placeholder="toi@savills.fr")
            pwd_s = st.text_input("Mot de passe (6 caractères min.)", type="password",
                                   key="signup_pwd", placeholder="••••••••")
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
            email_r = st.text_input("Ton email", key="reset_email", placeholder="toi@savills.fr")
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

    st.markdown("<div class='login-footer-space'></div></div></div>", unsafe_allow_html=True)

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
