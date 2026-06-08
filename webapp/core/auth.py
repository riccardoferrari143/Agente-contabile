"""
Autenticazione: registrazione, login, sessione Streamlit
"""
import bcrypt, streamlit as st
from core.database import crea_utente, get_utente_by_email, init_db

def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()

def verify_password(password: str, hashed: str) -> bool:
    return bcrypt.checkpw(password.encode(), hashed.encode())

def registra(email, nome, password):
    if len(password) < 8:
        return False, "La password deve essere di almeno 8 caratteri."
    if "@" not in email:
        return False, "Email non valida."
    h = hash_password(password)
    return crea_utente(email, nome, h)

def login(email, password):
    utente = get_utente_by_email(email)
    if not utente:
        return False, "Email o password errati."
    if not verify_password(password, utente["password_hash"]):
        return False, "Email o password errati."
    return True, utente

def init_session():
    """Inizializza le chiavi di sessione necessarie."""
    for key, default in [
        ("logged_in", False), ("utente", None), ("azienda_id", None),
        ("azienda_nome", ""), ("page", "dashboard")
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

def require_login():
    """Blocca la pagina se l'utente non è loggato."""
    init_session()
    if not st.session_state.logged_in:
        st.warning("Accedi per continuare.")
        st.stop()

def logout():
    for key in ["logged_in","utente","azienda_id","azienda_nome"]:
        st.session_state[key] = None if key!="logged_in" else False
    st.rerun()
