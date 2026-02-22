"""
auth.py — Módulo de autenticación para FarmaAnalytics
======================================================
Usa streamlit-authenticator con contraseñas hasheadas con bcrypt.

Flujo:
  1. Lee config.yaml
  2. Si alguna contraseña NO está hasheada (no empieza con '$2b$'),
     la hashea con bcrypt y guarda el archivo actualizado.
  3. Crea y devuelve el objeto Authenticate listo para usar.
"""

import os
import yaml
import bcrypt
import streamlit as st
import streamlit_authenticator as stauth

CONFIG_FILE = "config.yaml"


# ─────────────────────────────────────────────────────────────────────────────
# Carga y hashing automático de contraseñas
# ─────────────────────────────────────────────────────────────────────────────

def _load_and_hash_config() -> dict:
    """
    Carga config.yaml y hashea con bcrypt cualquier contraseña en texto plano.

    Las contraseñas ya hasheadas (empiezan con '$2b$' o '$2a$') no se tocan.
    Si se producen cambios, el archivo se sobreescribe con los hashes nuevos.

    Returns:
        dict: Configuración completa lista para usar con streamlit-authenticator.

    Raises:
        FileNotFoundError: Si config.yaml no existe en el directorio de trabajo.
    """
    if not os.path.exists(CONFIG_FILE):
        raise FileNotFoundError(
            f"No se encontró '{CONFIG_FILE}'. "
            "Asegúrate de que el archivo está en el directorio raíz del proyecto."
        )

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    # Revisar y hashear contraseñas en texto plano
    needs_update = False
    usernames = config.get("credentials", {}).get("usernames", {})

    for username, user_data in usernames.items():
        pw = str(user_data.get("password", ""))
        if not (pw.startswith("$2b$") or pw.startswith("$2a$")):
            # Contraseña en texto plano → hashear con bcrypt (12 rounds)
            hashed = bcrypt.hashpw(
                pw.encode("utf-8"), bcrypt.gensalt(rounds=12)
            ).decode("utf-8")
            config["credentials"]["usernames"][username]["password"] = hashed
            needs_update = True

    # Guardar el config actualizado con hashes
    if needs_update:
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                yaml.dump(config, f, default_flow_style=False, allow_unicode=True)
        except IOError:
            # En sistemas de solo lectura continuamos con los hashes en memoria
            pass

    return config


# ─────────────────────────────────────────────────────────────────────────────
# Creación del objeto Authenticate
# ─────────────────────────────────────────────────────────────────────────────

def create_authenticator():
    """
    Crea y devuelve el objeto Authenticate de streamlit-authenticator.

    Returns:
        tuple: (authenticator, config)
            - authenticator: instancia de stauth.Authenticate
            - config: diccionario con la configuración completa
    """
    config = _load_and_hash_config()

    authenticator = stauth.Authenticate(
        config["credentials"],
        config["cookie"]["name"],
        config["cookie"]["key"],
        config["cookie"]["expiry_days"],
    )

    return authenticator, config


# ─────────────────────────────────────────────────────────────────────────────
# Utilidades de usuario
# ─────────────────────────────────────────────────────────────────────────────

def get_user_role(config: dict, username: str) -> str:
    """
    Devuelve el rol del usuario ('admin' o 'farmacia').

    Args:
        config: Configuración cargada del config.yaml
        username: Nombre de usuario autenticado

    Returns:
        str: Rol del usuario. Si no se encuentra, devuelve 'farmacia'.
    """
    try:
        return config["credentials"]["usernames"][username].get("role", "farmacia")
    except (KeyError, TypeError):
        return "farmacia"


# ─────────────────────────────────────────────────────────────────────────────
# Interfaz de login
# ─────────────────────────────────────────────────────────────────────────────

def render_login_ui(authenticator) -> tuple:
    """
    Renderiza la página de login con diseño oscuro profesional.

    Muestra el logo, mensaje de bienvenida y el formulario de credenciales.

    Args:
        authenticator: Objeto Authenticate de streamlit-authenticator

    Returns:
        tuple: (name, authentication_status, username)
            - name (str | None): Nombre del usuario si está autenticado
            - authentication_status (bool | None): True=ok, False=error, None=no intentado
            - username (str | None): Username si está autenticado
    """
    # ── Cabecera centrada ─────────────────────────────────────────────────
    col_l, col_c, col_r = st.columns([1, 2, 1])
    with col_c:
        st.markdown(
            """
            <div style="text-align:center; padding: 2.5rem 0 1rem 0;">
                <div style="font-size:5rem; margin-bottom:0.5rem;">💊</div>
                <h1 style="color:#2ECC71; font-size:2.8rem; margin:0; font-weight:700;">
                    FarmaAnalytics
                </h1>
                <p style="color:#BDC3C7; font-size:1.1rem; margin-top:0.5rem;">
                    Sistema profesional de análisis de ventas para farmacias
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            """
            <div style="background:#16213e; border:1px solid #2ECC71;
                        border-radius:12px; padding:1.5rem 2rem; margin:1rem 0 0.5rem 0;">
                <p style="color:#95A5A6; text-align:center; margin:0; font-size:0.95rem;">
                    Introduce tus credenciales para acceder al panel de análisis
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ── Formulario de login ───────────────────────────────────────────────
    # Intentamos la firma de la API 0.3.x primero; si falla, usamos la 0.2.x
    try:
        name, auth_status, username = authenticator.login(location="main")
    except TypeError:
        name, auth_status, username = authenticator.login("Iniciar Sesión", "main")

    # ── Pie de página de login ────────────────────────────────────────────
    st.markdown(
        """
        <p style="text-align:center; color:#4A5568; font-size:0.8rem; margin-top:2rem;">
            Los datos se procesan en memoria y no se almacenan en ningún servidor.
        </p>
        """,
        unsafe_allow_html=True,
    )

    return name, auth_status, username
