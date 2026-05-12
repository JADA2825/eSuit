"""
auth.py — Autenticación local para eSuit.

Implementa un sistema simple de inicio de sesión con:
  - Almacenamiento en JSON con hash SHA256 + salt por usuario
  - Roles: 'admin' (gestión de usuarios) y 'usuario' (uso normal)
  - Usuario por defecto en primer arranque: admin / admin123
"""
import hashlib
import json
import secrets
from pathlib import Path
from typing import Optional

# Ruta del archivo de usuarios — relativa al script
USERS_FILE = Path(__file__).parent / "data" / "users.json"

# Usuario por defecto (se crea automáticamente la primera vez)
DEFAULT_ADMIN = {
    "user": "admin",
    "password": "admin123",
    "rol": "admin",
    "nombre": "Administrador",
}


def _ensure_data_dir() -> None:
    USERS_FILE.parent.mkdir(parents=True, exist_ok=True)


def _hash_password(password: str, salt: str) -> str:
    """SHA256 sobre salt+password (suficiente para uso interno / local)."""
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def cargar_usuarios() -> dict:
    """Carga el archivo users.json. Si no existe, crea uno con admin/admin123."""
    _ensure_data_dir()
    if not USERS_FILE.exists():
        # Bootstrap: admin por defecto
        salt = secrets.token_hex(16)
        users = {
            DEFAULT_ADMIN["user"]: {
                "salt": salt,
                "hash": _hash_password(DEFAULT_ADMIN["password"], salt),
                "rol": DEFAULT_ADMIN["rol"],
                "nombre": DEFAULT_ADMIN["nombre"],
                "default_pwd": True,
            }
        }
        _guardar(users)
        return users
    with USERS_FILE.open(encoding="utf-8") as f:
        return json.load(f)


def _guardar(users: dict) -> None:
    _ensure_data_dir()
    with USERS_FILE.open("w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def verificar(user: str, password: str) -> Optional[dict]:
    """Devuelve dict con datos del usuario si las credenciales son correctas,
       o None si no."""
    users = cargar_usuarios()
    u = users.get(user)
    if not u:
        return None
    if _hash_password(password, u["salt"]) == u["hash"]:
        return {
            "user": user,
            "rol": u.get("rol", "usuario"),
            "nombre": u.get("nombre", user),
            "default_pwd": u.get("default_pwd", False),
        }
    return None


def agregar_usuario(user: str, password: str, rol: str = "usuario",
                     nombre: str = "") -> tuple[bool, str]:
    """Crea un usuario nuevo. Retorna (ok, mensaje)."""
    user = user.strip()
    if not user:
        return False, "El nombre de usuario está vacío."
    if not password or len(password) < 4:
        return False, "La contraseña debe tener al menos 4 caracteres."
    if rol not in ("admin", "usuario"):
        return False, f"Rol inválido: {rol}"
    users = cargar_usuarios()
    if user in users:
        return False, f"El usuario '{user}' ya existe."
    salt = secrets.token_hex(16)
    users[user] = {
        "salt": salt,
        "hash": _hash_password(password, salt),
        "rol": rol,
        "nombre": nombre.strip() or user,
    }
    _guardar(users)
    return True, f"Usuario '{user}' creado con rol '{rol}'."


def eliminar_usuario(user: str) -> tuple[bool, str]:
    """Elimina un usuario. No permite eliminar al último admin."""
    users = cargar_usuarios()
    if user not in users:
        return False, f"El usuario '{user}' no existe."
    # Proteger al último admin
    admins = [u for u, d in users.items() if d.get("rol") == "admin"]
    if user in admins and len(admins) <= 1:
        return False, "No se puede eliminar al único administrador."
    del users[user]
    _guardar(users)
    return True, f"Usuario '{user}' eliminado."


def cambiar_password(user: str, nueva: str) -> tuple[bool, str]:
    """Cambia la contraseña de un usuario."""
    if not nueva or len(nueva) < 4:
        return False, "La nueva contraseña debe tener al menos 4 caracteres."
    users = cargar_usuarios()
    if user not in users:
        return False, f"El usuario '{user}' no existe."
    salt = secrets.token_hex(16)
    users[user]["salt"] = salt
    users[user]["hash"] = _hash_password(nueva, salt)
    users[user].pop("default_pwd", None)
    _guardar(users)
    return True, "Contraseña actualizada."


def cambiar_rol(user: str, nuevo_rol: str) -> tuple[bool, str]:
    """Cambia el rol. No permite degradar al último admin."""
    if nuevo_rol not in ("admin", "usuario"):
        return False, f"Rol inválido: {nuevo_rol}"
    users = cargar_usuarios()
    if user not in users:
        return False, f"El usuario '{user}' no existe."
    admins = [u for u, d in users.items() if d.get("rol") == "admin"]
    if user in admins and len(admins) <= 1 and nuevo_rol != "admin":
        return False, "No se puede quitar el rol al único administrador."
    users[user]["rol"] = nuevo_rol
    _guardar(users)
    return True, f"Rol de '{user}' actualizado a '{nuevo_rol}'."


def listar_usuarios() -> list[dict]:
    """Lista todos los usuarios (sin exponer hashes)."""
    users = cargar_usuarios()
    return [
        {
            "user": u,
            "nombre": d.get("nombre", u),
            "rol": d.get("rol", "usuario"),
            "default_pwd": d.get("default_pwd", False),
        }
        for u, d in users.items()
    ]


# ═════════════════════════════════════════════════════════
# LOGO SVG INLINE — usado en splash y header
# ═════════════════════════════════════════════════════════
LOGO_SVG = """
<svg width="120" height="120" viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="bg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0071e3"/>
      <stop offset="100%" stop-color="#1d1d1f"/>
    </linearGradient>
    <linearGradient id="bolt" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="#ffd60a"/>
      <stop offset="100%" stop-color="#ff9f0a"/>
    </linearGradient>
  </defs>
  <circle cx="100" cy="100" r="92" fill="url(#bg)" stroke="#0a84ff" stroke-width="3"/>
  <path d="M 110 30 L 60 110 L 95 110 L 80 170 L 145 85 L 110 85 Z"
        fill="url(#bolt)" stroke="#1d1d1f" stroke-width="2.5" stroke-linejoin="round"/>
  <circle cx="100" cy="100" r="92" fill="none" stroke="#ffffff"
          stroke-width="0.5" opacity="0.4"/>
</svg>
"""

LOGO_SVG_SMALL = """
<svg width="44" height="44" viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="bg2" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0071e3"/>
      <stop offset="100%" stop-color="#1d1d1f"/>
    </linearGradient>
    <linearGradient id="bolt2" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="#ffd60a"/>
      <stop offset="100%" stop-color="#ff9f0a"/>
    </linearGradient>
  </defs>
  <circle cx="100" cy="100" r="92" fill="url(#bg2)" stroke="#0a84ff" stroke-width="3"/>
  <path d="M 110 30 L 60 110 L 95 110 L 80 170 L 145 85 L 110 85 Z"
        fill="url(#bolt2)" stroke="#1d1d1f" stroke-width="2.5" stroke-linejoin="round"/>
</svg>
"""
