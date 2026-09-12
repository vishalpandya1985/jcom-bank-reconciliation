import os
import secrets
import bcrypt
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired

SECRET_KEY_PATH = os.path.join(os.environ.get('DATA_DIR', os.path.join(os.path.dirname(__file__), 'data')), 'secret.key')


def _get_secret_key():
    os.makedirs(os.path.dirname(SECRET_KEY_PATH), exist_ok=True)
    if os.path.exists(SECRET_KEY_PATH):
        with open(SECRET_KEY_PATH) as f:
            return f.read().strip()
    key = secrets.token_hex(32)
    with open(SECRET_KEY_PATH, 'w') as f:
        f.write(key)
    return key


SECRET_KEY = _get_secret_key()
serializer = URLSafeTimedSerializer(SECRET_KEY)

SESSION_COOKIE_NAME = 'recon_session'
SESSION_MAX_AGE = 60 * 60 * 24 * 14  # 14 days


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode('utf-8')[:72], bcrypt.gensalt()).decode('utf-8')


def verify_password(password: str, password_hash: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode('utf-8')[:72], password_hash.encode('utf-8'))
    except Exception:
        return False


def create_session_token(user_id: int) -> str:
    return serializer.dumps({'user_id': user_id})


def read_session_token(token: str):
    if not token:
        return None
    try:
        data = serializer.loads(token, max_age=SESSION_MAX_AGE)
        return data.get('user_id')
    except (BadSignature, SignatureExpired):
        return None
