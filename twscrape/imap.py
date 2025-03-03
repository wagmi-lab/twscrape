import asyncio
import email as emaillib
import imaplib
import os
import time
from datetime import datetime
import msal


from .logger import logger


def env_int(key: str | list[str], default: int) -> int:
    key = [key] if isinstance(key, str) else key
    val = [os.getenv(k) for k in key]
    val = [int(x) for x in val if x is not None]
    return val[0] if val else default


TWS_WAIT_EMAIL_CODE = env_int(["TWS_WAIT_EMAIL_CODE", "LOGIN_CODE_TIMEOUT"], 30)

CLIENT_SECRET = os.getenv("CLIENT_SECRET", "")
CLIENT_ID = os.getenv("CLIENT_ID", "")
TENANT_ID = os.getenv("TENANT_ID", "")

if CLIENT_SECRET == "" or CLIENT_ID == "" or TENANT_ID == "":
    raise Exception("Client secret, client id, or tenant id not set")

AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPES = ['https://outlook.office365.com/.default']

# Create a public client application
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET,
)


class EmailLoginError(Exception):
    def __init__(self, message="Email login error"):
        self.message = message
        super().__init__(self.message)


class EmailCodeTimeoutError(Exception):
    def __init__(self, message="Email code timeout"):
        self.message = message
        super().__init__(self.message)


IMAP_MAPPING: dict[str, str] = {
    "yahoo.com": "imap.mail.yahoo.com",
    "icloud.com": "imap.mail.me.com",
    "outlook.com": "imap-mail.outlook.com",
    "hotmail.com": "imap-mail.outlook.com",
}


def add_imap_mapping(email_domain: str, imap_domain: str):
    IMAP_MAPPING[email_domain] = imap_domain


def _get_imap_domain(email: str) -> str:
    email_domain = email.split("@")[1]
    if email_domain in IMAP_MAPPING:
        return IMAP_MAPPING[email_domain]
    return f"imap.{email_domain}"


def _wait_email_code(imap: imaplib.IMAP4_SSL, count: int, min_t: datetime | None, to_email:str) -> str | None:
    for i in range(count, 0, -1):
        _, rep = imap.fetch(str(i), "(RFC822)")
        for x in rep:
            if isinstance(x, tuple):
                msg = emaillib.message_from_bytes(x[1])

                # https://www.ietf.org/rfc/rfc9051.html#section-6.3.12-13
                msg_time = msg.get("Date", "").split("(")[0].strip()
                msg_time = datetime.strptime(msg_time, "%a, %d %b %Y %H:%M:%S %z")

                msg_from = str(msg.get("From", "")).lower()
                msg_subj = str(msg.get("Subject", "")).lower()
                msg_to = str(msg.get("To", "")).lower()
                logger.info(f"({i} of {count}) {msg_from} - {msg_to} - {msg_time} - {msg_subj}")

                if min_t is not None and msg_time < min_t:
                    return None

                if "info@x.com" in msg_from and "confirmation code is" in msg_subj and to_email in msg_to:
                    # eg. Your Twitter confirmation code is XXX
                    return msg_subj.split(" ")[-1].strip()

    return None


async def imap_get_email_code(
    imap: imaplib.IMAP4_SSL, email: str, min_t: datetime | None = None
) -> str:
    try:
        logger.info(f"Waiting for confirmation code for {email}...")
        start_time = time.time()
        while True:
            _, rep = imap.select("INBOX")
            msg_count = int(rep[0].decode("utf-8")) if len(rep) > 0 and rep[0] is not None else 0
            code = _wait_email_code(imap, msg_count, min_t, email)
            if code is not None:
                return code

            if TWS_WAIT_EMAIL_CODE < time.time() - start_time:
                raise EmailCodeTimeoutError(f"Email code timeout ({TWS_WAIT_EMAIL_CODE} sec)")

            await asyncio.sleep(5)
    except Exception as e:
        imap.select("INBOX")
        imap.close()
        raise e


def _checking_email(receiver_email:str=None,sender_email:str=None):
    def get_token():
        # Acquire a token using client credentials
        result = app.acquire_token_silent(SCOPES, account=None)
        if not result:
            # If no token is found in the cache, acquire a new one
            result = app.acquire_token_for_client(scopes=SCOPES)
        return result

    def get_access_token():
        token_response = get_token()
        access_token = token_response['access_token']
        return access_token

    def check_token_expiry(token_response):
        # Check if the access token is about to expire in the next 5 minutes
        expiry_time = token_response['expires_in']
        current_time = time.time()
        if expiry_time - current_time < 300:
            return True
        return False

    token_response = get_token()
    access_token = token_response['access_token']

    # Check if the token is about to expire
    if check_token_expiry(token_response):
        token_response = get_token()
        access_token = token_response['access_token']
        
        
def generate_token():
    def get_token():
        # Acquire a token using client credentials
        result = app.acquire_token_silent(SCOPES, account=None)
        if not result:
            # If no token is found in the cache, acquire a new one
            result = app.acquire_token_for_client(scopes=SCOPES)
        return result

    def get_access_token():
        token_response = get_token()
        access_token = token_response['access_token']
        return access_token

    def check_token_expiry(token_response):
        # Check if the access token is about to expire in the next 5 minutes
        expiry_time = token_response['expires_in']
        current_time = time.time()
        if expiry_time - current_time < 300:
            return True
        return False

    token_response = get_token()
    access_token = token_response['access_token']

    # Check if the token is about to expire
    if check_token_expiry(token_response):
        token_response = get_token()
        access_token = token_response['access_token']
        
    return access_token

def generate_auth_string(user, token):
    auth_string = f"user={user}\1auth=Bearer {token}\1\1"
    return auth_string

async def imap_login(email: str, password:str):
    # domain = _get_imap_domain(email)
    domain = "outlook.office365.com"
    
    imap_email = os.getenv("IMAP_EMAIL", "")
    if imap_email == "":
        raise Exception("IMAP_EMAIL not set")
    
    token = generate_token()
    imap = imaplib.IMAP4_SSL(domain)

    try:
        # imap.login(email, password)
        imap.authenticate("XOAUTH2", lambda x:generate_auth_string(imap_email, token))
        imap.select("INBOX", readonly=True)
    except imaplib.IMAP4.error as e:
        logger.error(f"Error logging into {imap_email} on {domain}: {e}")
        raise EmailLoginError() from e

    return imap

