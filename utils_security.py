import base64
import os

# Scrambled keys (Base64 encoded)
# These act as an internal fallback if environment variables are not set.
_G = "QUl6YVN5QUJoMlV3WURWRWRYblo1bkRMVjdxQ1Z2WjlVR2FDS2hz"
_T = "ODMxMTE1MDMwMjpBQUdGZVIzN0NuOU5ZanlJcFdaOVFWaTJEaFhOTWc3U1dBMA=="

def get_gemini_key():
    return os.getenv("GEMINI_API_KEY") or base64.b64decode(_G).decode()

def get_telegram_token():
    return os.getenv("TELEGRAM_TOKEN") or base64.b64decode(_T).decode()
