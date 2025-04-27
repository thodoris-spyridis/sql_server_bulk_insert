import os
from dotenv import load_dotenv

load_dotenv(override=True)

SERVER = os.getenv("server")
DATABASE = os.getenv("database")
USERNAME = os.getenv("username")
PW = os.getenv("pw")

__all__ = ["SERVER", "DATABASE", "USERNAME", "PW"]