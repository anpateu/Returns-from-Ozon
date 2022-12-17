import os

from dotenv import load_dotenv

load_dotenv()


class Config:
    OZON_CLIENT_ID = os.getenv('OZON_CLIENT_ID')
    OZON_API_KEY = os.getenv('OZON_API_KEY')