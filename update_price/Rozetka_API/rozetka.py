import os
import base64

import requests

from dotenv import load_dotenv

load_dotenv()

USERNAME = os.getenv('username_rozetka')
PASSWORD = os.getenv('password_rozetka')


# print(base64.b64encode(PASSWORD.encode('ascii')))

params = {
    'username': USERNAME,
    'password': base64.b64encode(PASSWORD.encode('ascii'))
}

response = requests.get(url='https://api-seller.rozetka.com.ua/sites', params=params).json()

print(response)
