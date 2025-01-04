from typing import Dict

headers: Dict[str, str] = {
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'content-type': 'application/json; charset=UTF-8',

    'referer': 'https://my.prom.ua/cms/order/edit/', 
    'x-requested-with': 'XMLHttpRequest',
    'priority': 'u=1, i',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
}

cookies: Dict[str, str] = {
    'auth': '7d47231bc274b63fcb8a58aa0e5fa42ef9a4eb51',
    'csrf_token': 'bf8bfdaa1e744bdbbe0e378a176e4e51',
    'evoauth': 'w33bd016ca7544eb893812febd46e64a8',

    'cabinet': 'company',
    'lid': '3054143', 
    'social_auth': '1',
}
