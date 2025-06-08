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
    'auth': '67a999baada1a92ed4be91d15582759533d0b4ff',
    'csrf_token': '8197d4ec000346948977d34fb6af904d',
    'evoauth': 'wbca16ef2b10b48cb86d2c225fd17404c',

    'cabinet': 'company',
    'lid': '3054143', 
    'social_auth': '1',
}
