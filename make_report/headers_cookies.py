from typing import Dict

cookies: Dict[str, str] = {
    'auth': 'd7662ab17f9a80298f34bf9a9aabbdcf53dd7e16',
    'csrf_token': '8480392f6a624caaa0df666fab79d68c',
    'evoauth': 'w09161b2059644c19a9164995a3e36914',

    'cabinet': 'company',
    'lid': '3054143',
    'social_auth': '1',
}

headers: Dict[str, str] = {
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'content-type': 'application/json; charset=UTF-8',

    'referer': 'https://my.prom.ua/cms/order/edit/', 
    'x-requested-with': 'XMLHttpRequest',
    "x-csrftoken": cookies["csrf_token"],
    'priority': 'u=1, i',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
}


