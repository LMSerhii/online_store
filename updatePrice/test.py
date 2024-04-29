import re

# Регулярний вираз
pattern = r'[-рсолхэпик ]\s*(\d+)'

# Рядки для тестування
strings = [
    "дропмард-685, 20450787350035",
    "дропмард-2275, 20450787082332",
    "0504371266036, олхрс760",
    "эпик, 1393",
    "1020, лендинг"
]

# Пошук числа у кожному рядку
for string in strings:
    match = re.search(pattern, string)
    if match:
        number = match.group(1)
        print(number)