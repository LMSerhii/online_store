import json
import math

def royalty(price, rate):
    return math.ceil((price / (100 - rate)) * 100)

def add_marg(price, rate):
    return math.ceil(price + price * (rate / 100))

def calculate_price(clear_price, marg_rate, royalty_rate):
    return royalty(add_marg(clear_price, marg_rate), royalty_rate)

def main():

    royalty_rate = 15
    marg_rate = 18
    sell_rate = 20

    items = {
        1.5: 266,
        3: 326,
        4: 416,
        5: 439,
        6: 457,
        7: 481,
        8: 497,
        10: 519,
        12: 559
    }

    for k, v in items.items():
        value = calculate_price(v,marg_rate, royalty_rate )
        items[k] = royalty(value, sell_rate)
        # items[k] = value

    with open("price.json", "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=4)


if __name__ == '__main__':
    main()
