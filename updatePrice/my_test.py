import json 



with open('prom_rate.json', 'r', encoding='utf-8') as file:
        f = json.load(file)

for item in f:
        print(type(item.get('cat_id')))

# print(f)