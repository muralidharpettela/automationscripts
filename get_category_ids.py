
from woocommerce import API
from benedict import benedict
import json

wcapi = API(
    url="https://www.lotus-grocery.eu/",
    consumer_key="ck_a1a83db1a7931bc4c965bf0e3d281ac63bea7264",
    consumer_secret="cs_3fd453e8a22d1d3da2a1376c1d8906130f6de1e4",
    timeout=1000
)

page = 1
products = []
categories = []
category_name = []
while True:
    prods = wcapi.get('products', params={'per_page': 100, 'page': page}).json()
    page += 1
    if not prods:
        break
    products.append(prods)

for product_list_100 in products:
    for each_product in product_list_100:
        dict_you_want = benedict(each_product).subset(keys=['id', 'name', 'sku', "weight", "categories"])
        if len(dict_you_want["categories"]) == 1:
            if dict_you_want["categories"][0]["name"] not in category_name:
                category_name.append(dict_you_want["categories"][0]["name"])
                categories.append(dict_you_want["categories"])
        #all_products.append(dict_you_want)
json_string = json.dumps(categories)
with open("product_categories.json", "w") as jsonfile:
    jsonfile.write(json_string)
jsonfile.close()




