
from woocommerce import API
from benedict import benedict
import json
wcapi = API(
    url="https://www.staging4.lotus-grocery.eu/",
    consumer_key="ck_ae5271d4aef30767be04eca894f5289b24ddebd6",
    consumer_secret="cs_7fcb2ea53e0cdce73247f4f3b91ef0a5aea1f6eb",
    timeout=30
)

page = 1
products = []
while True:
    prods = wcapi.get('products', params={'per_page': 100, 'page': page}).json()
    page += 1
    if not prods:
        break
    products.append(prods)
all_products = list()
for product_list_100 in products:
    for each_product in product_list_100:
        dict_you_want = benedict(each_product).subset(keys=['id', 'name', 'sku'])
        all_products.append(dict_you_want)
json_string = json.dumps(all_products)
with open("products.json", "w") as jsonfile:
    jsonfile.write(json_string)
jsonfile.close()




