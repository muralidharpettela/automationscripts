
from woocommerce import API
from benedict import benedict
import json


def load_wp_credentials(json_file_path):
    # load the json file
    # Opening JSON file
    f = open(json_file_path)
    # returns JSON object as
    # a dictionary
    data = json.load(f)
    return data


credentials = load_wp_credentials("/common/wp_credentials_live.json")
wcapi = API(
    url=credentials["url"],
    consumer_key=credentials["consumer_key"],
    consumer_secret=credentials["consumer_secret"],
    timeout=1000
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
        dict_you_want = benedict(each_product).subset(keys=['id', 'name', 'sku', 'stock_quantity'])
        if dict_you_want['stock_quantity'] == 0:
            all_products.append(dict_you_want)

with open("products_out_of_stock.txt", "w+") as txt:
    for product in all_products:
        txt.write(product['name'])
        txt.write("\n")
    txt.close()




