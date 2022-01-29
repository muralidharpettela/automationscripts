
from woocommerce import API
from benedict import benedict

wcapi = API(
    url="https://www.lotus-grocery.eu/",
    consumer_key="ck_a1a83db1a7931bc4c965bf0e3d281ac63bea7264",
    consumer_secret="cs_3fd453e8a22d1d3da2a1376c1d8906130f6de1e4",
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




