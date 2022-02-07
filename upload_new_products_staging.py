import openpyxl as xl
from os import listdir
from os.path import isfile, join
from common_functions import CommonFunctions
from woocommerce import API
import base64, requests, json
from datetime import datetime
import re
import os
import timeit

class UploadProducts(CommonFunctions):
    def __init__(self, kassen_system_filepath, new_products_excel_filepath, images_path):
        super().__init__(kassen_system_filepath, " ")
        self.images_path = images_path
        self.onlyfiles = [f for f in listdir(images_path) if isfile(join(images_path, f))]
        self.workbook = self.csv_to_excel()
        self.kassen_system_dict, row, col = self.load_kassen_system_excel_file(self.workbook)
        self.col_name, self.col_category = self.load_new_products_excel(new_products_excel_filepath)
        self.wcapi = API(
            url="https://www.staging4.lotus-grocery.eu/",
            consumer_key="ck_54f1c0d3cbc119670a8bc8cbb2a6835c0da94eda",
            consumer_secret="cs_e5e28b2e60e685c213b2ed5bcd67a5f83509fea5",
            timeout=1000
        )
        # flags for checking conditions, No image matched and no matching found in KS
        self.match_found_in_ks = False
        self.image_matched = False
        self.all_products_data_list = list()
        self.no_image_products = open("no_image_products.txt", "w+")
        self.num_no_match_found = 0
        self.match_of_stock_cells_count = 0
        self.num_no_image_match_found = 0


    def load_new_products_excel(self, new_products_excel_filepath):
        new_products_workbook = xl.load_workbook(new_products_excel_filepath)
        new_products_ws = new_products_workbook.worksheets[0]
        # column of sort id source file
        col_name = new_products_ws['A'][1:]
        # category data
        col_category = new_products_ws['B'][1:]
        return col_name, col_category

    def generate_sku(self, product_name):
        list_of_words = str(product_name.value).split()
        for j, word in enumerate(list_of_words):
            if word == "-":
                del list_of_words[j]
        final_string = "_".join(list_of_words) # SKU generate
        return final_string

    def _calculate_weight(self, product_website, products_kassen_system_dict):
        # weight calculate
        try:
            weight_value_with_unit = re.search(r'([0-9]+[" "]+(g|ml|kg|l))', str(product_website)).group(1)
            split_weight_unit = weight_value_with_unit.split(" ")
            try:
                if split_weight_unit[1] == "g":
                    weight = float(split_weight_unit[0]) / 1000
                    # weight
                    products_kassen_system_dict['weight'] = str(weight).replace(",", ".")
                elif split_weight_unit[1] == "kg":
                    weight = float(split_weight_unit[0])
                    # weight
                    products_kassen_system_dict['weight'] = str(weight).replace(",", ".")
                elif split_weight_unit[1] == "l":
                    weight = float(split_weight_unit[0])
                    # weight
                    products_kassen_system_dict['weight'] = str(weight).replace(",", ".")
                elif split_weight_unit[1] == "ml":
                    weight = float(split_weight_unit[0]) / 1000
                    # weight
                    products_kassen_system_dict['weight'] = str(weight).replace(",", ".")
                else:
                    pass
            except:
                print(str(product_website).rstrip())
        except:
            self.products_without_weight_txt.write(str(product_website))
            self.products_without_weight_txt.write("\n")

    def match_products_and_update(self, product_name, product_dict, kassen_system_data_dict):

        for j, product_kassen_system in enumerate(kassen_system_data_dict["product_names"]):
            # check the sort id source and destination are same, if yes update the stock of destination with stock of source
            if str(product_name).rstrip() == str(product_kassen_system.value).rstrip():
                self.match_found_in_ks = True
                # stock update
                product_dict['stock_quantity'] = kassen_system_data_dict["stock"][j].value
                # price update
                # products_kassen_system_dict['price'] = str(kassen_system_data_dict["price"][j].value)
                # sale price
                if kassen_system_data_dict["sale_price"][j].value != 0:
                    product_dict['sale_price'] = str(kassen_system_data_dict["price"][j].value)
                    product_dict['regular_price'] = str(
                        kassen_system_data_dict["sale_price"][j].value)
                else:
                    # regular price update
                    product_dict['regular_price'] = str(kassen_system_data_dict["price"][j].value)
                # tax class update
                if kassen_system_data_dict["tax_class"][j].value == 7:
                    product_dict['tax_class'] = "Tax 7 Per"
                else:
                    product_dict['tax_class'] = "Tax 19 Per"
                self.match_of_stock_cells_count = self.match_of_stock_cells_count + 1
                break
            if (j == len(kassen_system_data_dict["product_names"]) - 1):
                if str(product_name) not in self.no_match_products_list:
                    self.match_found_in_ks = False
                    self.no_match_products_list.append(product_name)
                    self.no_match_products_txt.write(product_name)
                    self.no_match_products_txt.write("\n")
                    self.num_no_match_found = self.num_no_match_found + 1

    def assign_category(self, category, products_dict):
        # category
        if category == "RI":
            products_dict["categories"][0]["id"] = 23
        elif category == "AT":
            products_dict["categories"][0]["id"] = 16
        elif category == "FL":
            products_dict["categories"][0]["id"] = 87
        elif category == "RA":
            products_dict["categories"][0]["id"] = 86
        elif category == "BS":
            products_dict["categories"][0]["id"] = 89
        elif category == "PS":
            products_dict["categories"][0]["id"] = 60
        elif category == "SM":
            products_dict["categories"][0]["id"] = 61
        elif category == "WS":
            products_dict["categories"][0]["id"] = 88
        elif category == "TCB":
            products_dict["categories"][0]["id"] = 67
        elif category == "CE":
            products_dict["categories"][0]["id"] = 68
        elif category == "OG":
            products_dict["categories"][0]["id"] = 90
        elif category == "GM":
            products_dict["categories"][0]["id"] = 91
        elif category == "PL":
            products_dict["categories"][0]["id"] = 92
        elif category == "PIC":
            products_dict["categories"][0]["id"] = 47
        elif category == "PAS":
            products_dict["categories"][0]["id"] = 93
        elif category == "PSY":
            products_dict["categories"][0]["id"] = 100
        elif category == "SAU":
            products_dict["categories"][0]["id"] = 54
        elif category == "CF":
            products_dict["categories"][0]["id"] = 46
        elif category == "CBC":
            products_dict["categories"][0]["id"] = 94
        elif category == "DFS":
            products_dict["categories"][0]["id"] = 28
        elif category == "SN":
            products_dict["categories"][0]["id"] = 58
        elif category == "SW":
            products_dict["categories"][0]["id"] = 62
        elif category == "PA":
            products_dict["categories"][0]["id"] = 31
        elif category == "BC":
            products_dict["categories"][0]["id"] = 96
        elif category == "HC":
            products_dict["categories"][0]["id"] = 97
        elif category == "HCO":
            products_dict["categories"][0]["id"] = 98
        elif category == "IM":
            products_dict["categories"][0]["id"] = 52
        elif category == "RE":
            products_dict["categories"][0]["id"] = 57
        elif category == "NO":
            products_dict["categories"][0]["id"] = 53
        elif category == "ID":
            products_dict["categories"][0]["id"] = 104
        elif category == "PI":
            products_dict["categories"][0]["id"] = 105
        elif category == "PM":
            products_dict["categories"][0]["id"] = 130
        elif category == "KW":
            products_dict["categories"][0]["id"] = 106
        elif category == "HI":
            products_dict["categories"][0]["id"] = 107
        else:
            products_dict["categories"][0]["id"] = "None"

    def header(self, user, password):
        credentials = user + ':' + password
        token = base64.b64encode(credentials.encode())
        header_json = {'Authorization': 'Basic ' + token.decode('utf-8')}
        return header_json

    def upload_image_to_wordpress(self, file_path, url, header_json):
        media = {'file': open(file_path, "rb"), 'caption': 'My great demo picture'}
        responce = requests.post(url + "wp-json/wp/v2/media", headers=header_json, files=media)
        newDict = responce.json()
        newID = newDict.get('id')
        link = newDict.get('guid').get("rendered")
        return link

    def upload_image_append_link(self, path, product_json):
        hed = self.header("muralidhar", "e2Yk Ba0a 3RbH vTyl PQUo WDfk")  # username, application password
        link = self.upload_image_to_wordpress(path, 'https://www.staging4.lotus-grocery.eu/', hed)
        product_json['images'][0]['src'] = link

    def find_match_upload(self, new_product_name, product_dict):
        extensions = ('.jpg', '.jpeg', '.png')
        for j, filename in enumerate(self.onlyfiles):
            if any(filename.endswith(extension) for extension in extensions):
                if os.path.splitext(filename)[0] == new_product_name.value:
                    filepath = os.path.join(self.images_path, filename)
                    print(filepath)
                    self.upload_image_append_link(filepath, product_dict)
                    self.image_matched = True
                    break
                if (j == len(self.onlyfiles) - 1):
                    self.image_matched = False
                    self.no_image_products.write(filename)
                    self.no_image_products.write("\n")
                    self.num_no_image_match_found = self.num_no_image_match_found + 1

    def process(self):
        start = timeit.default_timer()
        product_uploaded_successfully = 0
        for new_product_name, new_product_category in zip(self.col_name, self.col_category):
            product_dict = {"name": None,
                            "type": "simple",
                            "regular_price": 0.0,
                            "description": "",
                            "short_description": "",
                            "weight": "0",
                            "stock_quantity": 0,
                            "sale_price": None,
                            "tax_class": None,
                            "categories": [
                                {
                                    "id": 0
                                }
                            ],
                            "images": [
                                {
                                    "src": ""
                                }
                            ]
                            }
            product_dict["name"] = new_product_name.value
            product_dict["sku"] = self.generate_sku(new_product_name)
            # weight calculate
            self._calculate_weight(new_product_name.value, product_dict)
            # regular price, sale price, stock, tax class
            self.match_products_and_update(new_product_name.value, product_dict, self.kassen_system_dict)
            self.assign_category(new_product_category.value, product_dict)
            # image
            self.find_match_upload(new_product_name, product_dict)
            if self.match_found_in_ks == True and self.image_matched == True:
                #print(self.wcapi.post("products", product_dict).json())
                product_uploaded_successfully = product_uploaded_successfully + 1
                self.all_products_data_list.append(product_dict)
        self.no_match_products_txt.close()
        self.no_image_products.close()
        MAX_API_BATCH_SIZE = 50

        def chunks(l, n):
            """Yield successive n-sized chunks from l."""
            for i in range(0, len(l), n):
                yield l[i:i + n]

        try:
            for batch in chunks(self.all_products_data_list, MAX_API_BATCH_SIZE):
                print(len(batch))
                print(self.wcapi.put("products/batch", {"create": batch}).json())
            subject = '[Staging] lotus-grocery.eu - new products uploaded ' + datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S")
            prod = [prod['name'] for prod in self.all_products_data_list]
            message = "This is an automated mail, receives this mail, when new products are uploaded. " \
                      "The statistics are as follows:\n\n\n" \
                      "Total no of Products target to upload:{}\n " \
                      "Total no of Products uploaded successfully:{}\n" \
                      "Number of Products Matched their names in KS:{}\n" \
                      "Number of Products not Matched their names in KS:{}\n" \
                      "Number of Products images are not matched:{}\n" \
                      "Time elasped:{} seconds\n"\
                      "Uploaded products are:{}".format(len(self.col_name), product_uploaded_successfully, self.match_of_stock_cells_count,
                                                       self.num_no_match_found, self.num_no_image_match_found,
                                                       timeit.default_timer() - start, "\n".join(prod))
            content = [message, "./no_match_products.txt", "./no_image_products.txt"]
            self.send_email(subject, content)
        except:
            subject = '[Staging] lotus-grocery.eu - Stock Updated not successfully on ' + datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S")
            message = "products update not successful. Please re-run the scripts again"
            content = [message]
            self.send_email(subject, content)
        print("Total no of Products target to upload:{}".format(len(self.col_name)))
        print("Total no of Products uploaded successfully:{}".format(product_uploaded_successfully))
        print("Number of Products Matched their names in KS:{}".format(self.match_of_stock_cells_count))
        print("Number of Products not Matched their names in KS:{}".format(self.num_no_match_found))
        print("Number of Products images are not matched:{}".format(self.num_no_image_match_found))

        stop = timeit.default_timer()

        print('Time: ', stop - start)
        print("Finished")


if __name__ == "__main__":
    filepath_kassen_system = r"/Users/muralidharpettela/Downloads/BK_Artikeldaten_07022022.csv"
    new_products_excel_path = r"/Users/muralidharpettela/Downloads/07022022.xlsx"
    images_path = r"/Users/muralidharpettela/Downloads/07022022_images"
    staging_products_update = UploadProducts(filepath_kassen_system, new_products_excel_path, images_path)
    staging_products_update.process()
