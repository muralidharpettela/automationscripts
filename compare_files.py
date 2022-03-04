from update_products_data_staging import csv_to_excel, load_json_data_website_products, load_kassen_system_excel_file, assign_data_from_ks

filepath_kassen_system = r"/Users/muralidharpettela/Downloads/BK_Artikeldaten_04032022_1.csv"
json_file_path = "products.json"
no_match_products_list = list()
no_match_products_txt = open("no_match_products.txt", "w+")


def match_products_and_update(json_data_dict, kassen_system_data_dict):
    num_no_match_found = 0
    match_of_stock_cells_count = 0
    for i, product_kassen_system in enumerate(kassen_system_data_dict["product_names"]):
        for j, product_website in enumerate(json_data_dict):
            # check the sort id source and destination are same, if yes update the stock of destination with stock of source
            if str(product_website['name']).rstrip() == str(product_kassen_system.value).rstrip():

                match_of_stock_cells_count = match_of_stock_cells_count + 1
                break

            if (j == len(json_data_dict) - 1):
                if str(str(product_kassen_system.value).rstrip()) not in no_match_products_list:
                    no_match_products_list.append(str(product_kassen_system.value).rstrip())
                    no_match_products_txt.write(str(product_kassen_system.value).rstrip())
                    no_match_products_txt.write("\n")
                    num_no_match_found = num_no_match_found + 1
    no_match_products_txt.close()
    return match_of_stock_cells_count, num_no_match_found


def main():
    path_of_excel = csv_to_excel(filepath_kassen_system)
    ws1, mr_s, mc_s = load_kassen_system_excel_file(path_of_excel)
    json_data = load_json_data_website_products(json_file_path)
    kassen_system_data = assign_data_from_ks(ws1)
    match_of_stock_cells_count, num_no_match_found = match_products_and_update(json_data, kassen_system_data)
    print()





if __name__ == "__main__":
    main()