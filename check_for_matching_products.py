from common_functions import CommonFunctions
import timeit


class CheckForMatchingProducts(CommonFunctions):
    def __init__(self, filepath_kassen_system, json_file_path="products.json"):
        super().__init__(filepath_kassen_system, json_file_path)

    def process(self):
        start = timeit.default_timer()
        workbook = self.csv_to_excel()
        kassen_system_data_dict, mr_s, mc_s = self.load_kassen_system_excel_file(workbook)
        json_data = self.load_json_data_website_products()
        weight_updated_products, match_of_stock_cells_count, num_no_match_found = self.match_products_and_update(json_data, kassen_system_data_dict)
        print("Total no of Rows/Products in Source file from Shop File:{}".format(mr_s))
        print("Total no of Rows/Products in Destination file in Website:{}".format(len(json_data)))
        print("Number of Products Matched:{}".format(match_of_stock_cells_count))
        print("Number of Products are no matched:{}".format(num_no_match_found))
        print("Number of Products Weights updated:{}".format(weight_updated_products))
        stop = timeit.default_timer()

        print('Time: ', stop - start)


if __name__ == "__main__":
    filepath_kassen_system = r"/Users/muralidharpettela/Downloads/BK_Artikeldaten_09032022.csv"
    live_products_update = CheckForMatchingProducts(filepath_kassen_system)
    live_products_update.process()


