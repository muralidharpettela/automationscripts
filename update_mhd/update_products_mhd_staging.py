from common.common_functions import CommonFunctions
import timeit
from woocommerce import API
from datetime import datetime
import openpyxl as xl


class StagingUpdateProductsMHD(CommonFunctions):
    def __init__(self, filepath_kassen_system, json_file_path="products.json"):
        super().__init__(filepath_kassen_system, json_file_path)
        credentials = self.load_wp_credentials("/common/wp_credentials_staging.json")
        self.wcapi = API(
            url=credentials["url"],
            consumer_key=credentials["consumer_key"],
            consumer_secret=credentials["consumer_secret"],
            timeout=1000
        )

    def process(self):
        start = timeit.default_timer()
        #workbook = self.csv_to_excel()
        workbook = xl.load_workbook(self.filepath_kassen_system)
        kassen_system_data_dict, mr_s, mc_s = self.load_kassen_system_expiry_excel_file(workbook)
        json_data = self.load_json_data_website_products_mhd("products_mhd.json")
        weight_updated_products, match_of_stock_cells_count, num_no_match_found = self.match_products_and_update_mhd(json_data, kassen_system_data_dict)
        MAX_API_BATCH_SIZE = 100
        def chunks(l, n):
            """Yield successive n-sized chunks from l."""
            for i in range(0, len(l), n):
                yield l[i:i + n]
        try:
            for batch in chunks(self.products_list, MAX_API_BATCH_SIZE):
                print(len(batch))
                print(self.wcapi.put("products/batch", {"update": batch}).json())
            subject = '[Staging] lotus-grocery.eu - Stock Updated successfully on ' + datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S")
            message = "This is an automated mail, receives this mail once the stock updates successfully. " \
                      "The statistics are as follows:\n\n\n" \
                      "Total no of Rows/Products in Source file from Shop File:{}\n " \
                      "Total no of Rows/Products in Destination file in Website:{}\n" \
                      "Number of Products Matched:{}\n" \
                      "Number of Products are no matched:{}\n" \
                      "Number of Products Weights updated:{}\n" \
                      "Time elasped:{} seconds".format(mr_s, len(json_data), match_of_stock_cells_count,
                                                       num_no_match_found, weight_updated_products, timeit.default_timer() - start)
            content = [message, "./no_match_products.txt", "./products_without_weight.txt"]
            self.send_email(subject, content)
        except:
            subject = '[Staging] lotus-grocery.eu - Stock Updated not successfully on ' + datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S")
            message = "Stock update not successful. Please re-run the scripts again"
            content = [message]
            self.send_email(subject, content)
        print("Total no of Rows/Products in Source file from Shop File:{}".format(mr_s))
        print("Total no of Rows/Products in Destination file in Website:{}".format(len(json_data)))
        print("Number of Products Matched:{}".format(match_of_stock_cells_count))
        print("Number of Products are no matched:{}".format(num_no_match_found))
        print("Number of Products Weights updated:{}".format(weight_updated_products))
        stop = timeit.default_timer()

        print('Time: ', stop - start)


if __name__ == "__main__":
    filepath_kassen_system = r"C:\Users\e04ux6p\Downloads\products_expiry_list.xlsx"
    staging_products_update = StagingUpdateProductsMHD(filepath_kassen_system)
    staging_products_update.process()
