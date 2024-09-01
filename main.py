from extract_data import DataExtractor
from modify_data import SplitSalesByCountry
from save_data import LoadWorkbook, SaveData

try:
	from config_reader import EU_VAT, DATE_COL, COUNTRY_COL, TOTAL_COL
except Exception as e:
	print("[‼] Failed to import config. See the error below:")
	print(type(e), e)
	input("\nPress Enter to exit...")
	exit()


def print_title():
	print(r"""
     ╔═════════════════════════════════════════════════════════════════════════════════════════╗
    ╔╝   _____  __       _____             _       __                           _   _          ╚══╗
   ╔╝   | ___ \/  |     |  ___|           | |     / _|                         | | | |            ╚╗
   ║    | |_/ /`| |     | |____  _____ ___| |    | |_ ___  _ __ _ __ ___   __ _| |_| |_ ___ _ __   ║
   ║    | ___ \ | |     |  __\ \/ / __/ _ \ |    |  _/ _ \| '__| '_ ` _ \ / _` | __| __/ _ \ '__|  ║
   ║    | |_/ /_| |_    | |___>  < (_|  __/ |    | || (_) | |  | | | | | | (_| | |_| ||  __/ |     ║
   ║    \____/ \___/    \____/_/\_\___\___|_|    |_| \___/|_|  |_| |_| |_|\__,_|\__|\__\___|_|     ║
   ║                                                                                              ╔╝
   ╚══════════════════════════════════════════════════════════════════════════════════════════════╝
""")


def runner():
	print_title()
	print("[?] Enter the path (filename if the file is in the same folder) to the SALES REPORT FILE or drag it into this window:")
	input_filename = input("» ")
	try:
		wb = LoadWorkbook(input_filename, True)
		ext = DataExtractor(wb.sheet, DATE_COL, COUNTRY_COL, TOTAL_COL)
		data, status = ext.run()
		if status:
			sales = SplitSalesByCountry(data, EU_VAT)
			SaveData(data, sales.eu, sales.not_eu)
	except Exception as e:
		print("[‼] Exception occurred while running the app. See the error below:")
		print(type(e), e)
	finally:
		input("\nPress Enter to exit...")


if __name__ == '__main__':
	runner()
