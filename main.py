from utils import col_to_ind
from export_data import LoadWorkbook, SaveData
from vat import EU_VAT
from config import DATE_COL, COUNTRY_COL, TOTAL_COL


class DataExtractor():
	def __init__(self, sheet):
		self.sheet = sheet
		print(f"Getting data from the sheet '{self.sheet.title}'")
		self.headers = [self.sheet[col+'1'].value for col in (DATE_COL, COUNTRY_COL, TOTAL_COL)]
		print(f"Columns to parse: {self.headers[0]}, {self.headers[1]}, {self.headers[2]}\n")

	def run(self, start=2, stop=None) -> tuple[list, bool]:
		data = []
		skipped_count = 0
		error_count = 0

		print("# Extracting sales:")
		for i, row in enumerate(self.sheet.iter_rows(min_row=start, values_only=True), start=start):
			if stop and i > stop:
				i -= 1
				break

			row_data = self.get_row_data(row)
			if row_data:
				checked_data, is_valid = self.check_data(i, *row_data)
				if is_valid:
					data.append(checked_data)
				else:
					error_count += 1
			else:
				skipped_count += 1

		self.print_results(i - start + 1, len(data), skipped_count, error_count)
		return sorted(data, key=lambda x: x[0]), True if not error_count else False

	def get_row_data(self, row) -> tuple | None:
		# Get needed columns data
		date = row[col_to_ind(DATE_COL)]
		country = row[col_to_ind(COUNTRY_COL)]
		total = row[col_to_ind(TOTAL_COL)]

		# Return None if row is blank
		if not (date or country or total):
			print(f"🎨 No data in row {i}. Skipped")
			return None

		return (date, country, total)

	def check_data(self, i, date, country, total) -> tuple[tuple, bool]:
		is_row_valid = True

		# [Sale Date]
		new_date = date
		if date:
			try:
				(months, day, year_tens) = date.split('/')
				new_date = f'20{year_tens}-{months}-{day}'
			except Exception as e:
				print(f"❌ [{DATE_COL}{i}] Incorrect '{self.headers[0]}' in row {i}: '{date}'")
				is_row_valid = False
		else:
			print(f"❌ [{DATE_COL}{i}] No '{self.headers[0]}' in row {i}: ({date}, {country}, {total})")
			is_row_valid = False

		# [Ship Country]
		if not country:
			print(f"❌ [{COUNTRY_COL}{i}] No '{self.headers[1]}' in row {i}: ({date}, {country}, {total})")
			is_row_valid = False

		# [Order Total]
		if total:
			if isinstance(total, str):
				try:
					old_total, total = total, float(total.replace(',', ''))
					print(f"🎨 [{TOTAL_COL}{i}] Incorrect '{self.headers[2]}' in row {i}: '{old_total}'. Changed to '{total}'")
				except Exception as e:
					print(f"❌ [{TOTAL_COL}{i}] Incorrect '{self.headers[2]}' in row {i}: '{total}'. {type(e)}: {e}")
					is_row_valid = False
		else:
			print(f"❌ [{TOTAL_COL}{i}] No '{self.headers[2]}' in row {i}: ({date}, {country}, {total})")
			is_row_valid = False

		return (new_date, country, total), is_row_valid

	def print_results(self, n_rows_listened, n_rows_valid, n_rows_skipped, n_errors):
		print("# Extraction from Excel results:")
		print(f"{n_rows_listened} rows have been listened.")
		print(f" ├─ {n_rows_skipped} rows without data skipped.")
		print(f" └─ {n_rows_valid + n_errors} rows have been parsed.")
		if n_errors:
			print(f"     ├─ {n_rows_valid} rows have been saved.")
			print(f"     └─ {n_errors} rows are invalid! Please fix all ❌ marks first.\n")
		else:
			print("✔ All rows with data have been saved.")
			print("No critical errors found. Going further...\n")


class SplitSalesByCountry():
	eu = dict()
	not_eu = list()
	eu_countries = dict()
	not_eu_countries = dict()

	def __init__(self, sales):
		self.all = sales
		self.split_sales()
		self.count_vat_for_eu()
		self.print_results()

	def split_sales(self):
		for row in self.all:
			country = row[1]
			if country in EU_VAT.keys():
				if country not in self.eu:
					self.eu[country] = 0
				self.eu[country] += row[2]
				if country not in self.eu_countries:
					self.eu_countries[country] = 0
				self.eu_countries[country] += 1
			else:
				self.not_eu.append(row)
				if country not in self.not_eu_countries:
					self.not_eu_countries[country] = 0
				self.not_eu_countries[country] += 1

	def count_vat_for_eu(self):
		for country in self.eu.keys():
			total = self.eu[country]
			without_vat = total * 100 / (100 + EU_VAT[country])
			vat = total * EU_VAT[country] / (100 + EU_VAT[country])
			self.eu[country] = (without_vat, vat, total)

	def print_results(self):
		print("# Sales summary:")
		self.print_countries("EU", self.eu_countries)
		self.print_countries("not EU", self.not_eu_countries)
		print()

	def print_countries(self, title, countries):
		total = 0
		for n in countries.values():
			total += n
		print(f"{total} sales in {title} countries have been found. {len(countries)} countries in total:")
		for country, n in sorted(countries.items(), key=lambda item: item[1], reverse=True):
			print(f" ● {country} - {n}")


def main():
	input_filename = input("↪ Enter the path (filename if the file is in the same folder) to the SALES REPORT FILE or drag it into this window: ")
	wb = LoadWorkbook(input_filename, True)
	if not wb.sheet:
		return
	ext = DataExtractor(wb.sheet)
	data, status = ext.run()
	if status:
		sales = SplitSalesByCountry(data)
		SaveData(data, sales.eu, sales.not_eu)


if __name__ == '__main__':
	main()
