from openpyxl import load_workbook

DATE_COL = 'A'
COUNTRY_COL = 'O'
TOTAL_COL = 'X'

def col_to_ind(column:str, start:int=0) -> int:
	''' Converts a column name (e.g. 'A', 'AF', 'CK') to an index '''
	index = start - 1
	for i, letter in enumerate(column[::-1]): # Run through reversed column name
		index += (ord(letter) - ord('A') + 1) * 26**i
	# 'ABC'  ->  'CBA'  ->  'A'*(26^2) + 'B'*(26^1) + 'C'*(26^0)  ->  1*676 + 2*26 + 3*1  ->  731
	return index


class LoadWorkbook():
	def __init__(self, filename):
		workbook = load_workbook(filename, read_only=True)
		self.open_worksheet(workbook)

	def open_worksheet(self, workbook):
		self.sheet = workbook.active
		print(f"Data from sheet '{self.sheet.title}'")


class DataExtractor():
	data = []
	error_count = 0

	def __init__(self, sheet):
		self.sheet = sheet
		headers = [sheet[col+'1'].value for col in (DATE_COL, COUNTRY_COL, TOTAL_COL)]
		print(f"Columns to parse: {headers[0]}, {headers[1]}, {headers[2]}")
		self.extract(headers)
		self.print_results()

	def extract(self, headers):
		print("\nRunning:")
		for i, row in enumerate(self.sheet.iter_rows(min_row=2, values_only=True), start=2):
			# Get needed columns data
			date = row[col_to_ind(DATE_COL)]
			country = row[col_to_ind(COUNTRY_COL)]
			total = row[col_to_ind(TOTAL_COL)]

			# If row is blank
			if not (date or country or total):
				print(f"✔ No data in row {i}. Skipped")
				continue

			is_row_valid = True

			# [Sale Date]
			if date:
				try:
					(months, day, year_tens) = date.split('/')
					new_date = f'20{year_tens}-{months}-{day}'
				except Exception as e:
					print(f"❌ [{DATE_COL}{i}] Incorrect '{headers[0]}' in row {i}: '{date}'")
					new_date = date
					is_row_valid = False
			else:
				print(f"❌ [{DATE_COL}{i}] No '{headers[0]}' in row {i}: ({date}, {country}, {total})")
				is_row_valid = False

			# [Ship Country]
			if not country:
				print(f"❌ [{COUNTRY_COL}{i}] No '{headers[1]}' in row {i}: ({date}, {country}, {total})")
				is_row_valid = False

			# [Order Total]
			if total:
				if isinstance(total, str):
					try:
						old_total, total = total, float(total.replace(',', ''))
						print(f"✔ [{TOTAL_COL}{i}] Incorrect '{headers[2]}' in row {i}: '{old_total}'. Changed to '{total}'")
					except Exception as e:
						print(f"❌ [{TOTAL_COL}{i}] Incorrect '{headers[2]}' in row {i}: '{total}'. {type(e)}: {e}")
						is_row_valid = False
			else:
				print(f"❌ [{TOTAL_COL}{i}] No '{headers[2]}' in row {i}: ({date}, {country}, {total})")
				is_row_valid = False

			if is_row_valid:
				self.data.append((new_date if 'new_date' in locals() else date, country, total))
			else:
				self.error_count += 1

		self.data.sort(key=lambda x: x[0])

	def print_results(self):
		print("Results:")
		print(f"{len(self.data) + self.error_count} rows have been found.")
		if self.error_count:
			print(f" ├─ {len(self.data)} rows have been parsed.")
			print(f" └─ {self.error_count} rows are invalid! Please fix all ❌ marks first.")
		else:
			print(f" └─ {len(self.data)} rows have been parsed.")
			print("No critical errors found. Going further...")


def main():
	wb = LoadWorkbook('../EtsySoldOrders2024-7.xlsx')
	ext = DataExtractor(wb.sheet)
	print(f"\nData:")
	for (date, country, total) in ext.data:
		print(f"{date} | {country} | {total}")

if __name__ == '__main__':
	main()
