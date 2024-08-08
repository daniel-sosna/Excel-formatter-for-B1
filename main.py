from openpyxl import load_workbook

def col_to_ind(column:str, start:int=0) -> int:
	''' Converts a column name (e.g. 'A', 'AF', 'CK') to an index '''
	index = start - 1
	for i, letter in enumerate(column[::-1]): # Run through reversed column name
		index += (ord(letter) - ord('A') + 1) * 26**i
	# 'ABC'  ->  'CBA'  ->  'A'*(26^2) + 'B'*(26^1) + 'C'*(26^0)  ->  1*676 + 2*26 + 3*1  ->  731
	return index

DATE_COL = 'A'
COUNTRY_COL = 'O'
TOTAL_COL = 'X'

workbook = load_workbook('../EtsySoldOrders2024-7.xlsx', read_only=True)

sheet = workbook.active
print(f"Data from sheet '{sheet.title}'")

headers = [sheet[col+'1'].value for col in (DATE_COL, COUNTRY_COL, TOTAL_COL)]
print(f"Columns to parse: {headers[0]}, {headers[1]}, {headers[2]}")

data = []
error_count = 0
print("\nRunning:")
for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
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
		data.append((new_date if 'new_date' in locals() else date, country, total))
	else:
		error_count += 1

print("Results:")
print(f"{len(data) + error_count} rows have been found.")
if error_count:
	print(f" ├─ {len(data)} rows have been parsed.")
	print(f" └─ {error_count} rows are invalid! Please fix all ❌ marks first.")
else:
	print(f" └─ {len(data)} rows have been parsed.")
	print("No critical errors found. Going further...")

data.sort(key=lambda x: x[0])

print(f"\nData:")
for (date, country, total) in data:
	print(f"{date} | {country} | {total}")