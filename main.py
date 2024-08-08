from openpyxl import load_workbook

DATE_COL = 0
COUNTRY_COL = 14
TOTAL_COL = 23

workbook = load_workbook('../EtsySoldOrders2024-7.xlsx', read_only=True)

sheet = workbook.active
print(f"Data from sheet '{sheet.title}'")

headers = [sheet['1'][col].value for col in (DATE_COL, COUNTRY_COL, TOTAL_COL)]
print(f"Columns to parse: {headers[0]}, {headers[1]}, {headers[2]}")

print("\nRunning:")
for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
	date = row[DATE_COL]
	country = row[COUNTRY_COL]
	total = row[TOTAL_COL]

	if not (date or country or total):
		print(f"No data in row {i}. Skipped")
		continue

	# [Sale Date]
	if date:
		try:
			(months, day, year_tens) = date.split('/')
			new_date = f'20{year_tens}-{months}-{day}'
		except Exception as e:
			print(f"Incorrect '{headers[0]}' in row {i}: '{date}'")
			new_date = date
	else:
		print(f"No '{headers[0]}' in row {i}: ({date}, {country}, {total})")

	# [Ship Country]
	if not country:
		print(f"No '{headers[1]}' in row {i}: ({date}, {country}, {total})")

	# [Order Total]
	if total:
		if isinstance(total, str):
			try:
				old_total, total = total, float(total.replace(',', ''))
				print(f"Incorrect '{headers[2]}' in row {i}: '{old_total}'. Changed to '{total}'")
			except Exception as e:
				print(f"Incorrect '{headers[2]}' in row {i}: '{total}'. {type(e)}: {e}")
	else:
		print(f"No '{headers[2]}' in row {i}: ({date}, {country}, {total})")

	print(f"{i}) {new_date if 'new_date' in locals() else date} | {country} | {total}")
