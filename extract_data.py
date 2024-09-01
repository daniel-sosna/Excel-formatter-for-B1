from utils import col_to_ind


class DataExtractor():
    def __init__(self, sheet, DATE_COL, COUNTRY_COL, TOTAL_COL):
        self.sheet = sheet
        self.DATE_COL = DATE_COL
        self.COUNTRY_COL = COUNTRY_COL
        self.TOTAL_COL = TOTAL_COL

        print(f"Getting data from the sheet '{self.sheet.title}'")
        self.headers = [self.sheet[col+'1'].value for col in (self.DATE_COL, self.COUNTRY_COL, self.TOTAL_COL)]
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

            row_data = self.get_row_data(i, row)
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

    def get_row_data(self, i, row) -> tuple | None:
        # Get needed columns data
        date = row[col_to_ind(self.DATE_COL)]
        country = row[col_to_ind(self.COUNTRY_COL)]
        total = row[col_to_ind(self.TOTAL_COL)]

        # Return None if row is blank
        if not (date or country or total):
            print(f" ~ No data in row {i}. Skipped")
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
                print(f" X [{self.DATE_COL}{i}] Incorrect '{self.headers[0]}' in row {i}: '{date}'")
                is_row_valid = False
        else:
            print(f" X [{self.DATE_COL}{i}] No '{self.headers[0]}' in row {i}: ({date}, {country}, {total})")
            is_row_valid = False

        # [Ship Country]
        if not country:
            print(f" X [{self.COUNTRY_COL}{i}] No '{self.headers[1]}' in row {i}: ({date}, {country}, {total})")
            is_row_valid = False

        # [Order Total]
        if total:
            if isinstance(total, str):
                try:
                    old_total, total = total, float(total.replace(',', ''))
                    print(f" ~ [{self.TOTAL_COL}{i}] Incorrect '{self.headers[2]}' in row {i}: '{old_total}'. Changed to '{total}'")
                except Exception as e:
                    print(f" X [{self.TOTAL_COL}{i}] Incorrect '{self.headers[2]}' in row {i}: '{total}'. {type(e)}: {e}")
                    is_row_valid = False
        else:
            print(f" X [{self.TOTAL_COL}{i}] No '{self.headers[2]}' in row {i}: ({date}, {country}, {total})")
            is_row_valid = False

        return (new_date, country, total), is_row_valid

    def print_results(self, n_rows_listened, n_rows_valid, n_rows_skipped, n_errors):
        print("# Extraction from Excel results:")
        print(f"{n_rows_listened} rows have been listened.")
        print(f" ├─ {n_rows_skipped} rows without data skipped.")
        print(f" └─ {n_rows_valid + n_errors} rows have been parsed.")
        if n_errors:
            print(f"     ├─ {n_rows_valid} rows have been saved.")
            print(f"     └─ {n_errors} rows are invalid!")
            print("[!] Please fix all X marks first.\n")
        else:
            print("[+] All rows with data have been saved.")
            print("No critical errors found. Going further...\n")
