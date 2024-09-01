class SplitSalesByCountry():
    eu = dict()
    not_eu = list()
    eu_countries = dict()
    not_eu_countries = dict()

    def __init__(self, sales, EU_VAT):
        self.all = sales
        self.EU_VAT = EU_VAT
        self.split_sales()
        self.count_vat_for_eu()
        self.print_results()

    def split_sales(self):
        for row in self.all:
            country = row[1]
            if country in self.EU_VAT.keys():
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
            VAT_value = float(self.EU_VAT[country])
            without_vat = total * 100 / (100 + VAT_value)
            vat = total * VAT_value / (100 + VAT_value)
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
