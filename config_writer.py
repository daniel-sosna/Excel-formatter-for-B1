from configparser import ConfigParser
import os

CONFIG_FOLDER = 'config'

if not os.path.exists(CONFIG_FOLDER):
	os.makedirs(CONFIG_FOLDER)

config = ConfigParser()
vat = ConfigParser()

config['REQUIRED_COLUMNS'] = {
	'DATE': 'A',
	'COUNTRY': 'O',
	'TOTAL': 'X',
}
config['FILENAMES'] = {
	'TEMPLATE': 'config/template.xlsx',
	'SALES OUTPUT': 'pardavimai',
	'TEMPLATE OUTPUT': 'b1_import',
}
config['VARIABLES'] = {
	'date': 'A',
	'number': 'D',
	'country': 'L',
	'price': 'W',
}
config['CONSTANTS'] = {
	'C': 'SF',
	'E': 'Pardavimai',
	'F': 'EUR',
	'G': '',
	'N': '',
	'R': 'Pagrindinis',
	'S': 'Pardavimas, linas',
	'V': 1.000,
	'Y': 0.00,
	'Z': 'PVM12',
	'AA': 50001,
}
vat['EU_VAT'] = {
	'Ireland': 23,
	'Austria': 20,
	'Belgium': 21,
	'Bulgaria': 20,
	'Czech Republic': 21,
	'Denmark': 25,
	'Estonia': 20,
	'Greece': 24,
	'Spain': 21,
	'Italy': 22,
	'Cyprus': 19,
	'Croatia': 25,
	'Latvia': 21,
	'Poland': 23,
	'Luxembourg': 17,
	'Malta': 18,
	'The Netherlands': 21,
	'Portugal': 23,
	'France': 20,
	'Romania': 19,
	'Slovakia': 20,
	'Slovenia': 22,
	'Finland': 24,
	'Sweden': 25,
	'Hungary': 27,
	'Germany': 19,
}

with open(f'{CONFIG_FOLDER}/config.ini', 'w') as file:
	config.write(file)

with open(f'{CONFIG_FOLDER}/vat.ini', 'w') as file:
	vat.write(file)
