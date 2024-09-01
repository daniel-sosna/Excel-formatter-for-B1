from configparser import ConfigParser

CONFIG_FOLDER = 'config'

def parse_dict(section):
	dictionary = {}
	for key in section:
		dictionary[key] = section[key]
	return dictionary

config = ConfigParser()
config.optionxform = str
config.read(f'{CONFIG_FOLDER}/config.ini')

vat = ConfigParser()
vat.optionxform = str
vat.read(f'{CONFIG_FOLDER}/vat.ini')

DATE_COL = config['REQUIRED_COLUMNS']['DATE']
COUNTRY_COL = config['REQUIRED_COLUMNS']['COUNTRY']
TOTAL_COL = config['REQUIRED_COLUMNS']['TOTAL']

EU_VAT = parse_dict(vat['EU_VAT'])

VARIABLES = parse_dict(config['VARIABLES'])
CONSTANTS = parse_dict(config['CONSTANTS'])

TEMPLATE_PATH = config['FILENAMES']['TEMPLATE']
SALES_OUTPUT = config['FILENAMES']['SALES OUTPUT']
TEMPLATE_OUTPUT = config['FILENAMES']['TEMPLATE OUTPUT']
