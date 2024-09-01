from configparser import ConfigParser

CONFIG_FOLDER = 'config'

def parse_dict(section):
	dictionary = {}
	for key in section:
		dictionary[key] = section[key]
	return dictionary

config = ConfigParser()
config.read(f'{CONFIG_FOLDER}/config.ini')

vat = ConfigParser()
vat.read(f'{CONFIG_FOLDER}/vat.ini')

DATE_COL = config['REQUIRED_COLUMNS']['date']
COUNTRY_COL = config['REQUIRED_COLUMNS']['country']
TOTAL_COL = config['REQUIRED_COLUMNS']['total']

EU_VAT = parse_dict(vat['EU_VAT'])

VARIABLES = parse_dict(config['VARIABLES'])
CONSTANTS = parse_dict(config['CONSTANTS'])

TEMPLATE_PATH = config['FILENAMES']['template']
SALES_OUTPUT = config['FILENAMES']['sales output']
TEMPLATE_OUTPUT = config['FILENAMES']['template output']
