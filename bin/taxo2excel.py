#!/usr/bin/python
#
# Converts a taxonomy in YAML to a Excel file: first sheet lists the vocabularies, 
# following sheets list the terms for each of the vocabularies.
#
# Requires: 
# * XlsxWriter https://xlsxwriter.readthedocs.io/
# * PyYaml http://pyyaml.org/
#

import argparse
import yaml
import xlsxwriter

parser = argparse.ArgumentParser()
parser.add_argument('--taxo', required=True, help='Taxonomy YML file path.')
parser.add_argument('--out', required=True, help='Taxonmy Excel file path.')
args = parser.parse_args()

class Taxonomy(object):
	def __init__(self, values):
		self.name = values["name"]
		self.author = values["author"]
		self.license = values["license"]
		self.title = values["title"]
		self.description = values["description"]
		self.vocabularies = values["vocabularies"]

def taxonomy_constructor(loader, node):
	return Taxonomy(loader.construct_mapping(node, deep=True))

yaml.add_constructor('tag:yaml.org,2002:org.obiba.opal.core.domain.taxonomy.Taxonomy', taxonomy_constructor)

def write_taxonomy_object(taxonomy_object, ws, row):
	ws.write(row, 0, taxonomy_object['name'])
	ws.write(row, 1, taxonomy_object['title']['en'])
	ws.write(row, 2, taxonomy_object['title']['fr'])
	ws.write(row, 3, taxonomy_object['description']['en'])
	ws.write(row, 4, taxonomy_object['description']['fr'])

def write_taxonomy(taxonomy, wb):
	txws = wb.add_worksheet('_Vocabularies_')
	txws.write('A1', 'name')
	txws.write('B1', 'title:en')
	txws.write('C1', 'title:fr')
	txws.write('D1', 'description:en')
	txws.write('E1', 'description:fr')
	vocrow = 1
	for vocabulary in taxonomy.vocabularies:
		write_taxonomy_object(vocabulary, txws, vocrow)
		vocrow = vocrow + 1
		ws = wb.add_worksheet(vocabulary["name"][0:30])
		ws.write('A1', 'name')
		ws.write('B1', 'title:en')
		ws.write('C1', 'title:fr')
		ws.write('D1', 'description:en')
		ws.write('E1', 'description:fr')
		termrow = 1
		for term in vocabulary['terms']:
			write_taxonomy_object(term, ws, termrow)
			termrow = termrow + 1

workbook = xlsxwriter.Workbook(args.out)
with open(args.taxo, 'r') as stream:
    try:
    	taxo = yaml.load(stream)
        write_taxonomy(taxo, workbook)
        workbook.close()
    except yaml.YAMLError as exc:
        print(exc)
