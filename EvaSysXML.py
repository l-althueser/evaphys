#!/usr/bin/env python

import os, glob, sys, getopt
from glob import iglob
from xml.etree import ElementTree
from xml.dom import minidom
import pandas as pd
import unicodedata
import re

def usage():
	print('EvaSysXML.py --convert-to-html -split-by-<ID or ORG> -k <single --ID or --ORG> -i <HISLSF_XML> -o <filename>')
	sys.exit(2)

def iter_docs(Evasys):
    Evasys_attr = Evasys.attrib
    for Lecture in Evasys.iter('Lecture'):
        Lecture_dict = Evasys_attr.copy()
        for node in Lecture.getiterator():
            if node.tag in ["name", "orgroot", "short", "type", "period"]:
                Lecture_dict[node.tag] = node.text
        
        dozs = []
        for doz in Lecture.find('dozs'):
            doz_key = doz.find('EvaSysRef').attrib['key']
            
            title = ''
            firstname = ''
            lastname = ''
            email = ''
            for node in [x for x in Evasys.findall('Person') if x.attrib['key'] == doz_key][0].getiterator():
                if node.tag == "title":
                    title = str(node.text)
                if node.tag == "firstname":
                    firstname = str(node.text)
                if node.tag == "lastname":
                    lastname = str(node.text)
                if node.tag == "email":
                    email = str(node.text)
            dozs.append(" ".join([title,firstname,lastname,"("+email+")"]))
        Lecture_dict["dozs"] = ", ".join(dozs)
            
        Lecture_dict.update(Lecture.attrib)
        if Lecture_dict["name"]:
            yield Lecture_dict

def prettify(elem):
	"""Return a pretty-printed XML string for the Element.
	"""
	rough_string = ElementTree.tostring(elem, 'utf-8')
	reparsed = minidom.parseString(rough_string)
	#return reparsed.toprettyxml(indent="  ", encoding="UTF-8")
	return '\n'.join([line for line in reparsed.toprettyxml(indent=' '*2, encoding="UTF-8").decode('utf8').split('\n') if line.strip()])

def slugify(value):
    """
    Converts to lowercase, removes non-word characters (alphanumerics and
    underscores) and converts spaces to hyphens. Also strips leading and
    trailing whitespace.
    """
    value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
    value = re.sub('[^\w\s-]', '', value).strip().lower()
    return re.sub('[-\s]+', '-', value)
	
def process_XML(convert_format, convert, split_format, split_keys, split, input_file, output_file):
	"""Process XML file ...

	Keyword arguments:
	convert_format: ...
	convert: True or False
	split_keys: String key to filter the XML data (default: '01' to '15')
	split: True or False
	input_file: HISLSF export for EvaSys
	output_file: Basename of the output file(s)
	"""

	print("XML file: " + input_file)
		
	if not output_file:
		output_file = input_file
	print("Output filename: " + os.path.splitext(output_file)[0] + "...")
	
	if split:	
		if not split_keys and split_format == "ID": 
			split_keys = [str(i).zfill(2) for i in range(1,16)]
		elif not split_keys and split_format == "ORG": 
			ORGs = []
			EvaSys = ElementTree.parse(input_file).getroot()
			for Lecture in EvaSys.findall('Lecture'):
				for node in Lecture.getiterator():
					if node.tag == 'orgroot' and node.text:
						ORGs.append(node.text)
			split_keys = list(set(ORGs))
		else:
			split_keys = [split_keys]
		print("Split format: " + split_format)
		print("Split by key(s): " + ' | '.join(split_keys))
		
		for key in split_keys:
			if split_format == "ID":
				file_output = os.path.splitext(output_file)[0] + '-FB'+key
			if split_format == "ORG":
				file_output = os.path.splitext(output_file)[0] + '-'+slugify(key)
				
			EvaSys = ElementTree.parse(input_file).getroot()
			for Lecture in EvaSys.findall('Lecture'):
				for node in Lecture.getiterator():
					if split_format == "ID" and node.tag == 'short':
						if not node.text or not node.text.startswith(key):
							EvaSys.remove(Lecture)
					if split_format == "ORG" and node.tag == 'orgroot':
						if not node.text or not node.text == key:
							EvaSys.remove(Lecture)
			
			dozs_keys = []
			for Lecture in EvaSys.findall('Lecture'):
				for doz in Lecture.find('dozs'):
						dozs_keys.append(doz.find('EvaSysRef').attrib['key'])

			Persons = EvaSys.findall('Person')
			for Person in Persons:
				if not Person.attrib['key'] in dozs_keys:
					EvaSys.remove(Person)

			with open(file_output+'.xml', mode='w') as f:
				f.write(prettify(EvaSys))
				
			if convert:
				EvaSys_df = pd.DataFrame(list(iter_docs(EvaSys)))
				try:
					EvaSys_df = EvaSys_df[['key', 'period', 'orgroot', 'type', 'name', 'dozs', 'short']]
				except KeyError:
					pass
				EvaSys_df.to_html(file_output+'.html', justify="center")
				
	else:
		if convert:
			EvaSys = ElementTree.parse(input_file).getroot()
			EvaSys_df = pd.DataFrame(list(iter_docs(EvaSys)))
			try:
				EvaSys_df = EvaSys_df[['key', 'period', 'orgroot', 'type', 'name', 'dozs', 'short']]
			except KeyError:
				pass
			EvaSys_df.to_html(os.path.splitext(output_file)[0] + '.html', justify="center")
		
	print(".. EvaSysXML finished.")

if __name__ == "__main__":

	argv = sys.argv[1:]
	convert_format = ''
	convert = False
	split_format = ''
	split_keys = ''
	split = False
	input_file = ''
	output_file = ''

	try:
		opts, args = getopt.getopt(argv,"k:i:o:",["convert-to-html","ID=","ORG=","input=","output=","split-by-ID","split-by-ORG"])
	except getopt.GetoptError:
		usage()
		
	for opt, arg in opts:
		if opt in ("--convert-to-html"):
			convert_format = "html"
			convert = True
		elif opt in ("--split-by-ID"):
			split_format = 'ID'
			split = True
		elif opt in ("--split-by-ORG"):
			split_format = 'ORG'
			split = True
		elif opt in ("-k", "--ID", "--ORG"):
			split_keys = arg
			split = True
		elif opt in ("-i", "--input"):
			input_file = arg
		elif opt in ("-o", "--output"):
			output_file = arg
			
	if not input_file:
		usage()
			
	if not convert and not split:
		usage()
		
	process_XML(convert_format, convert, split_format, split_keys, split, input_file, output_file)
