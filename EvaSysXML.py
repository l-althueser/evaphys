#!/usr/bin/env python
# coding: utf8

import os, glob, sys, getopt
from glob import iglob
import requests
from urllib.parse import urlparse, urlencode
from xml.etree import ElementTree
from xml.dom import minidom
import pandas as pd
import unicodedata
import re
import datetime
import getpass
import codecs

def usage():
	print('EvaSysXML.py --download-semester <20181> --convert-to <html, csv or excel> --split-by <ID or ORG> --ID <single ID e.g. 11> --ORG <single ORG> --include-type <Vorlesung,V/Ü> --exclude-type <Forschungsseminar,Kolloquium> -o <filename>')
	print('EvaSysXML.py --input <HISLSF XML export> --convert-to <html, csv or excel> --split-by <ID or ORG> --ID <single ID e.g. 11> --ORG <single ORG> --include-type <Vorlesung,V/Ü> --exclude-type <Forschungsseminar,Kolloquium> -o <filename>')
	sys.exit(2)

def iter_docs(Evasys):
	Evasys_attr = Evasys.attrib
	for Lecture in Evasys.iter('Lecture'):
		Lecture_dict = Evasys_attr.copy()

		p_o_study = []
		for node in Lecture.getiterator():
			if node.tag in ["name", "orgroot", "type", "period", "turnout"]:
				Lecture_dict[node.tag] = node.text
			elif node.tag == "short":
				Lecture_dict[node.tag] = node.text
				Lecture_dict["LSF_ID"] = node.text.split()[0]
			elif node.tag == "p_o_study":
				p_o_study.append(str(node.text))
		Lecture_dict["p_o_study"] = ", ".join(p_o_study)

		dozs = []
		dozs_mail = []
		for doz in Lecture.find('dozs'):
			doz_key = doz.find('EvaSysRef').attrib['key']

			title = ''
			firstname = ''
			lastname = ''
			email = ''
			for node in [x for x in Evasys.findall('Person') if x.attrib['key'] == doz_key][0].getiterator():
				xstr = lambda s: '' if s is None else str(s)
				if node.tag == "title":
					title = xstr(node.text)
				elif node.tag == "firstname":
					firstname = xstr(node.text)
				elif node.tag == "lastname":
					lastname = xstr(node.text)
				elif node.tag == "email":
					email = xstr(node.text)
			# write dozs formatted as: "John Smith" <johnsemail@hisserver.com>
			dozs.append(" ".join(filter(None, [firstname,lastname])))
			dozs_mail.append('"'+" ".join(filter(None, [title,firstname,lastname]))+'"'+" <"+email+">")
		Lecture_dict["dozs"] = ", ".join(dozs)
		Lecture_dict["dozs_mail"] = ", ".join(dozs_mail)

		Lecture_dict.update(Lecture.attrib)
		if Lecture_dict["name"]:
			yield Lecture_dict

def slugify(value):
    """
    Converts to lowercase, removes non-word characters (alphanumerics and
    underscores) and converts spaces to hyphens. Also strips leading and
    trailing whitespace.
    """
    value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
    value = re.sub('[^\w\s-]', '', value).strip().lower()
    return re.sub('[-\s]+', '-', value)

def process_XML(convert_format, split_format, split_keys, include_type, exclude_type, remove_duplicates, input_file, output_file):
	"""Process XML file ...

	Keyword arguments:
	convert_format: ...
	split_keys: String key to filter the XML data (default: '01' to '15')
	input_file: HISLSF export for EvaSys
	output_file: Basename of the output file(s)
	"""

	if not os.path.splitext(input_file)[-1] == ".xml":
		usage()
	print("==== XML file statistics ====")
	print("Filename: " + input_file)

	n_lectures = 0
	lecture_types = []
	lecture_orgroot = []
	n_persons = 0
	EvaSys = ElementTree.parse(input_file).getroot()
	for Lecture in EvaSys.findall('Lecture'):
		n_lectures += 1
		for node in Lecture.getiterator():
			if node.tag == 'type' and node.text:
				lecture_types.append(node.text)
			if node.tag == 'orgroot' and node.text:
				lecture_orgroot.append(node.text)
	for Person in EvaSys.findall('Person'):
		n_persons += 1

	print("Number of Lectures: " + str(n_lectures))
	print("Number of Persons:  " + str(n_persons))

	print("Lecture types:   " + " | ".join(list(set(lecture_types))))
	print("Lecture orgroot: " + " | ".join(list(set(lecture_orgroot))))

	print("")
	print("==== Process XML file ====")
	if not output_file:
		output_file = input_file
	print("Write result to: " + os.path.splitext(output_file)[0])

	if split_format:
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
		print("Include type(s): " + ' | '.join(include_type))
		print("Exclude type(s): " + ' | '.join(exclude_type))

		for key in split_keys:
			if split_format == "ID":
				output_file_split = os.path.splitext(output_file)[0] + '-FB'+key
			if split_format == "ORG":
				output_file_split = os.path.splitext(output_file)[0] + '-'+slugify(key)

			EvaSys = ElementTree.parse(input_file).getroot()
			remove_lectures = []
			for Lecture in EvaSys.findall('Lecture'):
				for node in Lecture.getiterator():
					if split_format == "ID" and node.tag == 'short':
						if not node.text or not node.text.startswith(key):
							remove_lectures.append(Lecture)
					if split_format == "ORG" and node.tag == 'orgroot':
						if not node.text or not node.text == key:
							remove_lectures.append(Lecture)
					if include_type and node.tag == 'type' and not node.text in include_type:
						remove_lectures.append(Lecture)
					if exclude_type and node.tag == 'type' and node.text in exclude_type:
						remove_lectures.append(Lecture)
			for Lecture in list(set(remove_lectures)):
				EvaSys.remove(Lecture)

			if remove_duplicates:
				remove_lectures = []
				seen_lectures = []

				for Lecture in EvaSys.findall('Lecture'):
					if Lecture.attrib['key'] in seen_lectures:
						remove_lectures.append(Lecture)
					else:
						seen_lectures.append(Lecture.attrib['key'])

				for Lecture in list(set(remove_lectures)):
					EvaSys.remove(Lecture)

			dozs_keys = []
			for Lecture in EvaSys.findall('Lecture'):
				for doz in Lecture.find('dozs'):
						dozs_keys.append(doz.find('EvaSysRef').attrib['key'])

			for Person in EvaSys.findall('Person'):
				if not Person.attrib['key'] in dozs_keys:
					EvaSys.remove(Person)

			with codecs.open(output_file_split+'.xml', mode='w', encoding="utf-8") as f:
				f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
				f.write(ElementTree.tostring(EvaSys, encoding='utf-8', method='html').decode("utf-8"))

			if convert_format:
				EvaSys_df = pd.DataFrame(list(iter_docs(EvaSys)))
				try:
					EvaSys_df = EvaSys_df[['key', 'period', 'LSF_ID', 'type', 'name', 'dozs', 'orgroot', 'short', 'turnout', 'dozs_mail', 'p_o_study']]
				except KeyError:
					pass
				if convert_format == "html":
					EvaSys_df.to_html(output_file_split+'.html', justify="center", index=False)
				elif convert_format == "csv":
					EvaSys_df.to_csv(output_file_split+'.csv', sep=";", index=False)
				elif convert_format == "excel":
					with pd.ExcelWriter(output_file_split+'.xlsx', engine='xlsxwriter') as writer:
						EvaSys_df.to_excel(writer, index=False, sheet_name='HISLSF-Export')
						worksheet = writer.sheets['HISLSF-Export']
						worksheet.set_column('D:D', 12, None)
						worksheet.set_column('E:E', 45, None)
						worksheet.set_column('F:F', 25, None)
						worksheet.set_column('G:G', 25, None)
						worksheet.set_column('H:H', 14, None)

	else:
		if convert_format:
			EvaSys = ElementTree.parse(input_file).getroot()
			for Lecture in EvaSys.findall('Lecture'):
				for node in Lecture.getiterator():
					if include_type and node.tag == 'type' and not node.text in include_type:
						EvaSys.remove(Lecture)
					if exclude_type and node.tag == 'type' and node.text in exclude_type:
						EvaSys.remove(Lecture)

			if remove_duplicates:
				remove_lectures = []
				seen_lectures = []

				for Lecture in EvaSys.findall('Lecture'):
					if Lecture.attrib['key'] in seen_lectures:
						remove_lectures.append(Lecture)
					else:
						seen_lectures.append(Lecture.attrib['key'])

				for Lecture in list(set(remove_lectures)):
					EvaSys.remove(Lecture)

			EvaSys_df = pd.DataFrame(list(iter_docs(EvaSys)))
			try:
				EvaSys_df = EvaSys_df[['key', 'period', 'LSF_ID', 'type', 'name', 'dozs', 'orgroot', 'short', 'turnout', 'dozs_mail', 'p_o_study']]
			except KeyError:
				pass
			if convert_format == "html":
				EvaSys_df.to_html(output_file_split+'.html', justify="center", index=False)
			elif convert_format == "csv":
				EvaSys_df.to_csv(output_file_split+'.csv', sep=";", index=False)
			elif convert_format == "excel":
				with pd.ExcelWriter(output_file_split+'.xlsx', engine='xlsxwriter') as writer:
					EvaSys_df.to_excel(writer, index=False, sheet_name='HISLSF-Export')
					worksheet = writer.sheets['HISLSF-Export']
					worksheet.set_column('D:D', 12, None)
					worksheet.set_column('E:E', 45, None)
					worksheet.set_column('F:F', 25, None)
					worksheet.set_column('G:G', 25, None)
					worksheet.set_column('H:H', 14, None)

	print("")
	print("==== EvaSysXML finished ====")

def download_XML(semester):
	# See https://uvweb.uni-muenster.de/lsf/xml/index.php
	url = 'https://uvweb.uni-muenster.de/lsf/xml/evasys.php'
	now = datetime.datetime.now()
	filename = "LSFExport-"+semester+"-"+now.strftime("%Y-%m-%d")+".xml"

	print("Download semester: " + semester)
	print("Save XML data in: " + filename)

	if os.path.isfile(filename) and not os.stat(filename).st_size == 0:
		print("Skip download and use existing file!")
		return filename

	print("Enter your username and password for WWU Münster!")
	username = input("Username: ")
	password = getpass.getpass("Password for " + username + ": ")

	r = requests.post(url, data={'semester':semester}, auth=(username,password), stream=True)
	if r.status_code == 200:
		with open(filename, 'wb') as out:
			for bits in r.iter_content(chunk_size=1024):
				if bits:
					out.write(bits)

	if not os.path.isfile(filename) or os.stat(filename).st_size == 0:
		print("Error (Server or User/Password). Try again!")
		sys.exit(3)

	return filename

if __name__ == "__main__":

	argv = sys.argv[1:]
	convert_format = ''
	split_format = ''
	split_keys = ''
	input_file = ''
	output_file = ''
	include_type = ''
	exclude_type = ''
	remove_duplicates = False
	semester = ''

	try:
		opts, args = getopt.getopt(argv,"i:o:",["download-semester=","convert-to=","ID=","ORG=","include-type=","exclude-type=","remove-duplicates","input=","output=","split-by="])
	except getopt.GetoptError:
		usage()

	for opt, arg in opts:
		if opt in ("--download-semester"):
			semester = arg
		elif opt in ("--convert-to"):
			convert_format = arg
		elif opt in ("--split-by"):
			split_format = arg
		elif opt in ("--ID"):
			split_format = "ID"
			if split_keys:
				usage()
			else:
				split_keys = arg
		elif opt in ("--ORG"):
			split_format = "ORG"
			if split_keys:
				usage()
			else:
				split_keys = arg
		elif opt in ("--remove-duplicates"):
			remove_duplicates = True
		elif opt in ("--include-type"):
			include_type = arg.split(',')
		elif opt in ("--exclude-type"):
			exclude_type = arg.split(',')
		elif opt in ("-i", "--input"):
			input_file = arg
		elif opt in ("-o", "--output"):
			output_file = arg

	if (not input_file and not semester) or (input_file and semester):
		usage()

	if semester:
		input_file = download_XML(semester)

	if convert_format in ["html", "csv", "excel"] or split_format in ["ID", "ORG"]:
		process_XML(convert_format, split_format, split_keys, include_type, exclude_type, remove_duplicates, input_file, output_file)
	else:
		usage()
