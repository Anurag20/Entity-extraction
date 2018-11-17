# coding: utf-8
import os
import cv2
import numpy as np
import re
import usaddress
import json
import shutil
import sys
from io import StringIO
import pandas as pd
import numpy as np
import ast
from pandas import ExcelWriter
import FileLocation as fl
from dateutil import parser


def findWholeWord(w):
	try:
		# print("126 function------------------")
		return re.compile(r'\b({0})\b'.format(w), flags=re.IGNORECASE).search
	except Exception as e:
		pass


def validate_date(date_text):
	try:
		dt = parser.parse(date_text)
		#		print("1")
		if re.search("[A-Za-z]", date_text):
			#			print("2")
			return 0
		elif re.search("^[0-9]+$", date_text):
			#			print("3")
			return 0
		elif re.search(":", date_text):
			#			print("4")
			return 0
		# print("5")
		return 1
	except Exception as e:
		#		print("6")
		return 0


def validate_zipcode(zipcode):
	try:
		if re.search("[A-Za-z]", zipcode):
			return 0
		else:
			return 1
	except Exception as e:
		return 0


def validate_agency(agency_text):
	try:
		if re.search("(inc)|(llc)|(agency)", agency_text, flags=re.I):
			return 1
		else:
			return 0
	except Exception as e:
		# print(e)
		return 0


def split_addresses(address):
	try:
		validate_ins = 0
		insured_namef = ""
		streetf = ""
		cityf = ""
		statef = ""
		ZipCodef = ""
		address = usaddress.parse(address)
		for value, key in address:
			if key == "ZipCode":
				ZipCodef = value
				break
			else:
				if key == "Recipient" and validate_ins != 1:
					insured_namef = insured_namef + value + " "
					# print(":::::::::::::::::",value,validate_ins)
					if re.search('^inc$|^llc$|^inc\.$|^llc\.$', value.strip(), flags=re.I):
						# print(";;;;;;;;;;;;;;;;;",value,validate_ins)
						validate_ins = 1
				elif key == "AddressNumber" or key == 'AddressNumberPrefix' or key == 'AddressNumberSuffix' or key == 'BuildingName' or key == 'CornerOf' or key == 'IntersectionSeparator' or key == 'LandmarkName' or key == "StreetName" or key == "StreetNamePostType" or key == "OccupancyIdentifier" or key == "StreetNamePreDirectional" or key == 'OccupancyType' or key == 'StreetNamePreModifier' or key == 'StreetNamePreType' or key == 'StreetNamePostDirectional' or key == 'StreetNamePostModifier' or key == 'SubaddressIdentifier' or key == 'SubaddressType' or key == 'USPSBoxGroupID' or key == 'USPSBoxGroupType' or key == 'USPSBoxID' or key == 'USPSBoxType':
					streetf = streetf + value + " "
				elif key == "PlaceName":
					cityf = cityf + value + " "
				elif key == "StateName":
					statef = value
		return insured_namef, streetf, cityf, statef, ZipCodef
	except Exception as e:
		# print(e)
		return insured_namef, streetf, cityf, statef, ZipCodef


def backend_process(rootOutputDir, ocr_output_directory_path, ocr_output_file_name, form_name, new_form_flag,
					form_count, pdf_completion_flag):
	# print("Attribute extrction")
	file = ocr_output_file_name
	file_name = file
	pdf_name = ocr_output_directory_path.split('/')[-2]
	email_name = ocr_output_directory_path.split('/')[-3]
	json_path = ocr_output_directory_path + 'Entity_Output' + "/"
	if pdf_completion_flag:
		ui_files_path = fl.ui_path
		if not os.path.exists(ui_files_path):
			os.mkdir(ui_files_path)

		# shutil.copy(,ui_files_path)
		for path, subdirs, files in os.walk(ocr_output_directory_path):
			# print("Json Path >>>>>> 1", json_path)
			# print("Output Dir path >>>>>> 1", ocr_output_directory_path)
			x = os.path.join(json_path, files[(form_count)])
			if os.path.exists(x):
				for file in files:
					if (file.split(".")[-1] == 'PDF' or file.split(".")[-1] == 'pdf'):
						# print("Found PDF")
						for path, subdirs, files2 in os.walk(json_path):
							for file123 in files2:
								if (file123.split(".")[-1] == 'json' or file123.split(".")[-1] == 'JSON'):
									# print("Found JSON")
									os.chmod(ocr_output_directory_path + file, 777)
									shutil.copy(ocr_output_directory_path + file, ui_files_path)
									break
							break
						break

			if os.path.exists(x):
				with open(x) as data_file:
					element = json.load(data_file)

					if 'contour_flag' in element.keys():
						del element['contour_flag']
						del element['flag_agency']
						del element['flag_applicant']
						del element['flag_date']
						del element['CONFIDENCE_AGENCY']
						del element['CONFIDENCE_INSUREDNAME']
						del element['CONFIDENCE_STREET1']
						del element['CONFIDENCE_STREET2']
						del element['CONFIDENCE_CITY']
						del element['CONFIDENCE_STATE']
						del element['CONFIDENCE_DATE']
						del element['CONFIDENCE_ZIPCODE']
					data_file.close()

		return 1

	if not os.path.exists(json_path):  # writing to json watever entities were available
		os.makedirs(json_path)
	if form_count > 0 and new_form_flag:  # going back to old json and chcking if all contours were found or not
		for path, subdirs, files in os.walk(json_path):
			# print("Json Path >>>>>>2 ", json_path)
			x = os.path.join(json_path, files[form_count - 1])
			# print(x)
			if os.path.exists(x):
				with open(x) as data_file:
					element = json.load(data_file)

					if 'contour_flag' in element.keys():
						del element['contour_flag']
						del element['flag_agency']
						del element['flag_applicant']
						del element['flag_date']
						del element['CONFIDENCE_AGENCY']
						del element['CONFIDENCE_INSUREDNAME']
						del element['CONFIDENCE_STREET1']
						del element['CONFIDENCE_STREET2']
						del element['CONFIDENCE_CITY']
						del element['CONFIDENCE_STATE']
						del element['CONFIDENCE_DATE']
						del element['CONFIDENCE_ZIPCODE']
					data_file.close()

	# print(file)
	# fil = open(source_path + '/' + file, "r", encoding="utf8")
	# print('ocr op dir path',ocr_output_directory_path)
	# print('ocr op fname',ocr_output_file_name)
	fil = open(ocr_output_directory_path + ocr_output_file_name + ".txt", "r", encoding="utf8")

	x = fil.readlines()
	#    #print(x)
	str123 = ''
	for i in x:
		if not i == '\n':
			str123 = str123 + i
	# #print(str123)

	x = re.split("__+", str123)
	#    #print(x)

	s = [" ", '', "  ", "   ", "    ", "     ", None]
	address = ""
	street = ""
	if new_form_flag:

		flag_agency = 0
		flag_applicant = 0
		flag_street1 = 0
		flag_street2 = 0
		flag_date = 0
		agency_name = ""
		confidence_agency_name = 0
		insured_name = ""
		insured_name126 = ""
		confidence_insured_name = 0
		street_1 = ""
		street_2 = ""
		confidence_street1 = 0
		confidence_street2 = 0
		city = ""
		confidence_city = 0
		state = ""
		confidence_state = 0
		address = ""
		str1 = ""
		ZipCode = ""
		confidence_ZipCode = 0
		date = ""
		confidence_date = 0
		dic = {}
		confidence_applicant = 0
		applicant_contour = []
		insured_name125 = ""

	else:
		for path, subdirs, files in os.walk(json_path):

			# print("Json Path >>>>>>3 ", json_path)
			x = os.path.join(json_path, files[form_count - 1])
			# print(x)
			if os.path.exists(x):
				with open(x) as data_file:
					data = json.load(data_file)
					for element in data:
						contour_flag = element['contour_flag']
						flag_agency = element['flag_agency']
						flag_applicant = element['flag_applicant']
						flag_date = element['flag_date']
						agency_name = element['AGENCY']
						confidence_agency_name = element['CONFIDENCE_AGENCY']
						insured_name = element['INSUREDNAME']
						confidence_insured_name = element['CONFIDENCE_INSUREDNAME']
						street_1 = element['STREET1']
						street_2 = element['STREET2']
						confidence_street1 = element['CONFIDENCE_STREET1']
						confidence_street2 = element['CONFIDENCE_STREET2']
						city = element['CITY']
						confidence_city = element['CONFIDENCE_CITY']
						state = element['STATE']
						confidence_state = element['CONFIDENCE_STATE']
						date = element['DATE']
						confidence_date = element['CONFIDENCE_DATE']
						ZipCode = element['ZIPCODE']
						confidence_ZipCode = element['CONFIDENCE_ZIPCODE']

	str1 = ""
	dic = {}
	try:
		for i in range(len(x)):
			#    #print(i,"  " ,r[i])
			if flag_date != 1 and re.search("date", x[i], flags=re.I):
				r = re.split("\n", x[i])
				r = [k for k in r if k not in s]
				# print(r)
				for j in range(len(r)):
					if re.search("[^policy]|[^proposed]|[^effective]", r[j], flags=re.I):
						##print('inside date 0')
						if re.search("^date", r[j].replace(" ", ""), flags=re.I):
							##print('inside date 1')
							if len(r[j:len(r) - 1]) > 1:
								# print('inside date 2')
								dte = re.sub("[A-Za-z]", "", r[j + 1])
								dte = dte.replace(" ", "")
								if validate_date(dte):
									# print('inside date 3')
									date = r[j + 1].replace(" ", "")
								if date:
									confidence_date = float(
										re.split("confidence", r[len(r) - 1], flags=re.I)[1].strip())
									flag_date = 1
								break
			if flag_agency != 1 and re.search("(agency)|(producer)", x[i].replace(' ', ''), flags=re.I):
				r = re.split("\n", x[i])
				r = [k for k in r if k not in s]
				# print(r)
				for j in range(len(r)):
					# print(r[j],len(r))
					# print('inside agency 0')
					r[j] = r[j].strip(' .,;:?|\/""()[]{}')
					if re.search("^agency|^producer", r[j].replace(' ', ''), flags=re.I):
						if re.search("agencycustomer|agencybill", r[j].replace(" ", ""), flags=re.I):
							break
						# print(r[j])
						# print('inside agency 1')
						if len(r[j:len(r) - 1]) > 1:
							# print('inside agency 2')
							# print(r[j])
							if re.search("(applicant)|(underwriter)", r[j + 1], flags=re.I) and len(
									r[j + 1:len(r) - 1]) > 1:
								# print('inside agency 3')
								agency_name = r[j + 2]
							else:
								# print('inside agency 3')
								agency_name = r[j + 1]
							if agency_name:
								if re.search("(^\.|^\d|^\+|^=)(\d+|\s+\d+)", agency_name.strip()):
									agency_name = ""
									break
								if re.search("(customer)|(bill)|(CLISTOMER)", agency_name, flags=re.I):
									agency_name = ""
									break
								# print("without space ",agency_name)
								#								confidence_agency_name=float(re.split("confidence",r[len(r)-1],flags=re.I)[1].strip())
								#								flag_agency = 1
								if len(agency_name) > 1:
									# print("without space ",agency_name)
									confidence_agency_name = float(
										re.split("confidence", r[len(r) - 1], flags=re.I)[1].strip())
									flag_agency = 1
									break
								else:
									agency_name = ""
							break
			if flag_applicant != 1 and re.search("applicantinformation", x[i].replace(' ', ''), flags=re.I):
				# print('inside applicant 0')
				r = re.split("\n", x[i])
				r = [k for k in r if k not in s]
				# print(r)
				#            #print("Applicant info  ",r)
				for j in range(len(r)):
					# print('inside applicant 1')
					if re.search("^applicantinformation", r[j].replace(' ', ''), flags=re.I) and len(
							r[j:len(r) - 1]) > 1:
						# print('inside applicant 2')
						if re.search("name|mame|bame|mane", r[j + 1].replace(" ", ""), flags=re.I) and re.search(
								"mailing", r[j + 1].replace(" ", ""), flags=re.I) and len(r[j + 1:len(r) - 1]) > 1:
							# print('inside applicant 3')
							#                        #print("Applicant info  ",r)
							str1 = " ".join(r[j + 2:len(r) - 1])
							#							address = usaddress.parse(str1)
							# print("name and mailing ",str1)
							if str1:
								confidence_applicant = float(
									re.split("confidence", r[len(r) - 1], flags=re.I)[1].strip())
								applicant_contour = r
								flag_applicant = 1
							break
						elif re.search("name|mame|bame|mane", r[j + 1], flags=re.I) and len(r[j + 1:len(r) - 1]) > 1:
							# print('inside applicant 4')
							insured_name125 = r[j + 2]
							if insured_name125:
								confidence_applicant = float(
									re.split("confidence", r[len(r) - 1], flags=re.I)[1].strip())
								flag_applicant = -1
							# print("name ",insured_name125)
							break
						elif len(r[j:len(r) - 1]) > 1:
							#                        #print("Applicant info  ",r)
							# print('inside applicant 5')
							str1 = " ".join(r[j + 1:len(r) - 1])
							#							address = usaddress.parse(str1)
							if str1:
								confidence_applicant = float(
									re.split("confidence", r[len(r) - 1], flags=re.I)[1].strip())
								applicant_contour = r
								flag_applicant = 1
							# print("NO name and mailing ",str1)
							break
		# #print("Address :",address)
		for i in range(len(x)):
			if flag_agency != 1 and re.search("(agency)|(producer)", x[i], flags=re.I):
				# print("------------------------\n",x[i])
				r = re.split("AGENCY", x[i])
				# print(r)
				r = r[1]
				r = re.split("\n", r)
				# print(r)
				r = [k for k in r if k not in s]
				# print(r)
				if len(r[:len(r) - 1]) > 0:
					# print('inside agency space 0')
					if not re.search("(customer)|(bill)|(CLISTOMER)", r[0].replace(" ", ""), flags=re.I):
						# print('inside agency space 1')
						agency_name = r[0]
					if agency_name:
						# print('inside agency space 2')
						confidence_agency_name = float(re.split("confidence", r[len(r) - 1], flags=re.I)[1].strip())
						flag_agency = 1
					# print("agency space",agency_name)
			if flag_applicant == -1:
				# print('inside mailing 0')
				if re.search("mailingaddress", x[i].replace(" ", ""), flags=re.I):
					# print('inside mailing 1')
					r = re.split("\n", x[i])
					r = [k for k in r if k not in s]
					for j in range(len(r)):
						# print('inside mailing 2')
						if re.search("^mailingaddress", r[j].replace(' ', ''), flags=re.I) and len(r[j:len(r) - 1]) > 1:
							# print('inside mailing 3')
							#                        #print("Applicant info  ",r)
							str1 = " ".join(r[j:len(r) - 1])
							#							address = usaddress.parse(str1)
							if str1:
								confidence_applicant = (confidence_applicant + float(
									re.split("confidence", r[len(r) - 1], flags=re.I)[1].strip())) / 2
								applicant_contour = r
								flag_applicant = 1
							# print("Mailing ",str1)
			if '126' in form_name and flag_applicant == 0:

				# print(r)
				if re.search("applicant", x[i].replace(" ", ""), flags=re.I):
					r = re.split("\n", x[i])
					r = [k for k in r if k not in s]
					for j in range(len(r) - 1):
						if len(r[j:len(r) - 1]) > 1:
							# print('inside applicant space 0')
							x1 = findWholeWord('LLC')(r[j + 1])
							x2 = findWholeWord('INC')(r[j + 1])
							#						x3 = findWholeWord('INC')(r[j])
							#						x4 = findWholeWord('LLC')(r[j])
							if x1 or x2:
								# print('inside applicant space 1')
								#							#print("126 if condition---------------------------")
								#							if len(r[j:len(r) - 1]) > 1:
								#							#print('inside applicant space 2')
								# print(r[j])
								insured_name126 = r[j + 1]
								if insured_name126:
									# print("without space ", insured_name126)
									confidence_applicant = float(
										re.split("confidence", r[len(r) - 1], flags=re.I)[1].strip())
									flag_applicant = 1
								#						elif x3 or x4:
								#							print('inside applicant space 2')
								#							#print("126 elif condition---------------------------")
								#							if len(r[j:len(r) - 1]) > 0:
								#								insured_name126 = re.split("applicant", r[j], flags=re.I)[1].strip()
								#								if insured_name126:
								#									#print("without space ", insured_name126)
								#									confidence_applicant = float(re.split("confidence", r[len(r) - 1], flags=re.I)[1].strip())
								#									flag_applicant = 1
		if flag_applicant == -1:
			flag_applicant = 1

		if flag_applicant == 0:
			for i in range(len(x)):
				# print('inside mailing2 0')
				if re.search("mailingaddress", x[i].replace(" ", ""), flags=re.I) and re.search("[^name|^bame|^mame]",
																								r[j].replace(" ", ""),
																								flags=re.I):
					# print('inside mailing2 1')
					r = re.split("\n", x[i])
					r = [k for k in r if k not in s]
					for j in range(len(r)):
						# print('inside mailing2 2')
						if re.search("^mailingaddress", r[j].replace(' ', ''), flags=re.I) and len(r[j:len(r) - 1]) > 1:
							# print('inside mailing2 3')
							#                        #print("Applicant info  ",r)
							str1 = " ".join(r[j:len(r) - 1])
							#							address = usaddress.parse(str1)
							if str1:
								confidence_applicant = (confidence_applicant + float(
									re.split("confidence", r[len(r) - 1], flags=re.I)[1].strip())) / 2
								applicant_contour = r
								flag_applicant = 1
							# print("Mailing2 ",str1)

			'''if flag_applicant==0:
				if re.search("applicant",x[i],flags=re.I):
					r = re.split("\n", x[i])
					r = [k for k in r if k not in s]'''

		'''for value, key in address:
			if key == "ZipCode":
				ZipCode=value
				break
			else:
				if key == "Recipient":
					insured_name = insured_name + value + " "
				elif key == "AddressNumber" or key=='AddressNumberPrefix' or key=='AddressNumberSuffix' or key=='BuildingName' or key=='CornerOf' or key=='IntersectionSeparator'or key=='LandmarkName' or key == "StreetName" or key == "StreetNamePostType" or key == "OccupancyIdentifier" or key == "StreetNamePreDirectional" or key=='OccupancyType' or key=='StreetNamePreModifier' or key=='StreetNamePreType' or key=='StreetNamePostDirectional' or key=='StreetNamePostModifier' or key=='SubaddressIdentifier' or key=='SubaddressType' or key=='USPSBoxGroupID' or key=='USPSBoxGroupType' or key=='USPSBoxID' or key=='USPSBoxType':
					street = street + value + " "
				elif key == "PlaceName":
					city = city + value + " "
				elif key == "StateName":
					state = value'''

		insured_name, street, city, state, ZipCode = split_addresses(str1)
		if insured_name126:
			insured_name = insured_name126
		if insured_name125:
			insured_name = insured_name125
		# print("+++++++++++++","i",insured_name,"s",street,"c",city,"s",state,"z",ZipCode)
		#		street_1=street
		#		street_2=street
		if applicant_contour:
			for i in applicant_contour[:len(applicant_contour) - 1]:
				a, b, c, d, e = split_addresses(str(i))
				# print("---------------a",a,"b",b,"c",c,"d",d,"e",e)
				if b and re.search(b.replace(" ", ""), street.replace(" ", ""), flags=re.I):
					if not flag_street1:
						street_1 = b
						flag_street1 = 1
					# print("street1 ",street_1)
					elif not flag_street2:
						street_2 = b
						flag_street2 = 1
						# print("street2 ",street_2)
						break

		agency_name = agency_name.strip(' ,;:?|\/()')
		insured_name = insured_name.strip(' ,;:?|\/()')
		street_1 = street_1.strip(' .,;:?|\/()')
		street_2 = street_2.strip(' .,;:?|\/()')
		city = city.strip(' .,;:?|\/()')
		state = state.strip(' .,;:?|\/()')
		ZipCode = ZipCode.strip(' .,;:?|\/()')
		date = date.strip(' .,;:?|\/()')

		if len(state) != 2:
			state = ""
		if len(street_1) < 2:
			street_1 = ""
		if len(street_2) < 2:
			street_2 = ""
		if not validate_zipcode(ZipCode):
			ZipCode = ""
		if len(insured_name) < 2:
			insured_name = ""

		if confidence_applicant != 0:
			if insured_name:
				confidence_insured_name = confidence_applicant
			if street_1:
				confidence_street1 = confidence_applicant
			if street_2:
				confidence_street2 = confidence_applicant
			if city:
				confidence_city = confidence_applicant
			if state:
				confidence_state = confidence_applicant
			if ZipCode:
				confidence_ZipCode = confidence_applicant


			#		agency_name= agency_name.strip(' ,;:?|\/()')
			#		insured_name = insured_name.strip(' ,;:?|\/()')
			#		street_1= street_1.strip(' .,;:?|\/()')
			#		street_2= street_2.strip(' .,;:?|\/()')
			#		city= city.strip(' .,;:?|\/()')
			#		state = state.strip(' .,;:?|\/()')
			#		ZipCode=ZipCode.strip(' .,;:?|\/()')
			#		date=date.strip(' .,;:?|\/()')

		dic["AGENCY"] = agency_name
		dic["INSUREDNAME"] = insured_name
		dic["STREET1"] = street_1
		dic["STREET2"] = street_2
		dic["CITY"] = city
		dic["STATE"] = state
		dic["ZIPCODE"] = ZipCode
		dic["DATE"] = date

		# print(dic)
		z = (flag_agency + flag_applicant + flag_date) / 3
		# print(z)
		contour_flag = z
		if new_form_flag:
			newFile = open(json_path + file_name.split('.txt')[0] + '.json', "w")
			newFile.write(json.dumps(dic))
			newFile.close()
			fil.close()
		if not contour_flag:

			dic['flag_agency'] = flag_agency
			dic['flag_applicant'] = flag_applicant
			dic['flag_date'] = flag_date
			if agency_name:
				dic['CONFIDENCE_AGENCY'] = confidence_agency_name
			else:
				dic['CONFIDENCE_AGENCY'] = 0
			if insured_name:
				dic['CONFIDENCE_INSUREDNAME'] = confidence_insured_name
			else:
				dic['CONFIDENCE_INSUREDNAME'] = 0
			if street_1:
				dic['CONFIDENCE_STREET1'] = confidence_street1
			else:
				dic['CONFIDENCE_STREET1'] = 0
			if street_2:
				dic['CONFIDENCE_STREET2'] = confidence_street2
			else:
				dic['CONFIDENCE_STREET2'] = 0
			if city:
				dic['CONFIDENCE_CITY'] = confidence_city
			else:
				dic['CONFIDENCE_CITY'] = 0
			if state:
				dic['CONFIDENCE_STATE'] = confidence_state
			else:
				dic['CONFIDENCE_STATE'] = confidence_state
			if ZipCode:
				dic['CONFIDENCE_ZIPCODE'] = confidence_ZipCode
			else:
				dic['CONFIDENCE_ZIPCODE'] = 0
			dic['CONFIDENCE_DATE'] = confidence_date

			dic['contour_flag'] = contour_flag
			newFile = open(json_path + file_name.split('.txt')[0] + '.json', "w")
			newFile.write(json.dumps(dic))
			newFile.close()
			fil.close()

		# df1 = pd.DataFrame(columns = ['File_Name','Form_No.','Attr1-AGENCY','Attr1-Confidence','Attr2-INSUREDNAME','Attr2-Confidence','Attr3-STREET','Attr3-Confidence','Attr4-CITY','Attr4-Confidence','Attr5-STATE','Attr5-Confidence','Attr6-DATE','Attr6-Confidence','Attr7-ZIPCODE','Attr7-Confidence'])
		if contour_flag:
			# writing to consolidated excel

			##print(df1)
			if confidence_agency_name == 0:
				confidence_agency_name = ""
			if confidence_applicant == 0:
				confidence_applicant = ""
			if confidence_city == 0:
				confidence_city = ""
			if confidence_date == 0:
				confidence_date = ""
			if confidence_insured_name == 0:
				confidence_insured_name = ""
			if confidence_state == 0:
				confidence_state = ""
			if confidence_street1 == 0:
				confidence_street1 = ""
			if confidence_street2 == 0:
				confidence_street2 = ""
			if confidence_ZipCode == 0:
				confidence_ZipCode = ""

			excel_path = rootOutputDir
			data = [email_name, pdf_name, ocr_output_file_name, form_name, agency_name, confidence_agency_name,
					insured_name, confidence_insured_name, street_1, confidence_street1, street_2, confidence_street2,
					city, confidence_city, state, confidence_state, date, confidence_date, ZipCode, confidence_ZipCode]
			df1 = pd.DataFrame([data],
							   columns=['email_name', 'Pdf_name', 'File_Name', 'Form_No', 'AGENCY', 'Agency_Confidence',
										'INSUREDNAME', 'InsuredName_Confidence', 'STREET_LINE_1', 'Street1_Confidence',
										'STREET_LINE_2', 'Street2_Confidence', 'CITY', 'City_Confidence', 'STATE',
										'State_Confidence', 'DATE', 'Date_Confidence', 'ZIPCODE', 'Zipcode_Confidence'])
			df = pd.read_excel(excel_path + 'Consolidated.xlsx')
			writer = ExcelWriter(excel_path + 'Consolidated.xlsx')
			df = df.append(df1)
			df = df.reset_index(drop=True)
			df.to_excel(writer, "sheet-1")
			writer.save()

			'''if new_form_flag:
				newFile = open(json_path + file_name.split('.txt')[0] + '.json', "w")
				newFile.write(json.dumps(dic))
				newFile.close()
				fil.close()'''

			if not contour_flag:
				# print("Json Path >>>>>>4 ", json_path)
				x = os.path.join(json_path, files[(form_count)])
				if os.path.exists(x):
					with open(x) as data_file:
						element = json.load(data_file)

						if 'contour_flag' in element.keys():
							del element['contour_flag']
							del element['flag_agency']
							del element['flag_applicant']
							del element['flag_date']
							del element['CONFIDENCE_AGENCY']
							del element['CONFIDENCE_INSUREDNAME']
							del element['CONFIDENCE_STREET1']
							del element['CONFIDENCE_STREET2']
							del element['CONFIDENCE_CITY']
							del element['CONFIDENCE_STATE']
							del element['CONFIDENCE_DATE']
							del element['CONFIDENCE_ZIPCODE']
						data_file.close()
				newFile = open(json_path + file_name.split('.txt')[0] + '.json', "w")
				newFile.write(json.dumps(dic))
				newFile.close()
				fil.close()


	except Exception as e:
		pass

	return contour_flag

	# df1 = pd.DataFrame(columns=['File_Name', 'Form_No.', 'Attr1-AGENCY', 'Attr1-Confidence', 'Attr2-INSUREDNAME',
	# 							'Attr2-Confidence', 'Attr3-STREET', 'Attr3-Confidence', 'Attr4-CITY',
	# 							'Attr4-Confidence', 'Attr5-STATE', 'Attr5-Confidence', 'Attr6-DATE',
	# 							'Attr6-Confidence', 'Attr7-ZIPCODE', 'Attr7-Confidence'])
	# df1['Form_No.'] = form_name
	# json_path = ocr_output_directory_path + "/" + 'Entity_Output' + "/"
	# for path, subdirs, files in os.walk(json_path):
	# 	with open(files[0]) as data_file:
	# 		df1['File_Name'] = files[0].strip('.')[0]
	# 		data = json.load(data_file)
	# 		for element in data:
	# 			df1['Attr1-AGENCY'] = element['AGENCY']
	# 			df1['Attr1-Confidence'] = element['CONFIDENCE_AGENCY']
	# 			df1['Attr2-INSUREDNAME'] = element['INSUREDNAME']
	# 			df1['Attr2-Confidence'] = element['CONFIDENCE_INSUREDNAME']
	# 			df1['Attr3-STREET'] = element['STREET']
	# 			df1['Attr3-Confidence'] = element['CONFIDENCE_STREET']
	# 			df1['Attr4-CITY'] = element['CITY']
	# 			df1['Attr4-Confidence'] = element['CONFIDENCE_CITY']
	# 			df1['Attr5-STATE'] = element['STATE']
	# 			df1['Attr5-Confidence'] = element['CONFIDENCE_STATE']
	# 			df1['Attr6-DATE'] = element['DATE']
	# 			df1['Attr6-Confidence'] = element['CONFIDENCE_DATE']
	# 			df1['Attr7-ZIPCODE'] = element['ZIPCODE']
	# 			df1['Attr7-Confidence'] = element['CONFIDENCE_ZIPCODE']
	# 			del element['flag_agency']
	# 			del element['flag_applicant']
	# 			del element['flag_date']
	# 			del element['CONFIDENCE_AGENCY']
	# 			del element['CONFIDENCE_INSUREDNAME']
	# 			del element['CONFIDENCE_STREET']
	# 			del element['CONFIDENCE_CITY']
	# 			del element['CONFIDENCE_STATE']
	# 			del element['CONFIDENCE_DATE']
	# 			del element['CONFIDENCE_ZIPCODE']
	# 			writer = ExcelWriter(json_path + 'Consolidated.xlsx')
	# 			df1.to_excel(writer,"sheet-1")
	# 			writer.save()
	# 			#print(df1)
	#
	# return 1
