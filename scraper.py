import requests
import lxml.html
import scraperwiki
import simplejson as json
from zipfile import ZipFile
from io import BytesIO
import os
import xlrd
import string
from operator import itemgetter


def getColPos(s):
	return string.uppercase.index(s)

def getRowPos(n):
	return n - 1	

def getCellValue(wb,sheet,col,row):
	return wb.sheet_by_name(sheet).cell_value(getRowPos(row),getColPos(col))

def getPercent(val,total):
	if val != 0 and total != 0:
		percent = (val/total*100)
	else:
		percent = 0	
	return percent

religionExclude = ["Christianity:", "Total","Other Religions:","Secular Beliefs and Other Spiritual Beliefs"]
languageExclude = ["Chinese languages:", "Total","Other Religions:"]
ancestryExclude = ['Other(e)','Ancestry not stated']

with open("referenceData/sa2.json") as json_file:
	sa2s = json.load(json_file)

for i, sa2 in enumerate(sa2s):
	url = "http://www.censusdata.abs.gov.au/CensusOutput/copsub2016.NSF/All%20docs%20by%20catNo/2016~Community%20Profile~{sa2}/$File/TSP_{sa2}.zip?OpenElement".format(sa2=str(sa2['SA2_MAIN16']))
	print i
	print "getting", url
	# Fetching the URL with requests
	
	r = requests.get(url, allow_redirects=False)
	print r.status_code

	if r.status_code != 200:
		print "can't get", sa2

	if r.status_code == 200:

		strFile = BytesIO()
		strFile.write(r.content)

		with open('files/{sa2}.zip'.format(sa2=str(sa2['SA2_MAIN16'])), 'wb') as f:
			f.write(r.content)

		input_zip = ZipFile('files/{sa2}.zip'.format(sa2=str(sa2['SA2_MAIN16'])), 'r')
		ex_file = input_zip.open("TSP_" + str(sa2['SA2_MAIN16']) + ".XLS")
		content = ex_file.read()

		wb = xlrd.open_workbook(file_contents=content)

		################################# 2006 ####################################

		data = {}

		data['year'] = 2006
		data['sa2_code'] = sa2['SA2_MAIN16']
		data['sa2_name'] = sa2['SA2_NAME16']

		data['persons'] = getCellValue(wb,'T 01','D',11)
		data['male'] = getCellValue(wb,'T 01','B',11)	
		data['female'] = getCellValue(wb,'T 01','C',11)

		data['percent_male'] = getPercent(data['male'],data['persons'])
		data['percent_female'] = getPercent(data['female'],data['persons'])

		data['median_age'] = getCellValue(wb,'T 02','B',11)
		data['median_household_income'] = getCellValue(wb,'T 02','B',17)
		data['median_mortgage'] = getCellValue(wb,'T 02','G',11)
		data['median_rent'] = getCellValue(wb,'T 02','G',13)
		data['persons_per_bedroom'] = getCellValue(wb,'T 02','G',15)
		data['average_household_size'] = getCellValue(wb,'T 02','G',17)
		
		data['married_males'] = getCellValue(wb,'T 05a','B',29)
		data['married_females'] = getCellValue(wb,'T 05a','C',29)
		data['married_persons'] = data['married_males'] + data['married_females']

		data['defacto_males'] = getCellValue(wb,'T 05a','E',29)
		data['defacto_females'] = getCellValue(wb,'T 05a','F',29)
		data['defacto_persons'] = data['defacto_males'] + data['defacto_females']

		data['notmarried_males'] = getCellValue(wb,'T 05a','H',29)
		data['notmarried_females'] = getCellValue(wb,'T 05a','I',29)
		data['notmarried_persons'] = data['notmarried_males'] + data['notmarried_females']

		data['total_relationship_males'] = getCellValue(wb,'T 05a','K',29)
		data['total_relationship_females'] = getCellValue(wb,'T 05a','L',29)
		data['total_relationship_persons'] = getCellValue(wb,'T 05a','M',29)

		data['percent_married_males'] = getPercent(data['married_males'],data['total_relationship_males'])
		data['percent_married_females'] = getPercent(data['married_females'],data['total_relationship_females'])

		data['percent_married_persons'] = getPercent(data['married_persons'],data['total_relationship_persons'])

		data['percent_defacto_males'] = getPercent(data['defacto_males'],data['total_relationship_males'])
		data['percent_defacto_females'] = getPercent(data['defacto_females'],data['total_relationship_females'])

		data['percent_defacto_persons'] = getPercent(data['defacto_persons'],data['total_relationship_persons'])

		data['percent_notmarried_males'] = getPercent(data['notmarried_males'],data['total_relationship_males'])
		data['percent_notmarried_females'] = getPercent(data['notmarried_females'],data['total_relationship_females'])
		data['percent_notmarried_persons'] = getPercent(data['notmarried_persons'],data['total_relationship_persons'])

		data['indig_males'] = getCellValue(wb,'T 06a','B',27)
		data['indig_females'] = getCellValue(wb,'T 06a','C',27)
		data['indig_persons'] = getCellValue(wb,'T 06a','D',27)

		data['non_indig_males'] = getCellValue(wb,'T 06a','F',27)
		data['non_indig_females'] = getCellValue(wb,'T 06a','G',27)
		data['non_indig_persons'] = getCellValue(wb,'T 06a','H',27)

		data['not_stated_indig_males'] = getCellValue(wb,'T 06a','J',27)
		data['not_stated_indig_females'] = getCellValue(wb,'T 06a','K',27)
		data['not_stated_indig_persons'] = getCellValue(wb,'T 06a','L',27)

		data['total_indig_status_males'] = getCellValue(wb,'T 06a','N',27)
		data['total_indig_status_females'] = getCellValue(wb,'T 06a','O',27)
		data['total_indig_status_persons'] = getCellValue(wb,'T 06a','P',27)

		data['percent_indig_persons'] = getPercent(data['indig_persons'],data['total_indig_status_persons'])

		# Countries of birth is currently set to work with the 2011 census community profile template to generate test data
		# and will need to be changed for 2016

		countriesOfBirth = []

		for x in range(11,45):
			countryItem = {}
			countryItem['label'] = getCellValue(wb,'T 08','A',x).strip()
			countryItem['males'] = getCellValue(wb,'T 08','B',x)
			countryItem['females'] = getCellValue(wb,'T 08','C',x)
			countryItem['persons'] = getCellValue(wb,'T 08','D',x)
			
			# countryItem['persons_percent'] = getPercent(getCellValue(wb,'T 08','D',x), getCellValue(wb,'T 08','D',49))
			countryItem['persons_percent'] = getPercent(getCellValue(wb,'T 08','D',x), getCellValue(wb,'T 08','D',49))

			countriesOfBirth.append(countryItem)

		countriesOfBirth = sorted(countriesOfBirth, key=itemgetter('persons'), reverse=True) 
		
		data['countries_of_birth'] = json.dumps(countriesOfBirth)

		data['born_in_australia'] = getCellValue(wb,'T 08','D',11)
		data['country_not_stated'] = getCellValue(wb,'T 08','D',47)
		data['total_country_persons'] = getCellValue(wb,'T 08','D',49)

		data['born_overseas'] = data['total_country_persons'] - data['born_in_australia'] - data['country_not_stated']

		data['percent_born_overseas'] = getPercent(data['born_overseas'],data['total_country_persons'])


		ancestries = []

		for x in range(12,43):
			if getCellValue(wb,'T 09a','A',x).strip() not in ancestryExclude:
				ancestryItem = {}
				ancestryItem['label'] = getCellValue(wb,'T 09a','A',x).strip()
				ancestryItem['persons'] = getCellValue(wb,'T 09a','G',x)
				ancestryItem['persons_percent'] = getPercent(getCellValue(wb,'T 09a','G',x),getCellValue(wb,'T 09a','G',45))

				ancestries.append(ancestryItem)

		ancestries = sorted(ancestries, key=itemgetter('persons'), reverse=True)

		data['ancestries'] = json.dumps(ancestries)

		languages = []

		for x in range(14,49):
			if getCellValue(wb,'T 10','A',x).strip() not in languageExclude:		
				langItem = {}

				if getCellValue(wb,'T 10','A',x).strip() == "Other(b)":
					langItem['label'] = "Other Chinese"
				else:
					langItem['label'] = getCellValue(wb,'T 10','A',x).strip()	
				langItem['persons'] = getCellValue(wb,'T 10','D',x)
				
				# langItem['persons_percent'] = getPercent(getCellValue(wb,'T 10','D',x),getCellValue(wb,'T 10','D',54))
				langItem['persons_percent'] = getPercent(getCellValue(wb,'T 10','D',x),getCellValue(wb,'T 10','D',54))

				languages.append(langItem)

		languages = sorted(languages, key=itemgetter('persons'), reverse=True)

		data['languages'] = json.dumps(languages)

		data['language_english'] = getCellValue(wb,'T 10','D',11)
		data['language_not_stated'] = getCellValue(wb,'T 10','D',52)
		data['total_language_persons'] = getCellValue(wb,'T 10','D',54)

		data['language_other'] = data['total_language_persons'] - data['language_english'] - data['language_not_stated']

		data['percent_language_other'] = getPercent(data['language_other'],data['total_language_persons'])


		religions = []

		for x in range(13,45):
			# print getCellValue(wb,'T 12a','A',x)
			if getCellValue(wb,'T 12a','A',x).strip() not in religionExclude:		
				religionItem = {}
				religionItem['label'] = getCellValue(wb,'T 12a','A',x).strip()
				religionItem['persons'] = getCellValue(wb,'T 12a','K',x)
				religionItem['persons_percent'] = getPercent(getCellValue(wb,'T 12a','K',x),getCellValue(wb,'T 12a','K',47))
				religions.append(religionItem)

		religions = sorted(religions, key=itemgetter('persons'), reverse=True)

		data['religions'] = json.dumps(religions)

		data['seperate_house'] = getCellValue(wb,'T 15a','H',13)
		data['seperate_house_percent'] = getPercent(getCellValue(wb,'T 15a','H',13),getCellValue(wb,'T 15a','H',36))

		data['semi_or_townhouse'] = getCellValue(wb,'T 15a','H',19)
		data['semi_or_townhouse_percent'] = getPercent(getCellValue(wb,'T 15a','H',19),getCellValue(wb,'T 15a','H',36))

		data['flat_or_unit'] = getCellValue(wb,'T 15a','H',26)
		data['flat_or_unit_percent'] = getPercent(getCellValue(wb,'T 15a','H',26),getCellValue(wb,'T 15a','H',36))

		data['housing_other_or_not_stated'] = getCellValue(wb,'T 15a','H',32) + getCellValue(wb,'T 15a','H',34)
		data['housing_other_or_not_stated_percent'] = getPercent(data['housing_other_or_not_stated'],getCellValue(wb,'T 15a','H',36))

		data['dwelling_owned_outright'] = getCellValue(wb,'T 18a','G',15)
		data['dwelling_owned_outright_percent'] = getPercent(getCellValue(wb,'T 18a','G',15),getCellValue(wb,'T 18a','G',30))

		data['dwelling_owned_mortgage'] = getCellValue(wb,'T 18a','G',16)
		data['dwelling_owned_mortgage_percent'] = getPercent(getCellValue(wb,'T 18a','G',16), getCellValue(wb,'T 18a','G',30))

		data['dwelling_rented'] = getCellValue(wb,'T 18a','G',25)
		data['dwelling_rented_percent'] = getPercent(getCellValue(wb,'T 18a','G',25),getCellValue(wb,'T 18a','G',30))

		data['dwelling_other_or_not_stated'] = getCellValue(wb,'T 18a','G',27) + getCellValue(wb,'T 18a','G',28)
		data['dwelling_other_or_not_stated_percent'] = getPercent(data['dwelling_other_or_not_stated'], getCellValue(wb,'T 18a','G',30))
		# print data
		scraperwiki.sqlite.save(unique_keys=["year","sa2_code"], data=data)


		################################ 2011 ############################################

		data = {}

		data['year'] = 2011
		data['sa2_code'] = sa2['SA2_MAIN16']
		data['sa2_name'] = sa2['SA2_NAME16']

		data['male'] = getCellValue(wb,'T 01','F',11)	
		data['female'] = getCellValue(wb,'T 01','G',11)
		data['persons'] = getCellValue(wb,'T 01','H',11)	

		data['percent_male'] = getPercent(data['male'],data['persons'])
		data['percent_female'] = getPercent(data['female'],data['persons'])

		data['median_age'] = getCellValue(wb,'T 02','C',11)
		data['median_household_income'] = getCellValue(wb,'T 02','C',17)
		data['median_mortgage'] = getCellValue(wb,'T 02','H',11)
		data['median_rent'] = getCellValue(wb,'T 02','H',13)
		data['persons_per_bedroom'] = getCellValue(wb,'T 02','H',15)
		data['average_household_size'] = getCellValue(wb,'T 02','H',17)
		
		data['married_males'] = getCellValue(wb,'T 05a','B',49)
		data['married_females'] = getCellValue(wb,'T 05a','C',49)

		data['defacto_males'] = getCellValue(wb,'T 05a','E',49)
		data['defacto_females'] = getCellValue(wb,'T 05a','F',49)

		data['notmarried_males'] = getCellValue(wb,'T 05a','H',49)
		data['notmarried_females'] = getCellValue(wb,'T 05a','I',49)

		data['total_relationship_males'] = getCellValue(wb,'T 05a','K',49)
		data['total_relationship_females'] = getCellValue(wb,'T 05a','L',49)
		data['total_relationship_persons'] = getCellValue(wb,'T 05a','M',49)

		data['percent_married_males'] = getPercent(data['married_males'],data['total_relationship_males'])
		data['percent_married_females'] = getPercent(data['married_females'],data['total_relationship_females'])

		data['percent_defacto_males'] = getPercent(data['defacto_males'],data['total_relationship_males'])
		data['percent_defacto_females'] = getPercent(data['defacto_females'],data['total_relationship_females'])

		data['percent_notmarried_males'] = getPercent(data['notmarried_males'],data['total_relationship_males'])
		data['percent_notmarried_females'] = getPercent(data['notmarried_females'],data['total_relationship_females'])


		data['married_persons'] = data['married_males'] + data['married_females']
		data['defacto_persons'] = data['defacto_males'] + data['defacto_females']
		data['notmarried_persons'] = data['notmarried_males'] + data['notmarried_females']
		data['percent_married_persons'] = getPercent(data['married_persons'],data['total_relationship_persons'])
		data['percent_defacto_persons'] = getPercent(data['defacto_persons'],data['total_relationship_persons'])
		data['percent_notmarried_persons'] = getPercent(data['notmarried_persons'],data['total_relationship_persons'])	

		data['indig_males'] = getCellValue(wb,'T 06a','B',46)
		data['indig_females'] = getCellValue(wb,'T 06a','C',46)
		data['indig_persons'] = getCellValue(wb,'T 06a','D',46)

		data['non_indig_males'] = getCellValue(wb,'T 06a','F',46)
		data['non_indig_females'] = getCellValue(wb,'T 06a','G',46)
		data['non_indig_persons'] = getCellValue(wb,'T 06a','H',46)

		data['not_stated_indig_males'] = getCellValue(wb,'T 06a','J',46)
		data['not_stated_indig_females'] = getCellValue(wb,'T 06a','K',46)
		data['not_stated_indig_persons'] = getCellValue(wb,'T 06a','L',46)

		data['total_indig_status_males'] = getCellValue(wb,'T 06a','N',46)
		data['total_indig_status_females'] = getCellValue(wb,'T 06a','O',46)
		data['total_indig_status_persons'] = getCellValue(wb,'T 06a','P',46)

		data['percent_indig_persons'] = getPercent(data['indig_persons'],data['total_indig_status_persons'])

		# Countries of birth is currently set to work with the 2011 census community profile template to generate test data
		# and will need to be changed for 2016

		countriesOfBirth = []

		for x in range(11,45):
			countryItem = {}
			countryItem['label'] = getCellValue(wb,'T 08','A',x).strip()
			countryItem['persons'] = getCellValue(wb,'T 08','H',x)
			# countryItem['persons_percent'] = getPercent(getCellValue(wb,'T 08','H',x),getCellValue(wb,'T 08','H',49))
			countryItem['persons_percent'] = getPercent(getCellValue(wb,'T 08','H',x),getCellValue(wb,'T 08','H',49))

			countriesOfBirth.append(countryItem)

		countriesOfBirth = sorted(countriesOfBirth, key=itemgetter('persons'), reverse=True) 
		
		data['countries_of_birth'] = json.dumps(countriesOfBirth)


		data['born_in_australia'] = getCellValue(wb,'T 08','H',11)
		data['country_not_stated'] = getCellValue(wb,'T 08','H',47)
		data['total_country_persons'] = getCellValue(wb,'T 08','H',49)

		data['born_overseas'] = data['total_country_persons'] - data['born_in_australia'] - data['country_not_stated']

		data['percent_born_overseas'] = getPercent(data['born_overseas'],data['total_country_persons'])


		ancestries = []

		for x in range(12,43):
			if getCellValue(wb,'T 09b','A',x).strip() not in ancestryExclude:
				ancestryItem = {}
				ancestryItem['label'] = getCellValue(wb,'T 09b','A',x).strip()
				ancestryItem['persons'] = getCellValue(wb,'T 09b','G',x)
				ancestryItem['persons_percent'] = getPercent(getCellValue(wb,'T 09b','G',x),getCellValue(wb,'T 09b','G',45))

				ancestries.append(ancestryItem)

		ancestries = sorted(ancestries, key=itemgetter('persons'), reverse=True)

		data['ancestries'] = json.dumps(ancestries)

		# Languages is currently set to work with the 2011 census community profile template to generate test data
		# and will need to be changed for 2016

		languages = []

		for x in range(14,49):
			if getCellValue(wb,'T 10','A',x).strip() not in languageExclude:		
				langItem = {}

				if getCellValue(wb,'T 10','A',x).strip() == "Other(b)":
					langItem['label'] = "Other Chinese"
				else:
					langItem['label'] = getCellValue(wb,'T 10','A',x).strip()	
				langItem['persons'] = getCellValue(wb,'T 10','H',x)
				
				# langItem['persons_percent'] = getPercent(getCellValue(wb,'T 10','G',x),getCellValue(wb,'T 10','L',54))

				langItem['persons_percent'] = getPercent(getCellValue(wb,'T 10','H',x),getCellValue(wb,'T 10','H',54))

				languages.append(langItem)

		languages = sorted(languages, key=itemgetter('persons'), reverse=True)

		data['languages'] = json.dumps(languages)

		data['language_english'] = getCellValue(wb,'T 10','H',11)
		data['language_not_stated'] = getCellValue(wb,'T 10','H',52)
		data['total_language_persons'] = getCellValue(wb,'T 10','H',54)

		data['language_other'] = data['total_language_persons'] - data['language_english'] - data['language_not_stated']

		data['percent_language_other'] = getPercent(data['language_other'],data['total_language_persons'])

		religions = []

		for x in range(13,45):
			if getCellValue(wb,'T 12b','A',x).strip() not in religionExclude:		
				religionItem = {}
				religionItem['label'] = getCellValue(wb,'T 12b','A',x).strip()
				religionItem['persons'] = getCellValue(wb,'T 12b','K',x)
				religionItem['persons_percent'] = getPercent(getCellValue(wb,'T 12b','K',x),getCellValue(wb,'T 12b','K',47))
				religions.append(religionItem)

		religions = sorted(religions, key=itemgetter('persons'), reverse=True)

		data['religions'] = str(religions)

		data['seperate_house'] = getCellValue(wb,'T 15b','H',13)
		data['seperate_house_percent'] = getPercent(getCellValue(wb,'T 15b','H',13),getCellValue(wb,'T 15b','H',36))

		data['semi_or_townhouse'] = getCellValue(wb,'T 15b','H',19)
		data['semi_or_townhouse_percent'] = getPercent(getCellValue(wb,'T 15b','H',19),getCellValue(wb,'T 15b','H',36))

		data['flat_or_unit'] = getCellValue(wb,'T 15b','H',26)
		data['flat_or_unit_percent'] = getPercent(getCellValue(wb,'T 15b','H',26),getCellValue(wb,'T 15b','H',36))

		data['housing_other_or_not_stated'] = getCellValue(wb,'T 15b','H',32) + getCellValue(wb,'T 15b','H',34)
		data['housing_other_or_not_stated_percent'] = getPercent(data['housing_other_or_not_stated'],getCellValue(wb,'T 15b','H',36))

		data['dwelling_owned_outright'] = getCellValue(wb,'T 18a','G',34)
		data['dwelling_owned_outright_percent'] = getPercent(getCellValue(wb,'T 18a','G',34) , getCellValue(wb,'T 18a','G',49))

		data['dwelling_owned_mortgage'] = getCellValue(wb,'T 18a','G',35)
		data['dwelling_owned_mortgage_percent'] = getPercent(getCellValue(wb,'T 18a','G',35) , getCellValue(wb,'T 18a','G',49))

		data['dwelling_rented'] = getCellValue(wb,'T 18a','G',44)
		data['dwelling_rented_percent'] = getPercent(getCellValue(wb,'T 18a','G',44) , getCellValue(wb,'T 18a','G', 49))

		data['dwelling_other_or_not_stated'] = getCellValue(wb,'T 18a','G',46) + getCellValue(wb,'T 18a','G',47)
		data['dwelling_other_or_not_stated_percent'] = getPercent(data['dwelling_other_or_not_stated'], getCellValue(wb,'T 18a','G',49))

		scraperwiki.sqlite.save(unique_keys=["year","sa2_code"], data=data)

		############################## 2016 ######################################

		data = {}

		data['year'] = 2016
		data['sa2_code'] = sa2['SA2_MAIN16']
		data['sa2_name'] = sa2['SA2_NAME16']

		data['male'] = getCellValue(wb,'T 01','J',11)	
		data['female'] = getCellValue(wb,'T 01','K',11)
		data['persons'] = getCellValue(wb,'T 01','L',11)	

		data['percent_male'] = getPercent(data['male'],data['persons'])
		data['percent_female'] = getPercent(data['female'],data['persons'])

		data['median_age'] = getCellValue(wb,'T 02','D',11)
		data['median_household_income'] = getCellValue(wb,'T 02','D',17)
		data['median_mortgage'] = getCellValue(wb,'T 02','I',11)
		data['median_rent'] = getCellValue(wb,'T 02','I',13)
		data['persons_per_bedroom'] = getCellValue(wb,'T 02','I',15)
		data['average_household_size'] = getCellValue(wb,'T 02','I',17)
		
		data['married_males'] = getCellValue(wb,'T 05b','B',29)
		data['married_females'] = getCellValue(wb,'T 05b','C',29)

		data['defacto_males'] = getCellValue(wb,'T 05b','E',29)
		data['defacto_females'] = getCellValue(wb,'T 05b','F',29)

		data['notmarried_males'] = getCellValue(wb,'T 05b','H',29)
		data['notmarried_females'] = getCellValue(wb,'T 05b','I',29)

		data['total_relationship_males'] = getCellValue(wb,'T 05b','K',29)
		data['total_relationship_females'] = getCellValue(wb,'T 05b','L',29)
		data['total_relationship_persons'] = getCellValue(wb,'T 05b','M',29)

		data['percent_married_males'] = getPercent(data['married_males'],data['total_relationship_males'])
		data['percent_married_females'] = getPercent(data['married_females'],data['total_relationship_females'])

		data['percent_defacto_males'] = getPercent(data['defacto_males'],data['total_relationship_males'])
		data['percent_defacto_females'] = getPercent(data['defacto_females'],data['total_relationship_females'])

		data['percent_notmarried_males'] = getPercent(data['notmarried_males'],data['total_relationship_males'])
		data['percent_notmarried_females'] = getPercent(data['notmarried_females'],data['total_relationship_females'])


		data['married_persons'] = data['married_males'] + data['married_females']
		data['defacto_persons'] = data['defacto_males'] + data['defacto_females']
		data['notmarried_persons'] = data['notmarried_males'] + data['notmarried_females']
		data['percent_married_persons'] = getPercent(data['married_persons'],data['total_relationship_persons'])
		data['percent_defacto_persons'] = getPercent(data['defacto_persons'],data['total_relationship_persons'])
		data['percent_notmarried_persons'] = getPercent(data['notmarried_persons'],data['total_relationship_persons'])	

		data['indig_males'] = getCellValue(wb,'T 06b','B',27)
		data['indig_females'] = getCellValue(wb,'T 06b','C',27)
		data['indig_persons'] = getCellValue(wb,'T 06b','D',27)

		data['non_indig_males'] = getCellValue(wb,'T 06b','F',27)
		data['non_indig_females'] = getCellValue(wb,'T 06b','G',27)
		data['non_indig_persons'] = getCellValue(wb,'T 06b','H',27)

		data['not_stated_indig_males'] = getCellValue(wb,'T 06b','J',27)
		data['not_stated_indig_females'] = getCellValue(wb,'T 06b','K',27)
		data['not_stated_indig_persons'] = getCellValue(wb,'T 06b','L',27)

		data['total_indig_status_males'] = getCellValue(wb,'T 06b','N',27)
		data['total_indig_status_females'] = getCellValue(wb,'T 06b','O',27)
		data['total_indig_status_persons'] = getCellValue(wb,'T 06b','P',27)

		data['percent_indig_persons'] = getPercent(data['indig_persons'],data['total_indig_status_persons'])

		# Countries of birth is currently set to work with the 2011 census community profile template to generate test data
		# and will need to be changed for 2016

		countriesOfBirth = []

		for x in range(11,45):
			countryItem = {}
			countryItem['label'] = getCellValue(wb,'T 08','A',x).strip()
			countryItem['persons'] = getCellValue(wb,'T 08','L',x)
			
			# countryItem['persons_percent'] = getPercent(getCellValue(wb,'T 08','L',x),getCellValue(wb,'T 08','L',49))
			countryItem['persons_percent'] = getPercent(getCellValue(wb,'T 08','L',x),getCellValue(wb,'T 08','L',49))

			countriesOfBirth.append(countryItem)

		countriesOfBirth = sorted(countriesOfBirth, key=itemgetter('persons'), reverse=True) 
		
		data['countries_of_birth'] = json.dumps(countriesOfBirth)


		data['born_in_australia'] = getCellValue(wb,'T 08','L',11)
		data['country_not_stated'] = getCellValue(wb,'T 08','L',47)
		data['total_country_persons'] = getCellValue(wb,'T 08','L',49)

		data['born_overseas'] = data['total_country_persons'] - data['born_in_australia'] - data['country_not_stated']

		data['percent_born_overseas'] = getPercent(data['born_overseas'],data['total_country_persons'])


		ancestries = []

		for x in range(12,43):
			if getCellValue(wb,'T 09c','A',x).strip() not in ancestryExclude:	
				ancestryItem = {}
				ancestryItem['label'] = getCellValue(wb,'T 09c','A',x).strip()
				ancestryItem['persons'] = getCellValue(wb,'T 09c','G',x)
				ancestryItem['persons_percent'] = getPercent(getCellValue(wb,'T 09c','G',x),getCellValue(wb,'T 09c','G',45))

				ancestries.append(ancestryItem)

		ancestries = sorted(ancestries, key=itemgetter('persons'), reverse=True)

		data['ancestries'] = json.dumps(ancestries)	

		languages = []

		for x in range(14,49):
			if getCellValue(wb,'T 10','A',x).strip() not in languageExclude:		
				langItem = {}

				if getCellValue(wb,'T 10','A',x).strip() == "Other(b)":
					langItem['label'] = "Other Chinese"
				else:
					langItem['label'] = getCellValue(wb,'T 10','A',x).strip()	
				langItem['persons'] = getCellValue(wb,'T 10','L',x)

				# langItem['persons_percent'] = getPercent(getCellValue(wb,'T 10','G',x),getCellValue(wb,'T 10','L',54))

				langItem['persons_percent'] = getPercent(getCellValue(wb,'T 10','L',x),getCellValue(wb,'T 10','L',54))

				languages.append(langItem)

		languages = sorted(languages, key=itemgetter('persons'), reverse=True)

		data['languages'] = json.dumps(languages)

		
		data['language_english'] = getCellValue(wb,'T 10','L',11)
		data['language_not_stated'] = getCellValue(wb,'T 10','L',52)
		data['total_language_persons'] = getCellValue(wb,'T 10','L',54)

		data['language_other'] = data['total_language_persons'] - data['language_english'] - data['language_not_stated']

		data['percent_language_other'] = getPercent(data['language_other'],data['total_language_persons'])

		religions = []

		for x in range(13,45):
			if getCellValue(wb,'T 12c','A',x).strip() not in religionExclude:		
				religionItem = {}
				religionItem['label'] = getCellValue(wb,'T 12c','A',x).strip()
				religionItem['persons'] = getCellValue(wb,'T 12c','K',x)
				religionItem['persons_percent'] = getPercent(getCellValue(wb,'T 12c','K',x),getCellValue(wb,'T 12c','K',47))
				religions.append(religionItem)

		religions = sorted(religions, key=itemgetter('persons'), reverse=True)

		data['religions'] = json.dumps(religions)

		data['seperate_house'] = getCellValue(wb,'T 15c','H',13)
		data['seperate_house_percent'] = getPercent(getCellValue(wb,'T 15c','H',13),getCellValue(wb,'T 15c','H',36))

		data['semi_or_townhouse'] = getCellValue(wb,'T 15c','H',19)
		data['semi_or_townhouse_percent'] = getPercent(getCellValue(wb,'T 15c','H',19),getCellValue(wb,'T 15c','H',36))

		data['flat_or_unit'] = getCellValue(wb,'T 15c','H',26)
		data['flat_or_unit_percent'] = getPercent(getCellValue(wb,'T 15c','H',26),getCellValue(wb,'T 15c','H',36))

		data['housing_other_or_not_stated'] = getCellValue(wb,'T 15c','H',32) + getCellValue(wb,'T 15c','H',34)
		data['housing_other_or_not_stated_percent'] = getPercent(data['housing_other_or_not_stated'],getCellValue(wb,'T 15c','H',36))

		data['dwelling_owned_outright'] = getCellValue(wb,'T 18b','G',15)
		data['dwelling_owned_outright_percent'] = getPercent(getCellValue(wb,'T 18b','G',15) , getCellValue(wb,'T 18b','G',30))

		data['dwelling_owned_mortgage'] = getCellValue(wb,'T 18b','G',16)
		data['dwelling_owned_mortgage_percent'] = getPercent(getCellValue(wb,'T 18b','G',16) , getCellValue(wb,'T 18b','G',30))

		data['dwelling_rented'] = getCellValue(wb,'T 18b','G',25)
		data['dwelling_rented_percent'] = getPercent(getCellValue(wb,'T 18b','G',25) , getCellValue(wb,'T 18b','G', 30))

		data['dwelling_other_or_not_stated'] = getCellValue(wb,'T 18b','G',27) + getCellValue(wb,'T 18b','G',28)
		data['dwelling_other_or_not_stated_percent'] = getPercent(data['dwelling_other_or_not_stated'], getCellValue(wb,'T 18b','G',30))

		scraperwiki.sqlite.save(unique_keys=["year","sa2_code"], data=data)

		print "done"