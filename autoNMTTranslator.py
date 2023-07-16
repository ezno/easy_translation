#-*- coding:utf-8-*-
'''
* version 7.0 (date: 23.6.24)
	- Add DeepL API 
	- Add Glossary part
* version 6.0 (date: 19.07.22)
	- Add Google SpreadSheet API functions to share glossaries with other translators
* version 5.0 (date: 18.06.27)
	- Add Naver OpenApi
	 -> Client ID : wd49u739v3
	 -> Client Secret Key : 4CmfKfoTUsYfgoEJ8ZunZuz2GOUkPZIsGyiij3zW
* version 4.0 (date : 18.02.01)
	- Style #i# adds
* version 3.0 (date : 18.01.18)
	- Make functions for string modification
		- pre_string_changer, post_sting_changer, glossary_notifier
* version 2.0 (date : 17.12.15)
   - Add KaKao NMT translatr beta
   `-> Link(https://translate.kakao.com/)
   - Add Naver N2MT tralator(a.k.a PAPAGO) API
   	-> Client ID : 3yrYMArz_dA_9rEVRzur
   	-> Client Secret : hamJErs1qE

* author: EZNO(ezno84@gmail.com)
'''
import configparser
import re
import sys
import json
import http.client
import httplib2
import urllib.parse
import requests
import os.path
import pickle
import deepl
import argparse
import csv
import time
import base64
import hmac
import hashlib
from argparse import ArgumentParser, ArgumentDefaultsHelpFormatter

try:
	from textblob import TextBlob
	from textblob.compat import PY2, request, urlencode
except ImportError:
	raise ImportError('[Error] need textblob library.')

try:
	from googleapiclient.discovery import build
	from google.oauth2.credentials import Credentials
	from googleapiclient import discovery
	from googleapiclient.errors import HttpError
	from google_auth_oauthlib.flow import InstalledAppFlow
	from google.auth.transport.requests import Request	
	#from apiclient import discovery
	from google.cloud import translate_v3 as translate
	from google.cloud import storage
except ImportError:
	raise ImportError('[Error] need Google client API.')

# Global constant variables
# - Google Sheet, Goole trnalsation API, Naver Papago, DeepL
class G:
	# Dash board Google Sheet ID
	# For 'Machine learning in R' Google translation URL address is below
	# https://docs.google.com/spreadsheets/d/1F4AgXexD5xe6K0-ga0DfUxk9bn9xLCc4JBenDVfKJCw/edit#gid=0
	SPREADSHEET_ID = 'GARBAGEVALUE'
	SPREADSHEET_API_KEY = 'GARBAGEVALUE'

	# Google Translation API Key
	GOOGLE_TRNS_API_KEY = 'GARBAGEVALUE'

	# Google Translation Glossary related constants
	GT_LOCATION = ('us-central1')
	GT_PROJECT_ID = ('eznopub-142102')
	GT_GLOSSARY_PRJ_ID = ('655216823621')
	GT_GLOSSARY_ID = ('ML_in_R_202306')
	GT_GLOSSARY_FILE_NAME = ('ML_in_R.csv')

	# NAVER PAPAGO translation API Key	
	PAPAGO_CLIENT_ID = 'GARBAGEVALUE'
	PAPAGO_CLIENT_SECRET = 'GARBAGEVALUE'
	PAPAGO_GLOSSARY_ID = 'GARBAGEVALUE'
	NCLOUD_CLINET_ID = 'GARBAGEVALUE'
	NCLOUD_SECRET_KEY = 'GARBAGEVALUE'

	# NAVER PAPAGO translation spare API Key - by Yena
	''' 
	client_id = "rGvQAlQZKPflGRXFQE_8"
	client_secret = "lb04NeyJLY"
	'''
	# DeepL translation API Key
	DEEPL_AUTH_KEY = 'GARBAGEVALUE'

# Common string manipulation
#  - Replace garbage characters, wrong characters
def pre_string_changer(modified_text):
	pre_word_dictionary = {
		":\n" : ":.\n",
		"ﬁ" : "fi",
		"ﬁ " : "fi",
		"ﬂ " : "fl",
		'“' : '"',
		'”' : '"',
		"•" : " ",
		r"\s\s" : " ",
		r"\s-" : "",
		r"^\d+$" : "",
		r"\f" : "",
		'launchd' : '"launchd"',
		'powerd' : '"powerd"',
		'(cid:123)' : '',
	}

	for key,value in pre_word_dictionary.items():
		if modified_text.find(key) != -1:
			modified_text = re.sub( key, value, modified_text )
	
	return modified_text

# Common string manipulation
#  - After translating English to Koream, automatically replace Koream
def post_string_changer(modified_text):
	postWordDictionary = {
		u"합니다": u"한다",
		u"됩니다": u"된다",
		u"됩니다": u"된다",
		u"습니다": u"다",
		u"만듭니다": u"만든다",
		u"줍니다": u"준다",
		u"입니다": u"이다",
		u"낸니다": u"낸다",
		u"닙니다": u"니다",
		u'전자 메일' : u'이메일',
		u'지역' : u'리전',
		u'매개 변수' : u'파라미터',
		u'HackerOne' : u'해커원',
		u'Facebook' : u'페이스북',
		u'Twitter' : u'트위터',
	}

	for key, value in postWordDictionary.items():
		if modified_text.find(key) != -1:
			modified_text = re.sub( key, value, modified_text )

	return modified_text


def post_string_sanitizer(origText):

	modifiedText = origText

	postReplaceDictionary = {
		u"&quot;" : '"',
		u"&amp;" : '&',
		u"&gt;" : '>',
		u"&lt;" : '<',
		u"&#39;" : '"',
		u"와이어 샤크" : u"와이어샤크",
		u"브로드 캐스팅" : u"브로드캐스팅",
		u"멀티 캐스팅" : u"멀티캐스팅",
		u"루프 백" : u"루프백",
		u"플로" : u"플로우",
		u"프락시" : u"프록시",
		u'스크린숏' : u'스크린샷',
		u"명령 줄" : u"명령줄",
		u"명령행" : u"명령줄",
		u"설루션" : u"솔루션",	
		u'윈도' : u'윈도우',
		u'에러' : u'오류',
		u'진입 전' : u'진입점',
		u'WordPress' : u'워드프레스',
	}

	for key, value in postReplaceDictionary.items():
		if origText.find(key) != -1:
			modifiedText = re.sub( key, value, modifiedText )

	return modifiedText


def naver_spell_check(q):
	params = {'_callback': 'window.__jindo2_callback._spellingCheck_0', 'q': q}

	headers = {
		"User-Agent": "Mozilla/5.0 (iPad; CPU OS 11_0 like Mac OS X) AppleWebKit/604.1.34 (KHTML, like Gecko) Version/11.0 Mobile/15A5341f Safari/604.1",
		"Content-type": "application/x-www-form-urlencoded; charset=UTF-8", 
		"Accept": "application/javascript, */*;q=0.8",
		"Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
		"Host": "m.search.naver.com",
		"Cookie": "npic=BUJFh0UHQ/Vui5UVdABc/8pO/h93QNOHTB0SsspX37LyufL7PrrceWUDeK6M/l6wCA==; NNB=JNJFCFYP7HNVU; _ga=GA1.2.1110102603.1524441957; ASID=afdf1eaa00000162f0ab8bd60000005c; nx_open_so=1; nid_iplevel=1; nx_ssl=2; BMR=s=1532836444251&r=https%3A%2F%2Fpost.naver.com%2Fviewer%2FpostView.nhn%3FvolumeNo%3D16336242&r2=https%3A%2F%2Fwww.naver.com%2F; _naver_usersession_=QXYOfxkrxVqJdpL2rRPB/g==; page_uid=T2ctAlpVuENssuIyegdssssss4C-514218",
		"Referer": "https://search.naver.com/search.naver?sm=tab_hty.top&where=nexearch&query=%EB%84%A4%EC%9D%B4%EB%B2%84+%EB%9D%84%EC%96%B4%EC%93%B0%EA%B8%B0+%EA%B2%80%EC%82%AC%EA%B8%B0&oquery=in+conjunction+with&tqi=T2ct4spVuE0ssvTTu3RssssssrG-171059",
	}

	jsonp_text = requests.get("https://m.search.naver.com/p/csearch/ocontent/spellchecker.nhn", params=params, headers=headers).text
	jsonp_text = jsonp_text.replace(params['_callback'] + '(', '')
	json_text = jsonp_text.replace(');', '')

	try:
		result = json.loads(json_text)
		print ('* Errota Cnt : ' + str(result['message']['result']['errata_count']))
		result_html = result['message']['result']['html']
		result_text = re.sub(r'<\/?.*?>', '', result_html)
	except:
		print ('* Skip Spell check.')
		result_text = q

	return post_string_sanitizer(result_text)


def pre_string_sanitizer(origText):
	modifiedText = origText

	preReplaceDictionary = {
		r'> ' : '',
		r'●' : '',
		r'○' : '',
	}

	for key, value in preReplaceDictionary.items():
		if origText.find(key) != -1:
			modifiedText = re.sub( key, value, modifiedText )

	if( (re.search("#h#", modifiedText) or re.search("#d#", modifiedText) or re.search("#d2#", modifiedText)) and modifiedText[0:1] == ' '):
		modifiedText = modifiedText[1:]
	
	return modifiedText


# For Google Translation v3 - glossary ======================================
# 	- Get glossary data from DashBoard Google Sheet
#	- Create a CSV file on the client, 
def get_glossary_from_sheet():
	discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')

	service = discovery.build(
		'sheets',
		'v4',
		http=httplib2.Http(),
		discoveryServiceUrl=discoveryUrl,
		developerKey=G.SPREADSHEET_API_KEY,
	)
	sheet_range = 'glossary!A:B'

	result = service.spreadsheets().values().get(
				spreadsheetId=G.SPREADSHEET_ID,
				range = sheet_range,
			).execute()

	values = result.get('values', [])

	glossary_dictionary = {}

	if not values:
		print('\t[ERROR] No data found. Check Google spreadsheets!')
	else:
		glossary_dictionary = {
			values[i][0].lower() : values[i][1] for i in range(1, len(values)) 
		}
		csv_content = "\n".join([",".join(row) for row in values])

	return glossary_dictionary, csv_content


def dict_to_csv(dictionary, filename):
    # Create a list of key-value pairs
    rows = [(key, value) for key, value in dictionary.items() if key is not None and value is not None]

    # Create a CSV file and write the key-value pairs
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['en', 'ko'])  # Write column headers
        writer.writerows(rows)  # Write key-value pairs

    print(f"\tDictionary successfully converted and saved as '{filename}'.")

#  - Upload CSV file on the Google Cloud bucket
#  - Create Google translation glossary using above file 
def create_google_nmt_glossary_on_GCloud(csv_content, glossary_dict_cnt):
	# 1.Upload CSV file on the Google Cloud bucket
	storage_client = storage.Client(project=G.GT_PROJECT_ID)
	buckets = storage_client.list_buckets()

	for bucket in buckets:
		if G.GT_PROJECT_ID in bucket.name:
			glossary_bucket_name = bucket.name

	print(f'\tBucket name: {glossary_bucket_name}')

	glossary_bucket = storage_client.get_bucket(glossary_bucket_name)
	blob = bucket.blob(G.GT_GLOSSARY_FILE_NAME)
	blob.upload_from_string(csv_content, content_type='text/csv')

	print(f'\tUplaod glossary csv file to GCloud storage as a {G.GT_GLOSSARY_FILE_NAME} file')
	print(f'\tMaking Google Translation glossary')

	# 2.Create Google translation glossary using above file 
	client = translate.TranslationServiceClient()

	name = client.glossary_path(
	    G.GT_PROJECT_ID,
	    G.GT_LOCATION,
	    G.GT_GLOSSARY_ID)

	# gs://eznopub-142102.appspot.com/Functional programming - glossary.csv
	gcloud_uri = f'gs://{G.GT_PROJECT_ID}.appspot.com/{G.GT_GLOSSARY_FILE_NAME}'

	language_codes_set = translate.types.Glossary.LanguageCodesSet(
	    language_codes=['en', 'ko'])

	gcs_source = translate.types.GcsSource(
	    input_uri=gcloud_uri)

	input_config = translate.types.GlossaryInputConfig(
	    gcs_source=gcs_source)

	glossary = translate.types.Glossary(
	    name=name,
	    language_codes_set=language_codes_set,
	    input_config=input_config )

	parent = f'projects/{G.GT_GLOSSARY_PRJ_ID}/locations/{G.GT_LOCATION}'

	glosssary_input_uri, glossary_cnt = list_glossaries(client, parent, glossary)

	#if (glosssary_input_uri is not None) and (len(glossary_dict_cnt)-1) != glossary_cnt:
	if glosssary_input_uri is not None:
		delete_all_glossaries()
	
	operation = client.create_glossary(parent=parent, glossary=glossary)

	result = operation.result(timeout=90)
	print(f'\tCreated GCloud glossary: {result.name}')
	print(f'\tGCloud glossary Input Uri: {result.input_config.gcs_source.input_uri}')


def list_glossaries(client, parent, glossary):

	response = client.list_glossaries(parent=parent)

	for glossary in client.list_glossaries(parent=parent):
		print(f"\tGlossary name: {glossary.name}")
		print(f"\tInput URI: {glossary.input_config.gcs_source.input_uri}")
		print(f"\tEntry count: {glossary.entry_count}")

	return glossary.input_config.gcs_source.input_uri, glossary.entry_count


def delete_all_glossaries():
    client = translate.TranslationServiceClient()

    parent = f"projects/{G.GT_GLOSSARY_PRJ_ID}/locations/{G.GT_LOCATION}"

    response = client.list_glossaries(parent=parent)

    for glossary in response.glossaries:
        glossary_name = glossary.name
        client.delete_glossary(name=glossary_name)

        print(f"\tDeleted glossary: {glossary_name}")


#	- Load glossary for translation
def get_google_nmt_glossaray():
	client = translate.TranslationServiceClient()

	glossary_name = client.glossary_path(
	        G.GT_PROJECT_ID, 
	        G.GT_LOCATION,
	        G.GT_GLOSSARY_ID
	    )

	glossary_resp = client.get_glossary(name=glossary_name)

	return glossary_resp


# For Google spread sheet ======================================
#  - Find words in the string based on the glossary
#	- write down glossaries on the Google Sheet
# (Updated 2023.6.27) - This features added on Google spread sheet by Donghee

def glossary_notifier(inText, glossary):
	grepWord = u''

	inText = inText.lower() # chagne to lowercase
	'''
	glossaryDict = {
		# chapter 1 -------------------------------------------------
		u'region' : u'리전',
		#==============================================
	}
	'''
	for wordKey, wordValue in glossary.items():
		if inText.find(wordKey) != -1:
			grepWord = grepWord + wordKey + ' : ' + wordValue + ', '

	return grepWord


# Core part 
# 	- Load raw text and strings
#  - Translate raw text from Kakao, Naver Papago, Google translation v3, DeepL
# - Writ
#def translator(text_path, pageNum, glossary_config):
def translator(text_path, pageNum):
	# Record translation result on thie file in the client 
	TranText = 'trns_' + text_path

	fpin = open(text_path, 'r+')
	fpout = open(TranText, 'w+')

	try:
		if os.path.isfile(fpout):
			print('* transltated file backup : ' + fpout)
			cmd = 'copy "%s"  "%s"' %(fpout, 'bk'+fpout)
			os.system(cmd.encode('cp949'))
	except:
		pass	

	# Load text file to translate
	orig_text = fpin.read()
	origLines = orig_text.split('\n\n')

	textBlock = {
		'source' : '',
		'naverTrns' : 'naver :',
		'googleTrns' : '',
		'kakaoTrns' : 'kakao :',
		'deepLTrns' : '',
		'style' : '',
	}

	# Get Google Spread sheet id from Dashbaord
	spreadsheetId = get_sheet_id(pageNum)

	for orig_text in origLines:
		# Sanitize string 
		modifiedText = pre_string_changer(orig_text)

		if modifiedText != '':
			# Get Style information from the string
			textBlock['style'] = modifiedText[-3:]
			sentence_index = 0 
			
			if textBlock['style'] == '#p#' or textBlock['style'] == '#h#' or textBlock['style'] == '#d#':
				modifiedText = modifiedText[0:-3]

				# To get strings from the paragraph use TextBlob libary
				zen = TextBlob(modifiedText)
				for sentence in zen.sentences:
					sentence_index = sentence_index + 1
					
					# Updated 2023.6.27 - This features added on Google spread sheet by Donghee
					#wordMemResult = glossary_notifier(sentence.raw, glossary)
					#fpout.write('< ' + str(wordMemResult) + '\n')
					sentence.raw = sentence.raw.replace('\n', '')
					#textBlock['source'] += ' ' + sentence.raw
					textBlock['source'] = sentence.raw

					# NaverNMT2API results: English > Korean
					#textBlock['naverTrns'] = naver_neural_machine2_translate(sentence.raw, 'en', 'ko')
					textBlock['naverTrns'] = papago_translate(sentence.raw, 'en', 'ko')
		
					# Google NeuralMachineTranlate results: English > Korean
					textBlock['googleTrns'] = google_neural_machine_translate_v3(sentence.raw, 'en', 'ko') + ' '

					# KaKao translator beta: English > Korean Only
					textBlock['kakaoTrns'] = kakao_neural_machine2_translate(sentence.raw, 'en', 'ko')

					# DeepL translator: English > Korean
					# DeepL doesn't support Enlish, Korean glossary yet
					textBlock['deepLTrns'] = deepL_translate(sentence.raw, 'en', 'ko')
				
					fpout.write(textBlock['source'] + '\n')
					fpout.write(textBlock['naverTrns'] + '\n')
					fpout.write(textBlock['kakaoTrns'] + '\n')
					fpout.write(textBlock['deepLTrns'] + '\n')

					if len(zen.sentences) == sentence_index:
						fpout.write('> ' + textBlock['googleTrns'] + textBlock['style'] + '\n\n')
					else:
						fpout.write('> ' + textBlock['googleTrns'] + '\n\n')

					write_on_google_sheets(
						pageNum, 
						sentence.raw, 
						textBlock['googleTrns'],
						'',
						textBlock['deepLTrns'],
						textBlock['naverTrns'],
						textBlock['kakaoTrns'],
						textBlock['style'],
						sentence_index,
						spreadsheetId
					)

			elif textBlock['style'] == '#i#':
				fpout.write('[그림 시작]\n[그림 종료]\n')
				sentence_index = sentence_index + 1
				write_on_google_sheets(
						pageNum, 
						'', 
						'[스크린샷 삽입]',
						'',
						'',
						'',
						'',
						textBlock['style'],
						sentence_index,
						spreadsheetId
					)				

			elif textBlock['style'] == '#c#':
				sentence_index = sentence_index + 1
				orig_text = orig_text[0:-3]
				write_on_google_sheets(
						pageNum, 
						'', 
						'[코드 시작]'+orig_text+'\n[코드 종료]',
						'',
						'',
						'',
						'',
						textBlock['style'],
						sentence_index,
						spreadsheetId
					)		

			else:
				sentence_index = sentence_index + 1
				orig_text = orig_text[0:-3]
				write_on_google_sheets(
						pageNum, 
						'', 
						orig_text,
						'',
						'',
						'',
						'',
						textBlock['style'],
						sentence_index,
						spreadsheetId
					)				

				orig_text = re.sub(r"\n", "\n> ", orig_text)
				fpout.write('> ' + orig_text + textBlock['style'] + '\n\n')

			for key, value in textBlock.items():
				textBlock[key] = ''
	
	fpin.close()
	fpout.close()
	print('[INFO] Translating Complete.')

# Old version for Google translation
def google_translate(source, from_lang, to):
	# Configuration for Google tanslate API.
	# input your developer Key form Google Cloud Platfrom. <http://code.google.com/apis/console>
	service = build('translate', 'v2', developerKey=G.GOOGLE_TRNS_API_KEY)
	source = source.encode('utf-8')
	request = service.translations().list(q=source , source=from_lang, target=to)
	response = request.execute()

	return response['translations'][0]['translatedText']

# Old version for Google translation
def google_neural_machine_translate(source, from_lang, to):
	trnsURL = "https://translation.googleapis.com/language/translate/v2"
	source = source.encode('utf-8')
	reqValues = {
			"q": source,
			"source": from_lang,
			"target": to,
			"key" : G.GOOGLE_TRNS_API_KEY,
			"model" : "nmt",
		}

	reqData = urllib.parse.urlencode(reqValues)
	reqURL = trnsURL + '?' + reqData
	r = requests.get(reqURL)
	responseData = json.loads(r.content)

	#print responseData
	if r.status_code == 200:
		translatedTxt = responseData['data']['translations'][0]['translatedText']
	else:
		translatedTxt = '\t[ERROR] fail googleNMT. check-out'
	r.connection.close()

	return translatedTxt

# For Google Translation v3 
# 	- Can use glossaries for translation
def google_neural_machine_translate_v3(source, from_lang, to):
	from google.cloud import translate

	client = translate.TranslationServiceClient()
	parent = f'projects/{G.GT_GLOSSARY_PRJ_ID}/locations/{G.GT_LOCATION}'

	glossary = client.glossary_path(G.GT_PROJECT_ID, "us-central1", G.GT_GLOSSARY_ID)

	glossary_config = translate.TranslateTextGlossaryConfig(glossary=glossary)

	result = client.translate_text(
			request= {
				"contents": [source],
				"source_language_code": from_lang,
				"target_language_code": to,
				"parent": parent,
				"glossary_config": glossary_config,
			}
	)

	translated_txt = ''

	for glossary_translations in result.glossary_translations:
		results_text = glossary_translations.translated_text
		translated_txt = translated_txt + results_text

	if translated_txt == '':
		translated_txt = '\t[Error] Failed to translate by Google NMT v3'

	return post_string_changer(translated_txt)


# NAVER Papago translation - free version
# Need free translation API keys - client ID, secrey
def naver_neural_machine2_translate(source, from_lang, to):
	params = urllib.parse.urlencode({
			'source':from_lang,
			'target':to,
			'text':source,
		})

	params = params.encode('utf-8')

	headers = {
			"X-Naver-Client-Id": G.PAPAGO_CLIENT_ID,
			"X-Naver-Client-Secret": G.PAPAGO_CLIENT_SECRET,
			"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
		}

	conn = http.client.HTTPSConnection("openapi.naver.com")
	conn.request("POST","/v1/papago/n2mt", params, headers)
	
	response = conn.getresponse()
	data = json.loads(response.read())

	try:
		translated_txt = data['message']['result']['translatedText']
	except:
		translated_txt = 'Fail.'
	conn.close()

	return translated_txt

# NAVER Papago translation - paid version
# Can work with glossary 
def make_signature(access_key, secret_key, timestamp, url, method):
	timestamp = str(timestamp)
	secret_key = bytes(secret_key, 'UTF-8')

	message = method + " " + url + "\n" + timestamp + "\n" + access_key
	message = bytes(message, 'UTF-8')
	signingKey = base64.b64encode(hmac.new(secret_key, message, digestmod=hashlib.sha256).digest())
	return signingKey.decode('UTF-8')


def papago_glossary_upload():
	baseurl = "https://papago.apigw.ntruss.com"
	url = "/glossary/v1/{}/upload"
	url = url.format(G.PAPAGO_GLOSSARY_ID)

	timestamp = int(time.time() * 1000)
	method = "POST"

	signature = make_signature(G.NCLOUD_CLINET_ID, G.NCLOUD_SECRET_KEY, timestamp, url, method)

	url = baseurl + url
	headers = {
		"x-ncp-apigw-timestamp": str(timestamp),
		"x-ncp-iam-access-key": G.NCLOUD_CLINET_ID,
		"x-ncp-apigw-signature-v2": str(signature),
	}
	file = {
		"file": (	G.GT_GLOSSARY_FILE_NAME, 
						open(G.GT_GLOSSARY_FILE_NAME, 'rb'), 
						"text/csv"	)
	}

	print(f'[INFO] Update Papago glossary data from Google sheet')
	response = requests.post(url=url, verify=True, headers=headers, files=file)
	data = json.loads(response.content.decode('utf-8'))
	print(f"\tPapago glossary ID: {data['data']['glossaryKey']}")

	return data['data']['glossaryKey']


# NAVER Papago translation - paid version
def papago_translate(source, from_lang, to):

	params = urllib.parse.urlencode({
			'source':from_lang,
			'target':to,
			'text':source,
			'glossaryKey': G.PAPAGO_GLOSSARY_ID,
		})

	params = params.encode('utf-8')

	headers = {
			"X-NCP-APIGW-API-KEY-ID": G.PAPAGO_CLIENT_ID,
			"X-NCP-APIGW-API-KEY": G.PAPAGO_CLIENT_SECRET,
			"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
		}

	conn = http.client.HTTPSConnection("naveropenapi.apigw.ntruss.com")
	conn.request("POST","/nmt/v1/translation", params, headers)
	
	response = conn.getresponse()
	data = json.loads(response.read())

	try:
		translated_txt = data['message']['result']['translatedText']
	except:
		translated_txt = 'Fail.'
	conn.close()

	return translated_txt


# KAKAO neural translation
def kakao_neural_machine2_translate(source, from_lang, to):
	params = {
		'queryLanguage': 'auto', 
		'resultLanguage' : 'kr', 
		'q': source,
	}

	headers = {
		"User-Agent": "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36, iLoveKakao",
		"Content-type": "application/x-www-form-urlencoded; charset=UTF-8", 
		"Accept": "*/*",
		"Accept-Language": "ko-KR,ko;q=0.8,en-US;q=0.6,en;q=0.4",
		"Host": "translate.kakao.com",
		"Cookie": "_ga=GA1.2.1565432128.1524450128; webid=fba36550468e11e8ab13000af759d260; _kadu=QLG8AipYIeRhrmtH_1557285958669; TIARA=LwBejXLW1_fntGreDRUPhJ8uu5QQO77G6s5R6_dAOb.vEGy5XOySMKOeaHiS4unoWWz5sm3vOvN5lldDmeUUJYlfzH5gptWc",
		"Accept-Encoding": "gzip, deflate, br",
		"Referer": "https://translate.kakao.com/",
	}

	json_text = requests.post(
							"https://translate.kakao.com/translator/translate.json", 
							params=params, 
							headers=headers
						).text

	data = json.loads(json_text)

	try:
		translated_txt = data['result']['output'][0][0]
	except:
		translated_txt = 'Fail.'

	return translated_txt

# DeepL API
def deepL_translate(source, from_lang, to):
	deepl_translator = deepl.Translator(G.DEEPL_AUTH_KEY) 
	deepl_result = deepl_translator.translate_text(source, target_lang=to)

	return deepl_result.text

#	- In DashBoard google sheet, in the 'sheetID' column, write down spreadSheet ID 
#	- get spreadSheet ID based on page
def get_sheet_id(currentPage):
	discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
	service = discovery.build(
		'sheets',
		'v4',
		http=httplib2.Http(),
		discoveryServiceUrl=discoveryUrl,
		developerKey=G.SPREADSHEET_API_KEY,
	)
	sheet_range = 'chapter!A:G'

	result = service.spreadsheets().values().get(
			spreadsheetId=G.SPREADSHEET_ID,
			range = sheet_range,
			).execute()

	values = result.get('values', [])

	if not values:
		print('\t[ERROR] Google Docs ID for this Chapter not found. Check Google spreadsheets!')
		return ''

	else:
		for i in range(1, len(values)):
			ch_num = values[i][0].encode('utf-8')
			ch_title = values[i][4].encode('utf-8')
			ch_start_page = values[i][5].encode('utf-8')
			ch_end_page = values[i][6].encode('utf-8')

			if int(currentPage) in range(int(ch_start_page), int(ch_end_page)):
				sheetID = values[i][2]
				print(f'\t[INFO] This page corresponds to - Chapter{ch_num.decode("utf-8")} : {ch_title.decode("utf-8")}')
				print(f'\t[INFO] spreadsheets ID: {sheetID}')

				return sheetID


def get_google_docs_sheets_id(currentCh):
	DISCOVERY_URL = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
	service = discovery.build(
		'sheets',
		'v4',
		http=httplib2.Http(),
		discoveryServiceUrl=DISCOVERY_URL,
		developerKey=G.SPREADSHEET_API_KEY,
	)
	sheet_range = 'chapter!A:C'

	result = service.spreadsheets().values().get(
			spreadsheetId=G.SPREADSHEET_ID,
			range = sheet_range,
			).execute()

	values = result.get('values', [])

	if not values:
		print('\t[Error] Can not get Google Docs ID for this chapter Check the google spreadsheet!')
		return '', ''
	else:
		for i in range(0, len(values)):
			if values[i][0] == str(currentCh):
				docs_id = values[i][1]
				sheets_id = values[i][2]

				return docs_id, sheets_id


def write_on_google_sheets(page, source, google, glossary, deepl, papago, kakao, style, sentenceNum, spreadsheetId):

	discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
	SCOPES = [
		'https://www.googleapis.com/auth/drive',
		'https://www.googleapis.com/auth/drive.file',
		'https://www.googleapis.com/auth/spreadsheets',
		'https://www.googleapis.com/auth/documents',
	]
	creds = None

	if os.path.exists('token.json'):
		creds = Credentials.from_authorized_user_file('token.json', SCOPES)

	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
			creds = flow.run_local_server(port=0)

		with open('token.json', 'w') as token:
			token.write(creds.to_json())

	service = build('sheets', 'v4', credentials=creds)

	sheet_range = 'segments!A:I'

	values = [
		[ page, source, google, glossary, deepl, papago, kakao, style, sentenceNum ],
	]

	body = {
		'values': values
	}

	print(body)

	service.spreadsheets().values().append(
		spreadsheetId=spreadsheetId, 
		range=sheet_range,
		valueInputOption='USER_ENTERED', 
		body=body).execute()


def write_google_docs(in_text, docsId):
	
	SCOPES = [
		'https://www.googleapis.com/auth/drive.file',
		'https://www.googleapis.com/auth/documents',
		'https://www.googleapis.com/auth/drive',
	]	
	
	requestsDoc = [
		{
				"insertText": {
					"endOfSegmentLocation": {
						"segmentId": "",
					},
				"text": in_text,
			}
		}
	]

	creds = None

	if os.path.exists('token.json'):
		creds = Credentials.from_authorized_user_file('token.json', SCOPES)

	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
			creds = flow.run_local_server(port=0)

	# Save the credentials for the next run
		with open('token.pickle', 'w') as token:
			pickle.dump(creds, token)

	service = build(
					'docs', 
					'v1', 
					credentials=creds,
				)
	# Retrieve the documents contents from the Docs service.
	document = service.documents().batchUpdate(
		documentId=docsId, 
		body={'requests': requestsDoc},
	).execute()


def make_translation_documents(sheets_id, docs_id):

	discovery_url = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
	SCOPES = [
		'https://www.googleapis.com/auth/drive',
		'https://www.googleapis.com/auth/drive.file',
		'https://www.googleapis.com/auth/spreadsheets',
		'https://www.googleapis.com/auth/documents',
	]
	creds = None

	'''
	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token, encoding="bytes")
		# If there are no (valid) credentials available, let the user log in.
	'''

	if os.path.exists('token.json'):
		creds = Credentials.from_authorized_user_file('token.json', SCOPES)


	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
			creds = flow.run_local_server(port=0)
	# Save the credentials for the next run

		'''
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)
		'''
	with open('token.json', 'w') as token:
		token.write(creds.to_json())
	
	service = build('sheets', 'v4', credentials=creds)

	sheet_range = 'segments!A:I'

	result = service.spreadsheets().values().get(
			spreadsheetId=sheets_id,
			range = sheet_range,
			).execute()

	translated_values = result.get('values', [])

	paragraph_text = ''

	if not translated_values:
		print('\t[!]No data found. Check Google spreadsheets!')
	else:
		for i in range(1, len(translated_values)):
			sentence = translated_values[i][2]
			style = translated_values[i][7]
			sentence_index = translated_values[i][8]

			if sentence_index == u'1' or i == len(translated_values):
				paragraph_text = '\n' + paragraph_text + '\n'
				write_google_docs(paragraph_text, docs_id)
				paragraph_text = ''

			if style == u'#c#':
				paragraph_text = paragraph_text + sentence
			else:
				replaced_text = post_string_sanitizer(sentence)
				speelchecked_text = naver_spell_check(replaced_text)
				paragraph_text = paragraph_text + speelchecked_text

			time.sleep(5)


if __name__ == '__main__':
	#reload(sys)
	#sys.setdefaultencoding("utf8")

	# Load API keys and variables 
	config = configparser.ConfigParser()
	config.read('config.ini')

	# Access the API keys from the configuration file
	G.SPREADSHEET_ID = config.get('API_KEYS', 'SPREADSHEET_ID')
	G.SPREADSHEET_API_KEY = config.get('API_KEYS', 'SPREADSHEET_API_KEY')
	G.GOOGLE_TRNS_API_KEY = config.get('API_KEYS', 'GOOGLE_TRNS_API_KEY')
	# For DeepL translation API
	G.DEEPL_AUTH_KEY = config.get('API_KEYS', 'DEEPL_AUTH_KEY')
	# For Papago translation API
	#	- Papago
	G.PAPAGO_CLIENT_ID = config.get('API_KEYS', 'PAPAGO_CLIENT_ID')
	G.PAPAGO_CLIENT_SECRET = config.get('API_KEYS', 'PAPAGO_CLIENT_SECRET')
	G.PAPAGO_GLOSSARY_ID = config.get('API_KEYS', 'PAPAGO_GLOSSARY_ID')
	#	- Glossary update
	G.NCLOUD_CLINET_ID = config.get('API_KEYS', 'NCLOUD_CLINET_ID')
	G.NCLOUD_SECRET_KEY = config.get('API_KEYS', 'NCLOUD_SECRET_KEY')

	# For argument parsing
	parser = argparse.ArgumentParser(
		description = "Translate Automator V8.1 (2023.07.16.)",
		formatter_class = ArgumentDefaultsHelpFormatter )

	subparsers = parser.add_subparsers(dest="command", help="Available commands")
	# Two options for this script
	# Select [1]. translator
	parser_trans = subparsers.add_parser("trns", help="Process transcripts")
	parser_trans.add_argument("-sp", "--start-page", type=int, help="Start page")
	parser_trans.add_argument("-ep", "--end-page", type=int, help="End page")

	# Select [2]. create translation goole doc
	parser_docs = subparsers.add_parser("docs", help="Process documents")
	parser_docs.add_argument("-ch", "--chapter", type=int, help="Chapter number")

	args = parser.parse_args()

	# Handle Option [1] - translator
	if args.command == "trns":
		if args.start_page is None or args.end_page is None:
			print('[Error] Both start page(-sp) and end page(-ep) numbers are required with "trns" options')
		# Set up parameters
		if args.start_page is not None:
		  startPage = args.start_page

		if args.end_page is not None:
		  # Sequence Numbers for this proeject - 1,3,5,7,10
	  	  endPage = args.end_page

		print(f'[INFO] Translating from {startPage}p to {endPage}p')

		print('[INFO] Get glossary data from Google sheet')
		glossary_dict, csv_content = get_glossary_from_sheet()
		
		dict_to_csv(glossary_dict, G.GT_GLOSSARY_FILE_NAME)
		glossary_cnt = len(glossary_dict)

		create_google_nmt_glossary_on_GCloud(csv_content, glossary_cnt)
		#glossary = get_google_nmt_glossaray()
		G.PAPAGO_GLOSSARY_ID = papago_glossary_upload()

		for pdf_page in range(int(startPage), int(endPage)):
			print(f'[INFO] Translating {pdf_page}p start!')
			txt_page = str(pdf_page) + '.txt'
			translator(txt_page, pdf_page)

	# Handle Option [2] - create translation goole doc
	elif args.command == "docs":
		if args.chapter is None:
			print('[Error] Chapter number(-ch) are required with "docs" options')

		if args.chapter is not None:
			print(f'[INFO] Create google docs for chapter {args.chapter}')

		docs_id, sheets_id = get_google_docs_sheets_id(args.chapter)
		print(f'\t[INFO] Spreadsheets ID: {sheets_id}')
		print(f'\t[INFO] Document ID: {docs_id}')

		make_translation_documents(sheets_id, docs_id)
		print ('[INFO] Complete!')