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
	#================================================================
	#parent = client.location_path(G.GT_PROJECT_ID, G.GT_LOCATION)
	parent = f'projects/{G.GT_GLOSSARY_PRJ_ID}/locations/{G.GT_LOCATION}'
	#================================================================

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
def translator(chapter, text_path, pageNum, glossary_config):
	# Record translation result on thie file in the client 
	TranText = os.path.join('outputs', 'trns_' + text_path)
	input_path = os.path.join('ch'+str(chapter), text_path)

	fpin = open(input_path, 'r+')
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
					textBlock['naverTrns'] = naver_neural_machine2_translate(sentence.raw, 'en', 'ko')
		
					# Google NeuralMachineTranlate results: English > Korean
					textBlock['googleTrns'] = google_neural_machine_translate_v3(sentence.raw, 'en', 'ko', glossary_config) + ' '

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
def google_neural_machine_translate_v3(source, from_lang, to, glossary_config):
	client = translate.TranslationServiceClient()

	parent = f'projects/{G.GT_GLOSSARY_PRJ_ID}/locations/{G.GT_LOCATION}'

	result = client.translate_text(
			request= {
				"parent": parent,
				"contents": [source],
				"source_language_code": from_lang,
				"target_language_code": to,
				"glossary_config": glossary_config,
			}
	)

	translated_txt = ''

	for glossary_translations in result.translations:
		results_text = glossary_translations.translated_text
		translated_txt = translated_txt + results_text

	if translated_txt == '':
		translated_txt = '\t[Error] Failed to translate by Google NMT v3'

	return post_string_changer(translated_txt)


# NAVER Papago translation - free version
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
	G.DEEPL_AUTH_KEY = config.get('API_KEYS', 'DEEPL_AUTH_KEY')
	G.PAPAGO_CLIENT_ID = config.get('API_KEYS', 'PAPAGO_CLIENT_ID')
	G.PAPAGO_CLIENT_SECRET = config.get('API_KEYS', 'PAPAGO_CLIENT_SECRET')

	print(G.SPREADSHEET_ID)
	print(G.PAPAGO_CLIENT_SECRET)

	# For argument parsing
	parser = argparse.ArgumentParser(
					description = "Translate Automator V7 (2023.6.25.)",
					formatter_class = ArgumentDefaultsHelpFormatter )

	# Parse command line arguments
	parser.add_argument("-ch", "--chapter", help="Translation chapter", required=True)
	parser.add_argument("-sp", "--startPage", help="Translation start page", required=True)
	parser.add_argument("-ep", "--endPage", help="Translation end page", required=True)
	args = vars(parser.parse_args())

	# Set up parameters
	if args["chapter"] is not None:
		chapter = args["chapter"]
	
	if args["startPage"] is not None:
		startPage = (args["startPage"])

	if args["endPage"] is not None:
		# Sequence Numbers for this proeject - 1,3,5,7,10
		endPage = args["endPage"]

	print(f'[INFO] Translating from {startPage}p to {endPage}p in Chapter{chapter}')

	print('[INFO] Get glossary data from Google sheet')
	glossary_dict, csv_content = get_glossary_from_sheet()
	glossary_cnt = len(glossary_dict)
	create_google_nmt_glossary_on_GCloud(csv_content, glossary_cnt)

	glossary = get_google_nmt_glossaray()
	glossary_config = {
			'glossary': glossary.name,
			'ignore_case': True
	}

	for pdf_page in range(int(startPage), int(endPage)):
		print('[INFO] Translating ' + str(pdf_page) + 'p Start!')
		txt_page = str(pdf_page) + '.txt'
		translator(chapter, txt_page, pdf_page, glossary_config)