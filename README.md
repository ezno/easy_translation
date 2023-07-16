## 1. Overview

Translate each sentence of the page you want to translate to Google translation API v3, DeepL API, Naver Papago API, and Kakao translator. And organized the results in a Google Spreadsheet

* Google translation v3 beta API with glossary
* Papago tranlsation API with glossary (paid version)
* DeepL translation API (beta version)
* Kakao tranlsation

## 2. Installation

```
$ pip install -r requirements.txt
```

important python libraries

* Textblob (https://textblob.readthedocs.io/en/dev/)
```
$ pip install textblob
$ python -m textblob.download_corpora
```

* Google Translation Python client library (https://cloud.google.com/translate/docs/reference/libraries/v2/python)
```
$ pip install google-cloud-translate
```

* Python Client for Google Cloud Storage (https://cloud.google.com/python/docs/reference/storage/latest)
```
$ pip install google-cloud-storage
```

* Google cloud local setting
```
gcloud auth application-default login
gcloud auth application-default set-quota-project eznopub-142102
```

## 3. How to use

### Common

1. Clone this repo on your client
2. Download `config.ini`, `client_secret.json`, `token.json` files on the client 

### Translation

#### Step1 - Google spreadsheet

1.  Duplicate prior Google spreadsheet file 
2.  Copy new spreadsheet's Id value from URL of new spread sheet

```
https://docs.google.com/spreadsheets/d/{SPREAD_SHEET_ID}/edit#gid=0
```

3.  Go to Dashboard Google Spread sheet
4.  Paste spreadsheet's Id on your part of `sheetID` column

### Step 2 - run
1. Parepare `{page number.txt}` on the directory
2. Run autoNMTTranslator.py code

* -sp / --startPage : translation start page
* -ep / --endPage : translation end page

- example, translation 74.txt, 75.txt files

```
python3 autoNMTTranslator.py trns -sp 74 -ep 76
```

### Creating trnaslation Document

#### Step1 - Google spreadsheet

1. Create new a Google docs
2. Copy new document's Id value from URL of new document

```
https://docs.google.com/document/d/{DOCUMENT_ID}/edit 
```

3. Go to Dashboard Google Spread sheet
4. Paste document's Id on your part of `docID` column

#### Step 2 - run
1. Run autoNMTTranslator.py code

* -ch  : specific chapter number to translate

- example, translation chapter 3

```
python3 autoNMTTranslator.py docs -ch 3
```



