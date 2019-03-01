# To make print working for Python2/3
from __future__ import print_function

# Use six to import urllib so it is working for Python2/3
from six.moves import urllib
# If you don't want to use six, comment out the line above
# and use the line below instead (for Python3 only).
#import urllib.request, urllib.parse, urllib.error

import time
import pandas as pd
import xlsxwriter
import xlrd

# Build the cookie handler
cookier = urllib.request.HTTPCookieProcessor()
opener = urllib.request.build_opener(cookier)
urllib.request.install_opener(opener)

# Cookie and corresponding crumb
_cookie = None
_crumb = None

# Headers to fake a user agent
_headers={
	'User-Agent': 'Mozilla/5.0 (X11; U; Linux i686) Gecko/20071127 Firefox/2.0.0.11'
}

def _get_cookie_crumb():
	'''
	This function perform a query and extract the matching cookie and crumb.
	'''

	# Perform a Yahoo financial lookup on SP500
	req = urllib.request.Request('https://finance.yahoo.com/quote/^GSPC', headers=_headers)
	f = urllib.request.urlopen(req)
	alines = f.read().decode('utf-8')

	# Extract the crumb from the response
	global _crumb
	cs = alines.find('CrumbStore')
	cr = alines.find('crumb', cs + 10)
	cl = alines.find(':', cr + 5)
	q1 = alines.find('"', cl + 1)
	q2 = alines.find('"', q1 + 1)
	crumb = alines[q1 + 1:q2]
	_crumb = crumb

	# Extract the cookie from cookiejar
	global cookier, _cookie
	for c in cookier.cookiejar:
		if c.domain != '.yahoo.com':
			continue
		if c.name != 'B':
			continue
		_cookie = c.value

	# Print the cookie and crumb
	#print('Cookie:', _cookie)
	#print('Crumb:', _crumb)

def load_yahoo_quote(ticker, beginYr, beginMonth, begindate, endYr, endMonth, enddate, info = 'quote', format_output = 'list'):
	'''
	This function load the corresponding history/divident/split from Yahoo.
	'''
	# Check to make sure that the cookie and crumb has been loaded

	global _cookie, _crumb
	if _cookie == None or _crumb == None:
		_get_cookie_crumb()

	# Prepare the parameters and the URL
	ta = (beginYr, beginMonth, begindate, 4, 0, 0, 0, 0, 0)
	tc = (endYr, endMonth, enddate, 18, 0, 0, 0, 0, 0)
	tb = time.mktime(ta)
	te = time.mktime(tc)

	param = dict()
	param['period1'] = int(tb)
	param['period2'] = int(te)
	param['interval'] = '1d'
	if info == 'quote':
		param['events'] = 'history'
	elif info == 'dividend':
		param['events'] = 'div'
	elif info == 'split':
		param['events'] = 'split'
	param['crumb'] = _crumb
	params = urllib.parse.urlencode(param)
	url = 'https://query1.finance.yahoo.com/v7/finance/download/{}?{}'.format(ticker, params)
	#print(url)
	req = urllib.request.Request(url, headers=_headers)

	# Perform the query
	# There is no need to enter the cookie here, as it is automatically handled by opener
	f = urllib.request.urlopen(req)
	alines = f.read().decode('utf-8')
	#print(alines)
	if format_output == 'list':
		return alines.split('\n')

	if format_output == 'dataframe':
		nested_alines = [line.split(',') for line in alines.split('\n')[1:]]
		cols = alines.split('\n')[0].split(',')
		adf = pd.DataFrame.from_records(nested_alines[:-1], columns=cols)
		adf.to_excel(ticker + '.xlsx')
		return adf

portfolio = ['aapl','nflx','amzn']
tickerCount = 1
colPr = 5
tempCount = 2
wb2 = xlsxwriter.Workbook('port.xlsx')
s1 = wb2.add_worksheet()
for tick in portfolio:
	#df = load_yahoo_quote(tick, 2018, 1, 1, 2019, 1, 1, info='quote', format_output='dataframe')
	wb = xlrd.open_workbook(tick + '.xlsx') 
	sheet = wb.sheet_by_index(0)
	rowEnd = sheet.nrows - 1
	
	if tickerCount == 1:
		for counter in range(1,rowEnd):
			copyVal = sheet.cell_value(counter,tickerCount)
			s1.write(counter,tickerCount,copyVal)
			copyVal2 = sheet.cell_value(counter,colPr)
			s1.write(counter,tempCount,copyVal2)
			s1.write(0,tempCount,tick)
	else:
		for counter1 in range(1,rowEnd):
			copyVal3 = sheet.cell_value(counter,colPr)
			s1.write(counter,tempCount,copyVal3)
			s1.write(0,tempCount,tick)
	tempCount+=1
wb2.close()
