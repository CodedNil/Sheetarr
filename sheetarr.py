#!/usr/bin/python3

import gspread
from oauth2client.service_account import ServiceAccountCredentials

import json
import time
import os
import string
import pyarr
from colorama import Fore, Style
import requests
import datetime

#################################
############# TODOS #############
#################################

# Automate the remover script to pull from sheet, if text is coloured anything but default black, proceed with removal, make sure to reset formatting after this

# Local log.txt file which logs every debug change, with append write mode

# If sonarr, radarr or spreadsheet data file is missing, prompt user to provide the information
# Make discord optional and add readme info about it
# Run arguments to change information stored
# --help argument to display help

# Setup command which creates the spreadsheet with correct blank format

# Make the sheets non destructive, if rows added or deleted etc, always readjust them
# Reset row/col sizes
# Set resolution drop down choices
# Set top bar text and text sizes
# Set the alternating background colour

# Readme improvements


#################################
######### CONFIGURATION #########
#################################

# Manage auto run at 20 minute intervals with
# crontab -e
# */20 * * * * /DIRECTORY/run_sheetsmanager.sh

shouldPullResolution = False # Should the resolution be pulled from the sites or pushed to them


#################################
######### AUTHORISATION #########
#################################

# Read from credentials file
credentials = json.loads(open('credentials.json').read())

# Authorise google sheets
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(credentials['sheet']['keyfile'], scope)
client = gspread.authorize(creds)
gspreadsheet = client.open(credentials['sheet']['sheetname'])

# Setup pyarr api with auth
sonarrapi = pyarr.SonarrAPI(credentials['sonarr']['url'], credentials['sonarr']['api'])
sonarrapi.basic_auth(credentials['sonarr']['authuser'], credentials['sonarr']['authpass'])

radarrapi = pyarr.RadarrAPI(credentials['radarr']['url'], credentials['radarr']['api'])
radarrapi.basic_auth(credentials['sonarr']['authuser'], credentials['sonarr']['authpass'])

# Discord
discordWebhook = credentials['discord']


#################################
############# UTILS #############
#################################

def n2a(n,b=string.ascii_uppercase):
	"""Convert a number to a column letter, e.g. 1 -> A, 5 -> E, 27 -> AA"""
	d, m = divmod(n,len(b))
	return n2a(d-1,b)+b[m] if d else b[m]

def lstd(list, key, default=''):
	"""Get list data if exists or return default"""
	return list[key] if key in list else default

def TitleMatch(name, title, year):
	"""Check if title matches name, with or without the release year"""
	return title.lower() == name.lower() or (title + ' ' + str(year)).lower() == name.lower()

def sizeof_fmt(num, suffix="B"):
	""""Return the human readable size of a file from bytes, e.g. 1024 -> 1KB"""
	for unit in ["", "Ki", "Mi", "Gi", "Ti", "Pi", "Ei", "Zi"]:
		if abs(num) < 1024:
			return f"{num:3.1f}{unit}{suffix}"
		num /= 1024
	return f"{num:.1f}Yi{suffix}"

def fuzzyMatchList(listA, listB, fuzzyness=0.05):
	"""Match the items in listA to the items in listB with a fuzzy margin"""
	for index, itemA in enumerate(listA):
		itemB = listB[index]
		if abs(itemA - itemB) > fuzzyness:
			return False
	return True

# Load cache
cache = {'discord': [], 'quota': []}
def SaveCache():
	"""Saves the cache to a file"""
	with open('cache.json', 'w') as json_file:
		json.dump(cache, json_file)

if os.path.isfile('cache.json'):
	with open('cache.json') as json_file:
		cache = json.load(json_file)

	# Remove old entries from quota cache
	quotaList = [key for key in cache['quota'] if time.time() - key < 60]
	if cache['quota'] != quotaList:
		cache['quota'] = quotaList
		SaveCache()

def PostDiscord(colour, message):
	"""Posts a discord message"""
	# Send message once per day max
	debugmsg = datetime.date.today().strftime('%Y-%m-%d') + ' - ' + message
	print(colour + debugmsg + Style.RESET_ALL)
	if debugmsg not in cache['discord']:
		requests.post(discordWebhook, headers={"Content-Type": "application/json"}, data=json.dumps({'content': debugmsg}))
		cache['discord'].append(debugmsg)
		SaveCache()

# Profile id to resolution
qualityProfiles = {2: 'SD', 3: '720p', 4: '1080p', 5: '2160p', 6: '720p/1080p', 7: 'Any'}
qualityFromProfile = {'SD': 2, '720p': 3, '1080p': 4, '2160p': 5, '720p/1080p': 6, 'Any': 7}
qualityResolutions = {'SD': 480, '720p': 720, '1080p': 1080, '2160p': 2160, '720p/1080p': 900, 'Any': 0}


#################################
######## GRAB SITES DATA ########
#################################

sonarrList = sonarrapi.get_series()
radarrList = radarrapi.get_movie()

# Search against sonarr/radarr
def SearchAgainstSite(name, wantedres, isseries):
	"""Returns the series/movie data if found, else adds the series/movie to sonarr/radarr
	name, the name of the series/movie to search for
	wantedres, the resolution to search for if adding
	isseries, a boolean for if the search is for a series or a movie
	
	Returns the status of the search
	'found', item - the series/movie data object if the series/movie was found existing
	'adding' - the series/movie wasn't found but was matched correctly and was added to sonarr/radarr
	'failedmatch', matches - the series/movie did not match exactly but potential matches were found
	'failed' - the series/movie could not be matched
	"""

	# Find if item is already on sonarr/radarr
	found = False
	for item in isseries and sonarrList or radarrList:
		if TitleMatch(name, item['title'], item['year']):
			found = True
			return 'found', item

	# If not found, do a search to find closest match
	if not found:
		response = None
		if isseries:
			response = sonarrapi.lookup_series(name)
		else:
			response = radarrapi.lookup_movie(name)

		# Check if matches are found
		if len(response) > 0:
			titleyear = response[0]['title'] + ' -' + str(response[0]['year'])
			if TitleMatch(name, response[0]['title'], response[0]['year']):
				# Search result found and matches
				if isseries:
					# Add result to sonarr
					wantedquality = qualityFromProfile[wantedres]
					sonarrapi.add_series(response[0]['tvdbId'], wantedquality, '/tv', season_folder=True, monitored=True, search_for_missing_episodes=True)
					# Send debug message to discord
					PostDiscord(Fore.MAGENTA, 'Adding to Sonarr: ' + name)
				else:
					# Add result to radarr
					wantedquality = qualityFromProfile[wantedres]
					radarrapi.add_movie(response[0]['tmdbId'], wantedquality, '/movies', monitored=True, search_for_movie=True, tmdb=True)
					# Send debug message to discord
					PostDiscord(Fore.MAGENTA, 'Adding to Radarr: ' + name)

				return 'adding', None
			else:
				# Search result found but not exact match
				print(Fore.RED + 'FAILED: ', name, ' '*(50 - len(titleyear)), titleyear + Style.RESET_ALL)

				# Add first few results to match data
				matches = []
				for i in range(0, 2):
					if len(response) > i:
						matches.append(response[i]['title'])
				return 'failedmatch', matches
		else:
			# Search result not found
			print(Fore.RED + 'FAILED ' + name + Style.RESET_ALL)
			return 'failed', None


#################################
######## WRITE TO SHEETS ########
#################################

def CalculateQuota():
	"""Calculates the number of writes made in the last minute"""
	currentquota = 0
	for quota in cache['quota']:
		if time.time() - quota < 60:
			currentquota += 1
	return currentquota

def WriteSheet(gsheet, title, func, cell, *args):
	"""Writes to the sheet with the given function and arguments"""
	value = json.dumps([item for item in args]).replace('\n', ' ')
	print(Fore.YELLOW + title + ' ' + func + ' ' + cell + ' : ' + value + Style.RESET_ALL)

	# Wait for quota to be below limit
	currentquota = CalculateQuota()
	if currentquota > 55:
		while CalculateQuota() > 50:
			time.sleep(5)

	# Add current time to quota list and write cache
	cache['quota'].append(time.time())
	SaveCache()
	
	# Run function with args
	if func == 'format':
		gsheet.format(cell, *args)
	elif func == 'update':
		gsheet.update(cell, *args)
	elif func == 'update_acell':
		gsheet.update_acell(cell, *args)
	elif func == 'insert_note':
		gsheet.insert_note(cell, *args)


#################################
####### PROCESS SHEET DATA ######
#################################

def ProcessSheetMedia(gsheet, title, isSeries, cellData):
	"""
	Processes the sheets data and push any necessary changes to hyperlinks, notes, text color etc with data from sonarr/radarr

	title, the sheets title
	isSeries, a boolean for if the search is for a series or a movie
	cellData, the data from the cell in the sheet, a list of the 3 cells, text, resolution and status, each contains a dictionary with, cell, text, hyperlink, note, textColor
	"""
	mediaTitle = cellData[0]['text']
	wantedResolution = cellData[1]['text']

	# Variables for what the sheets data should be
	wantedMainHyperlink = ''
	wantedMainNote = ''
	wantedMainTextColor = [0, 0, 0]

	wantedResolutionText = '1080p'

	wantedStatusText = ''
	wantedStatusNote = ''
	wantedStatusTextColor = [0, 0, 0]

	fileSize = 0 # Size of media file(s)
	hasFile = 0 # 0 if no file, 1 if file exists

	# Search the sites for the media
	if mediaTitle != '':
		result, item = SearchAgainstSite(mediaTitle, wantedResolution, isSeries)

		duped = False # Todo check if media is a dupe
		dupedList = [] # Todo check if media is a dupe

		# Purple if duped, red if failed, orange if not on or black
		wantedMainTextColor = duped and [0.5, 0.05, 0.75]\
			or result == 'found' and [0, 0, 0]\
			or result == 'adding' and [0.5, 0.5, 0]\
			or [1, 0, 0]

		# Get required notes
		if result == 'failedmatch':
			# Couldn't match message to series alternative names
			wantedMainNote = "Couldn't match, did you mean\n" + '\n'.join(item)
		elif result == 'adding':
			# Adding message
			wantedMainNote = 'Not on {site} yet, will be added automatically soon'.format(site = isSeries and 'Sonarr' or 'Radarr')
		elif result == 'found':
			# Get media status and year
			wantedMainNote = str(item['year']) + '\n' + item['status'].capitalize()
		if duped:
			# If media is duped, add that to the note
			wantedMainNote += '\nThis exists on multiple sheets: ' + ', '.join(dupedList)

		# Get required status notes
		if result == 'found':
			if isSeries:
				# Get series season episode count and size breakdown
				seasonnotes = []
				for season in item['seasons']:
					seasonnumber = str(season['seasonNumber'])
					seasonepisodes = str(season['statistics']['episodeFileCount']) + '/' + str(season['statistics']['totalEpisodeCount'])
					seasonsize = sizeof_fmt(season['statistics']['sizeOnDisk'])
					seasonnotes.append('Season ' + seasonnumber + ':   ' + seasonepisodes + ' '*(8 - (len(seasonepisodes) - 1)) + seasonsize)
				wantedStatusNote = '\n'.join(seasonnotes)

		# Get required resolution value
		if result == 'found':
			wantedResolutionText = qualityProfiles[item['qualityProfileId']] # The resolution of the media on the site
			if wantedResolutionText != cellData[1]['text'] and not shouldPullResolution:
				# Adjust resolution on sonarr or radarr
				wantedQuality = qualityFromProfile[cellData[1]['text']]
				params = item
				params['profileId'] = wantedQuality
				params['qualityProfileId'] = wantedQuality
				if isSeries:
					sonarrapi.upd_series(params)
				else:
					radarrapi.upd_movie(params)
				# Send debug message to discord
				PostDiscord(Fore.MAGENTA, 'Adjusting {site} resolution: {media} from {old} to {new}'.format(site = isSeries and 'Sonarr' or 'Radarr', media = mediaTitle, old = cellData[1]['text'], new = wantedResolutionText))
				
		# Get required hyperlink value
		if result == 'found':
			wantedMainHyperlink = (isSeries and 'http://sonarr.ratatoskr.uk/series/' or 'http://radarr.ratatoskr.uk/movie/') + item['titleSlug']
		
		# Get required status text and color
		if result == 'found':
			if isSeries:
				# Series status data
				fileCount = item['episodeFileCount']
				episodeCount = item['episodeCount']
				wantedStatusText = sizeof_fmt(item['sizeOnDisk']) + ' ' + str(fileCount) + '/' + str(episodeCount)

				hasFile = episodeCount > 0 and fileCount / episodeCount or 0

				wantedStatusTextColor = episodeCount == 0 and [1, 0, 0] or fileCount == episodeCount and [0, 0.75, 0] or fileCount < episodeCount and [0.3, 0.3, 0] or [1, 0, 0]
			else:
				# Movie status data
				wantedStatusText = sizeof_fmt(item['sizeOnDisk'])

				resDifference = 0 # 0 if the same, 1 if above, -1 if below
				hasFile = item['hasFile'] and 1 or 0 # For return
				actualResNum = 0
				if item['hasFile']:
					radarrres = qualityProfiles[item['qualityProfileId']]
					actualResNum = item['movieFile']['quality']['quality']['resolution']
					targetresnum = qualityResolutions[radarrres]
					if actualResNum == targetresnum:
						resDifference = 0
					elif actualResNum > targetresnum:
						resDifference = 1
					else:
						resDifference = -1
				# Green if has file and resolution matches, purple if resolution is too high, orange if resolution is too low, red if no file, red if file is missing
				wantedStatusTextColor = item['hasFile'] and (resDifference == 0 and [0, 0.75, 0] or resDifference == 1 and [0.5, 0.05, 0.75] or resDifference == -1 and [0.5,0.5,0]) or [1, 0, 0]

				# Get required status notes
				fileRes = str(actualResNum) + 'p'
				wantedStatusNote = resDifference == 1 and 'Resolution of file is too high, file is ' + fileRes or resDifference == -1 and 'Resolution of file is too low, file is ' + fileRes or item['hasFile'] and 'Has file at correct resolution' or 'Missing file'

		# Add total size variable
		if result == 'found':
			fileSize += item['sizeOnDisk']

	# Update the sheets data where necessary
	if cellData[0]['note'] != wantedMainNote:
		WriteSheet(gsheet, title, 'insert_note', cellData[0]['cell'], wantedMainNote)
	if cellData[0]['hyperlink'] != wantedMainHyperlink:
		WriteSheet(gsheet, title, 'update_acell', cellData[0]['cell'], '=HYPERLINK("' + wantedMainHyperlink + '", "' + mediaTitle + '")')
	if not fuzzyMatchList(cellData[0]['textColor'], wantedMainTextColor):
		WriteSheet(gsheet, title, 'format', cellData[0]['cell'], {'textFormat': {'bold': True, 'foregroundColor': {'red': wantedMainTextColor[0], 'green': wantedMainTextColor[1], 'blue': wantedMainTextColor[2]}, 'link': {'uri': wantedMainHyperlink}}})

	# Write 1080p to any blank cells, and pull to sheet if that option is enabled
	if mediaTitle == '' and cellData[1]['text'] != '1080p':
		WriteSheet(gsheet, title, 'update', cellData[1]['cell'], '1080p')
	else:
		if shouldPullResolution:
			if cellData[1]['text'] != wantedResolutionText:
				WriteSheet(gsheet, title, 'update', cellData[1]['cell'], wantedResolutionText)

	if cellData[2]['note'] != wantedStatusNote:
		WriteSheet(gsheet, title, 'insert_note', cellData[2]['cell'], wantedStatusNote)
	if cellData[2]['text'] != wantedStatusText:
		WriteSheet(gsheet, title, 'update', cellData[2]['cell'], wantedStatusText)
	if not fuzzyMatchList(cellData[2]['textColor'], wantedStatusTextColor):
		WriteSheet(gsheet, title, 'format', cellData[2]['cell'], {'textFormat': {'bold': True, 'foregroundColor': {'red': wantedStatusTextColor[0], 'green': wantedStatusTextColor[1], 'blue': wantedStatusTextColor[2]}}})
	
	return fileSize, hasFile, mediaTitle != ''


#################################
######## GRAB SHEET DATA ########
#################################

# List all worksheets
start = time.time()
params = {
	#"ranges": "'Dan'",
	"spreadsheetId": gspreadsheet.id,
	"includeGridData": True
}

# Get data from cache or sheets
sheetsDict = {}
sheetsData = gspreadsheet._spreadsheets_get(params)

# Data for series, movies count and sizes
sheetSeriesData = {}
sheetMoviesData = {}
# List for items that are not on the sheet, and items that are duplicates on the sheets
missingSheetSeries = []
duplicateSheetSeries = {}
missingSheetMovies = []
duplicateSheetMovies = {}

def GetMissingMedia():
	"""Get lists of series and movies on the site but not on the sheet"""
	for siteItem in sonarrList:
		found = False
		for sheetItem in sheetSeriesData:
			if TitleMatch(sheetItem, siteItem['title'], siteItem['year']):
				found = True
				break
		if not found:
			missingSheetSeries.append(siteItem['title'])

	for siteItem in radarrList:
		found = False
		for sheetItem in sheetMoviesData:
			if TitleMatch(sheetItem, siteItem['title'], siteItem['year']):
				found = True
				break
		if not found:
			missingSheetMovies.append(siteItem['title'])


for sheet in sheetsData['sheets']:
	gsheet = gspreadsheet.get_worksheet_by_id(sheet['properties']['sheetId'])
	# Get data about the sheet
	properties = sheet['properties']
	sheetTitle = properties['title']
	rowCount = properties['gridProperties']['rowCount']
	columnCount = properties['gridProperties']['columnCount']

	print(Fore.YELLOW + 'Sheet: ' + sheetTitle + Style.RESET_ALL)

	# Start dictionary
	sheetsDict[sheetTitle] = {
		'properties': {
			'rowCount': rowCount,
			'columnCount': columnCount
		},
		'rows': []
	}

	# Get the data from the sheet
	startRow = lstd(sheet['data'][0], 'startRow', 0)
	startColumn = lstd(sheet['data'][0], 'startColumn', 0)
	rowData = sheet['data'][0]['rowData']

	totalSeriesSize, totalSeriesFiles, totalSeriesCount = 0, 0, 0
	totalMoviesSize, totalMoviesFiles, totalMoviesCount = 0, 0, 0

	if sheetTitle == 'Info':
		# Remove all items that are only on one sheet
		duplicateSheetSeries = {i:j for i,j in duplicateSheetSeries.items() if len(j) > 1}
		duplicateSheetMovies = {i:j for i,j in duplicateSheetMovies.items() if len(j) > 1}
		GetMissingMedia()

	# Loop through all rows and get indexes
	for y, row in enumerate(rowData):
		rowIndex = startRow + y
		# Write row to dictionary
		sheetsDict[sheetTitle]['rows'].append({'cells': []})

		# Loop through all columns and get indexes
		for x, column in enumerate(row['values']):
			columnIndex = startColumn + x

			# Get cell data
			formattedValue = lstd(column, 'formattedValue')
			note = lstd(column, 'note')
			hyperlink = lstd(column, 'hyperlink')
			# Get text rgb
			textColor = [0, 0, 0]
			if 'userEnteredFormat' in column and 'textFormat' in column['userEnteredFormat'] and 'foregroundColorStyle' in column['userEnteredFormat']['textFormat']:
				textColorList = column['userEnteredFormat']['textFormat']['foregroundColorStyle']['rgbColor']
				textColor = [lstd(textColorList, 'red', 0), lstd(textColorList, 'green', 0), lstd(textColorList, 'blue', 0)]

			# Write data to dictionary
			sheetsDict[sheetTitle]['rows'][rowIndex]['cells'].append({
				'cell': n2a(columnIndex) + str(rowIndex + 1),
				'text': formattedValue,
				'note': note,
				'hyperlink': hyperlink,
				'textColor': textColor
			})

		if sheetTitle == 'Info':
			# Process the info sheet, requires being last processed
			if rowIndex > 3:
				rowData = sheetsDict[sheetTitle]['rows'][rowIndex]['cells']
				# Set missing series
				removeSeries = (len(missingSheetSeries) > rowIndex - 4) and missingSheetSeries[rowIndex - 4] or ''
				if rowData[0]['text'] != removeSeries:
					WriteSheet(gsheet, sheetTitle, 'update', rowData[0]['cell'], removeSeries)

				# Set dupe series text or nothing
				dupeSeriesName, dupeSeriesSheets, dupeSeriesText = '', [], ''
				if len(duplicateSheetSeries) > rowIndex - 4:
					dupeSeriesName, dupeSeriesSheets = list(duplicateSheetSeries.keys())[0], list(duplicateSheetSeries.values())[0]
					dupeSeriesText = dupeSeriesName + ' (' + ', '.join(dupeSeriesSheets) + ')'
				if rowData[1]['text'] != dupeSeriesText:
					WriteSheet(gsheet, sheetTitle, 'update', rowData[1]['cell'], dupeSeriesText)

				# Set missing movies
				removeMovies = (len(missingSheetMovies) > rowIndex - 4) and missingSheetMovies[rowIndex - 4] or ''
				if rowData[2]['text'] != removeMovies:
					WriteSheet(gsheet, sheetTitle, 'update', rowData[2]['cell'], removeMovies)

				# Set dupe movies text or nothing
				dupeMoviesName, dupeMoviesSheets, dupeMoviesText = '', [], ''
				if len(duplicateSheetMovies) > rowIndex - 4:
					dupeMoviesName, dupeMoviesSheets = list(duplicateSheetMovies.keys())[0], list(duplicateSheetMovies.values())[0]
					dupeMoviesText = dupeMoviesName + ' (' + ', '.join(dupeMoviesSheets) + ')'
				if rowData[3]['text'] != dupeMoviesText:
					WriteSheet(gsheet, sheetTitle, 'update', rowData[3]['cell'], dupeMoviesText)

		else:
			# For every non info sheet
			if rowIndex > 0:
				# Series cells
				seriesSize, seriesFile, seriesCount = ProcessSheetMedia(gsheet, sheetTitle, True, sheetsDict[sheetTitle]['rows'][rowIndex]['cells'][0:3])
				totalSeriesSize += seriesSize
				totalSeriesFiles += seriesFile
				totalSeriesCount += seriesCount
				seriesName = lstd(sheetsDict[sheetTitle]['rows'][rowIndex]['cells'][0], 'text')
				if seriesName != '':
					if seriesName in duplicateSheetSeries:
						duplicateSheetSeries[seriesName].append(sheetTitle)
					else:
						duplicateSheetSeries[seriesName] = [sheetTitle]
					sheetSeriesData[seriesName] = {
						'size': seriesSize,
						'files': seriesFile
					}

				# Movies cells
				moviesSize, moviesFile, moviesCount = ProcessSheetMedia(gsheet, sheetTitle, False, sheetsDict[sheetTitle]['rows'][rowIndex]['cells'][3:6])
				totalMoviesSize += moviesSize
				totalMoviesFiles += moviesFile
				totalMoviesCount += moviesCount
				moviesName = lstd(sheetsDict[sheetTitle]['rows'][rowIndex]['cells'][3], 'text')
				if moviesName != '':
					if moviesName in duplicateSheetMovies:
						duplicateSheetMovies[moviesName].append(sheetTitle)
					else:
						duplicateSheetMovies[moviesName] = [sheetTitle]
					sheetMoviesData[moviesName] = {
						'size': moviesSize,
						'files': moviesFile
					}

	if sheetTitle == 'Info':
		# Process the info sheet, requires being last processed
		seriesCount = len(sheetSeriesData)
		seriesSize, seriesFiles = 0, 0
		for name in sheetSeriesData:
			seriesSize += sheetSeriesData[name]['size']
			seriesFiles += sheetSeriesData[name]['files']

		moviesCount = len(sheetMoviesData)
		moviesSize, moviesFiles = 0, 0
		for name in sheetMoviesData:
			moviesSize += sheetMoviesData[name]['size']
			moviesFiles += sheetMoviesData[name]['files']

		wantedTextSeries = str(seriesCount) + ' - ' + str(round(seriesFiles/seriesCount*100)) + '%\n' + str(sizeof_fmt(seriesSize))
		wantedTextMovies = str(moviesCount) + ' - ' + str(round(moviesFiles/moviesCount*100)) + '%\n' + str(sizeof_fmt(moviesSize))

		if sheetsDict[sheetTitle]['rows'][1]['cells'][0]['text'] != wantedTextSeries:
			WriteSheet(gsheet, sheetTitle, 'update', 'A2:B2', wantedTextSeries)

		if sheetsDict[sheetTitle]['rows'][1]['cells'][2]['text'] != wantedTextMovies:
			WriteSheet(gsheet, sheetTitle, 'update', 'C2:D2', wantedTextMovies)
	else:
		# First row, push the total sizes to the sheet
		wantedTextSeries = totalSeriesCount == 0 and 'N/A' or str(totalSeriesCount) + ' - ' + str(round(totalSeriesFiles/totalSeriesCount*100)) + '%\n' + str(sizeof_fmt(totalSeriesSize))
		wantedTextMovies = totalMoviesCount == 0 and 'N/A' or str(totalMoviesCount) + ' - ' + str(round(totalMoviesFiles/totalMoviesCount*100)) + '%\n' + str(sizeof_fmt(totalMoviesSize))

		if sheetsDict[sheetTitle]['rows'][0]['cells'][2]['text'] != wantedTextSeries:
			WriteSheet(gsheet, sheetTitle, 'update', 'C1', wantedTextSeries)

		if sheetsDict[sheetTitle]['rows'][0]['cells'][5]['text'] != wantedTextMovies:
			WriteSheet(gsheet, sheetTitle, 'update', 'F1', wantedTextMovies)
