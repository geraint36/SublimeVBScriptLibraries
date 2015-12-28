# sssumes that the files that this is being run on contain valid VB syntax

import sublime_plugin
import sublime
import re
import os

VBSCRIPT_ALLOW_VAR_NAME_REGEX = '\\b[a-zA-Z]{1}[a-zA-Z0-9_]{,254}\\b'
VBSCRIPT_LIBRARY_PARENT_FOLDER = '\\testlibrary\\'

class ImportedClassesMethods(sublime_plugin.EventListener):

	def on_query_completions(self, view, prefix, locations):
		words = []
		matches = []

		filePath = view.file_name()
		# to get the preceeding word (which could be a varaible storing a library)
		variableName, charFollowing = getViewWordBeforeCursorsWord(view)
		# gets a dictionary of all the imports used in the currently opened file 
		# looks at the saved version no sublime's opened one
		imports = extractImports(filePath)

		# gets the path of the current \TestLibrary\ directory
		libraryDirPath = filePath[:filePath.lower().find(VBSCRIPT_LIBRARY_PARENT_FOLDER) \
			+ len(VBSCRIPT_LIBRARY_PARENT_FOLDER)]

		# if a '.' character follows the keyword try to display the methods
		if charFollowing == '.':
			# goes through every file in the library directory
			for directory, subdirectories, files in os.walk(libraryDirPath):
				for file in files:
					# full path of the file
					path = os.path.join(directory, file)
					pos = path.lower().find(VBSCRIPT_LIBRARY_PARENT_FOLDER) \
						+ len(VBSCRIPT_LIBRARY_PARENT_FOLDER)
					relativePath = path[pos:]
					fileExtension = os.path.splitext(path)[1].lower()

					# ignore files that are not '.vbs' or '.qfl'
					if (fileExtension != '.vbs') and (fileExtension != '.qfl'):
						continue
					# ignore words that are not imported libraries
					if not (variableName in imports.keys()):
						continue
					# ignore files that don't match the library that is stored in the variable
					elif formatImportPath(imports[variableName]) != formatImportPath(relativePath):
						continue

					properties = extractProperties(path)
					for prop in properties:
						comment, scope, propertyName = prop
						# ignores the private properties
						if scope == 'private':
							continue

						trigger, contents = buildTriggerAndContents(comment, propertyName)
						matches.append(('$%s' % trigger, contents))

					# extract all the methods of the library and put them into the matches array 
					# elements of the matches array are tuples with a tiggers and the actualy contents
					methods = extractMethods(path)

					for method in methods:
						comment, scope, methodParamsStr = method
						# ignores the private methods
						if scope == 'private':
							continue

						trigger, contents = buildTriggerAndContents(comment, methodParamsStr)
						matches.append((trigger, contents))
					
					break

		# if an empty list is returned from this method then the standard sublime suggestions will be used
		# this means that after any keyword that stores a library none of the standard suggestions will be 
		# available but everywhere else it'll just display the standard auto-complete options
		return matches

# gets the word preceeding the word that the cursor is curently at
# returns the tuple (preceedingWord, charFollowing)
def getViewWordBeforeCursorsWord(view):
	# [0] is used because of the posiblility of multiple cursors
	region = view.sel()[0]
	# gets the start positions of the word that the cursor is currently at
	wordStart = view.word(region).begin()
	# get the word before the current one (possible a variable storing a library)
	previousWordRegion = view.word( sublime.Region(wordStart-1, wordStart-1) )
	# gets the string for the word from its region
	preceedingWord = view.substr( previousWordRegion ).lower()
	# gets the character after the word
	charFollowing = view.substr( sublime.Region(previousWordRegion.end(), previousWordRegion.end()+1) )
	return preceedingWord, charFollowing

# extracts the methods of a class from a library file
def extractMethods(path):
	content = returnClassString(path)

	# finds all the strings that represent functions or subs and their parameters
	# will return a iterable collection of MatchObjects (in this case) each with 7 
	# groups. The first will contain the comments above the method the third the 
	# scope of the method the forth the type of the method (e.g. Sub or Function)
	# the sixth will be the methods name and the seventh will be the paramater list 
	# or None if there aren't any
	matches = re.finditer('((\\s*\'.*\\n)*)^\\s*(\\bPrivate\\b|\\bPublic\\b|)\\s*' \
		+ '(\\bFunction\\b|\\bSub\\b|\\bProperty\\b\\s*(\\bLet\\b|\\bSet\\b|\\bGet\\b))\\s*(' \
		+ VBSCRIPT_ALLOW_VAR_NAME_REGEX + ')\\s*(\\([a-zA-Z0-9\\,\\s]*\\))?$', content, \
		re.IGNORECASE | re.MULTILINE)

	# builds an array of the function and sub strings
	methods = []
	for match in matches:
		comment, scope, methodType, methodParamsStr = formatMethodStr(match)
		methods.append((comment, str(scope), str(methodParamsStr)))

	return methods

def extractProperties(path):
	content = returnClassString(path)

	# finds all the strings that represent functions or subs and their parameters
	# will return a iterable collection of MatchObjects (in this case) each with 7 
	# groups. The first will contain the comments above the method the third the 
	# scope of the method the forth the type of the method (e.g. Sub or Function)
	# the sixth will be the methods name and the seventh will be the paramater list 
	# or None if there aren't any
	matches = re.finditer('((\\s*\'.*\\n)*)^\\s*(\\bPrivate\\b|\\bPublic\\b)\\s*(' \
		+ VBSCRIPT_ALLOW_VAR_NAME_REGEX + ')\\s*$', content, re.IGNORECASE | re.MULTILINE)

	# builds an array of the function and sub strings
	properties = []

	for match in matches:
		comment, scope, propertyName = formatPropertyStr(match)
		properties.append((comment, str(scope), str(propertyName)))

	return properties

# returns a sub string for the file that corresponds to the the class in the library 
# file allowing for vbScript removing line continuation characters and putting the lines 
# on one line instead
def returnClassString(path):
	content = ''
	inClass = False
	with open(path) as f:
		for line in f:
			writeLine = line.strip()

			if writeLine[:5].lower() == 'class':
				inClass = True
			elif inClass and (writeLine[:3].lower() == 'end') and (writeLine[-5:].lower() == 'class'):
				inClass = False

			# if not inside a class the do nothing and continue to the next line
			if not inClass:
				continue

			if len(writeLine) == 0:
				continue
			elif writeLine[-1] == '_':
				content += writeLine[:-1]
				continue

			content += writeLine + '\n'
	return content

# extracts all the libraries that are imported by the specified one
def extractImports(path):
	content = returnFileString(path)

	# finds all the strings that represent relative paths of the libraries that are 
	# imported will return a iterable collection of MatchObjects (in this case) each 
	# with 2 groups. The first will contain variable name that the class is stored in the 
	# second will contain the relative path of the library that the class comes from
	matches = re.finditer('\\bSet\\b\\s*(' + VBSCRIPT_ALLOW_VAR_NAME_REGEX + \
		')\\s*=\\s*\\bImport\\s*\\(\\s*"([a-zA-Z0-9\\.\\\\/]+)\\s*"\\s*\\)', content, re.IGNORECASE)

	# builds an dictionary of the variables and their relative paths
	imports = {}
	for match in matches:
		imports[match.group(1).lower()] = formatImportPath(match.group(2))

	return imports

# returns a string for a library file allowing for vbScript removing line continuation characters 
# and putting the lines on one line instead
def returnFileString(path):
	content = ''
	with open(path) as f:
		for line in f:
			writeLine = line.strip()

			if len(writeLine) == 0:
				continue
			elif writeLine[-1] == '_':
				content += writeLine[:-1]
				continue

			content += writeLine + '\n'
	return content

# removes newline character from comment and just returns the decription of the function
# (for fancier comments of the form in the  /lib/Methods.qfl library)
def getCommentDescription(inputComment):
	if (not isinstance(inputComment, str)):
		return None

	commentLines = inputComment.split('\n')
	output = ''
	METHOD_DESC_KEYWORD = 'description'.lower()

	# difference between re.search and re.match;
	#    re.search - looks in whole string for first match
	#    re.match - looks for a match that begins at the start of the string

	# case for comments like thoose in the Methods.qfl library (fancier)
	if (commentLines[0][:1] == "#") and \
		(re.search('(\\b' + METHOD_DESC_KEYWORD + '\\b)\\s*:', inputComment, re.IGNORECASE) != None):

		# boolean variable used to see if currently inside description part of the comment
		pastDescLine = False
		for line in commentLines:
			keyWordMatchObj = re.match('\\b([a-z]*)\\b\\s*:', line, re.IGNORECASE)
			# check to see if the line if of the form 'keyword :'
			if keyWordMatchObj != None:
				# gets the keyword from the regex match
				keyWord = keyWordMatchObj.group(1).lower()
				if keyWord == METHOD_DESC_KEYWORD:
					pastDescLine = True
					# adds a line to the output adding in a space if required (ignoring the 
					# 'description:' part)
					output = addLineAutoCompleteComment(output, \
						line[len(METHOD_DESC_KEYWORD):].lstrip(' :'))
				else:
					# stops building the comment if the description has already been found
					# and currently on a new keyword
					if pastDescLine:
						break
			# if the line is of a normal form and currently in the description section
			# then add that line to the comment
			elif pastDescLine:
				# adds a line to the output adding in a space if required
				output = addLineAutoCompleteComment(output, line)
	# default case
	else:
		for line in commentLines:
			# adds a line to the output adding in a space if required
			output = addLineAutoCompleteComment(output, line)
	# returns the built comment removing any white space to the right
	return output.rstrip()

# adds a line to the output adding in a space if required (and ignoring some lines)
def addLineAutoCompleteComment(comment, inputLine):
	# remove unwanted "'" characters from right of string (will cause occational
	# errors where something is meant the be quoted)
	# used as lots of people start and end comments with the "'" character as opposed
	# to just starting them with it
	line = inputLine.strip(" '")
	# ignore enpty lines
	if len(line) == 0:
		pass
	# ignore spacer lines of '#' characters (maybe achange so ignores lines that
	# are made of just one character)
	elif line == '#' * len(line):
		pass
	else:
		comment += line + ' '
	return comment

# formats the comment returned by the regular expression
def formatComment(inputComment):
	if isinstance(inputComment, str):
		if inputComment == '':
			return None
		else:
			comment = ''
			lines = inputComment.strip(' \n').split('\n')
			for line in lines:
				comment += line.lstrip(" '\t") + '\n'
			return comment
	else:
		return None
	output = ''

# returns a standard format for the library relative paths
def formatImportPath(str):
	str = str.lower()
	
	if str.find('.') != -1:
		str = str[:str.find('.')]

	str = str.replace('/', '\\')	
	return str.strip('\\')

def formatPropertyStr(match):
	COMMENT_POS = 1
	PROPERTY_SCOPE_POS = 3
	PROPERTY_NAME_POS = 4

	comment, scope = getCommentAndScope(match, COMMENT_POS, PROPERTY_SCOPE_POS)
	propertyName = match.group(PROPERTY_NAME_POS)
	return comment, scope, propertyName

# removes white spaces and the words Function, Sub, ByVal and ByRef from the function definition string
def formatMethodStr(match):
	pattern = re.compile("(\\s+|\\bbyval\\b|\\bbyref\\b)", re.IGNORECASE)
	COMMENT_POS = 1
	METHOD_SCOPE_POS = 3
	METHOD_TYPE_POS = 4
	METHOD_NAME_POS = 6
	PARAMATERS_POS = 7

	comment, scope = getCommentAndScope(match, COMMENT_POS, METHOD_SCOPE_POS)

	# case where there are no parmaters
	if match.group(PARAMATERS_POS) == None:
		methodParamsStr = match.group(METHOD_NAME_POS) + '()'
	# case when there are paramaters
	else:
		methodParamsStr = match.group(METHOD_NAME_POS) + match.group(PARAMATERS_POS)

	# removes white spaces and unwanted keywords
	methodParamsStr = pattern.sub('', methodParamsStr)

	# allows for different formatting of subs and functions
	methodType = match.group(METHOD_TYPE_POS).lower()
	if methodType in ('sub', 'property let', 'property set'):
		methodParamsStr = methodParamsStr.replace('()', '').replace('(', ' ').replace(')', '')
	return comment, scope, methodType, methodParamsStr

def getCommentAndScope(match, comment_pos, scope_pos):
	# gets the comment immedietly above the method name
	comment = match.group(comment_pos)
	if comment == None:
		comment = ''

	# removes whitespace and "'" character from left of comment lines
	comment = formatComment(comment)
	comment = getCommentDescription(comment)

	# gets the scope of the method
	if match.group(scope_pos) == None:
		scope = 'public'
	elif match.group(scope_pos) == '':
		scope = 'public'
	else:
		scope = match.group(scope_pos).lower()
	return comment, scope

def buildTriggerAndContents(comment, keywordStr):
	# options for both comments and no comments
	if comment == None:
		trigger = "%s" % (keywordStr,)
	else:
		trigger = "%s\t'%s" % (keywordStr, comment)

	# replace done as '$' is a special character that seems to stop the auto-complete
	contents = keywordStr.replace('$', '\\$')
	return trigger, contents