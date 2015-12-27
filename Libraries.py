# TODO - currently also displays private methods of the classes
# Assumes that the files that this is being run on contain valid VB syntax

import sublime_plugin
import sublime
import re
import os

VBSCRIPT_ALLOW_VAR_NAME_REGEX = '\\b[a-zA-Z]{1}[a-zA-Z0-9_]{,254}\\b'

class ImportedClassesMethods(sublime_plugin.EventListener):

	def on_query_completions(self, view, prefix, locations):
		words = []
		matches = []

		filePath = view.file_name()
		# to get the preceeding word (which could be a varaible storing a library)
		# [0] is used because of the posiblility of multiple cursors
		region = view.sel()[0]
		# gets the start positions of the word that the cursor is currently at
		wordStart = view.word(region).begin()
		# get the word before the current one (possible a variable storing a library)
		previousWordRegion = view.word(sublime.Region(wordStart-1,wordStart-1))
		# gets the string for the word from its region
		variableName = view.substr(previousWordRegion).lower()

		# gets a dictionary of all the imports used in the currently opened file 
		# looks at the saved version no sublime's opened one
		imports = extractImports(filePath)

		# gets the path of the current \TestLibrary\ directory
		libraryDirPath = filePath[:filePath.lower().find('\\testlibrary\\') + 13]

		# goes through every file in the library directory
		for directory, subdirectories, files in os.walk(libraryDirPath):
			for file in files:
				# full path of the file
				path = os.path.join(directory, file)
				pos = path.lower().find('\\testlibrary\\') + 13
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

				# extract all the methods of the library and put them into the matches array 
				# elements of the matches array are tuples with a tiggers and the actualy contents
				methods = extractMethods(path)
				for method in methods:
					comment = formatComment(method[0])
					scope = method[1]
					methodParamsStr = method[2]

					if scope == 'private':
						continue

					if comment == None:
						trigger = "%s\t(%s)" % (methodParamsStr, relativePath)
					else:
						trigger = "%s\t(%s)" % (methodParamsStr, relativePath)

					contents = methodParamsStr
					matches.append((trigger, contents))

		# could exit the function and return the matches here
		return matches

# Assumes that the file this is being run on contains valid VB syntax
def extractMethods(path):
	content = returnClassString(path)

	# finds all the strings that represent functions or subs and their parameters
	# will return a iterable collection of MatchObjects (in this case) each with 3 
	# groups. The first will contain Sub or Function the second will be the methods
	# name and the third will be the paramater list or None if there aren't any
	# TODO - add bit so ignores private methods
	matches = re.finditer('((\\s*\'.*\\n)*)\\s*(\\bPrivate\\b|\\bPublic\\b|)\\s*(\\bFunction\\b|\\bSub\\b) (' + VBSCRIPT_ALLOW_VAR_NAME_REGEX \
		+ ')(\\([a-zA-Z0-9\\,\\s]*\\))?', content, re.IGNORECASE)

	# builds an array of the function and sub strings
	methods = []
	for match in matches:
		comment, scope, methodParamsStr = formatMethodStr(match)
		methods.append((str(comment), str(scope), str(methodParamsStr)))

	return methods

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

# removes white spaces and the words Function, Sub, ByVal and ByRef from the function definition string
def formatMethodStr(match):
	pattern = re.compile("(\\s|\\bbyval\\b|\\bbyref\\b|\\bFunction\\b|\\bSub\\b)", re.IGNORECASE)
	COMMENT_POS = 1
	METHOD_SCOPE_POS = 3
	METHOD_TYPE = 4
	METHOD_NAME_POS = 5
	PARAMATERS_POS = 6

	# gets the comment immedietly above the method name
	if match.group(COMMENT_POS) == None:
		comment = ''
	else:
		comment = match.group(COMMENT_POS)

	# gets the scope of the method
	if match.group(METHOD_SCOPE_POS) == None:
		scope = 'public'
	elif match.group(METHOD_SCOPE_POS) == '':
		scope = 'public'
	else:
		scope = match.group(METHOD_SCOPE_POS).lower()

	# case where there are no parmaters
	if match.group(PARAMATERS_POS) == None:
		methodParamsStr = match.group(METHOD_NAME_POS) + '()'
	# case when there are paramaters
	else:
		methodParamsStr = match.group(METHOD_NAME_POS) + match.group(PARAMATERS_POS)

	# removes white spaces and unwanted keywords
	methodParamsStr = pattern.sub('', methodParamsStr)

	# allows for different formatting of subs and functions
	if match.group(METHOD_TYPE).lower() == 'sub':
		methodParamsStr = methodParamsStr.replace('()', '').replace('(', ' ').replace(')', '')
	return comment, scope, methodParamsStr