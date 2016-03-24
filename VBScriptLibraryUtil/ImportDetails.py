import os, codecs, re
#import time

VBSCRIPT_VAR_NAME_PATTERN = '\\b[a-zA-Z]{1}[a-zA-Z0-9_]{0,254}\\b'

class FileNotFoundException(Exception):
	def __init__(self,*args,**kwargs):
		Exception.__init__(self,*args,**kwargs)

class FileEncodingNotFoundException(Exception):
	def __init__(self,*args,**kwargs):
		Exception.__init__(self,*args,**kwargs)

class IncorrectFileExtensionException(Exception):
	def __init__(self,*args,**kwargs):
		Exception.__init__(self,*args,**kwargs)

class VBSScope:
	def __init__(self):
		self.content = ''

	def appendContent(text):
		self.content += (text)

class VBSGlobalScope(VBSScope):
	def __init__(self):
		VBSScope.__init__(self)

# used for all scopes apart from the global scope
class VBSBlockScope(VBSScope):
	SCOPE_MODIFIERS_PATTERN = '(\\bpublic\\b|\\bprivate\\b)?'
	startPattern = None
	endPattern = None

	def __init__(self):
		VBSScope.__init__(self)

	@classmethod
	def isStart(cls, line):
		return (None != re.match(cls.startPattern, line))

	@classmethod
	def isEnd(cls, line):
		return (None != re.match(cls.endPattern, line))

class VBSBlockScopeClass(VBSBlockScope):
	# setup up the regexp pattern static class variables
	VBSBlockScope.startPattern = ( '^%s\\s*\\bClass\\b\\s+%s$' % \
		(VBSBlockScope.SCOPE_MODIFIERS_PATTERN, VBSCRIPT_VAR_NAME_PATTERN) )
	VBSBlockScope.endPattern = ( '^\\bEnd\\b\\s+\\bClass\\b$' )

	def __init__(self, scope, name):
		VBSBlockScope.__init__(self)
		self.scope = scope
		self.name = name		

class VBSBlockScopeMethod(VBSBlockScope):
	PARAMS_TYPE_PATTERN = '\\bByVal\\b|\\bByRef\\b'
	METHOD_SINGLE_PARAM_PATTERN = ( '\\s*(%s)?\\s*(%s)\\s*' % (PARAMS_TYPE_PATTERN, VBSCRIPT_VAR_NAME_PATTERN) )
	METHOD_PARAMS_PATTERN = ( '\\((%s,)*(%s)?\\)' % \
		(METHOD_SINGLE_PARAM_PATTERN, METHOD_SINGLE_PARAM_PATTERN) )

	def __init__(self, blockStartLine):
		VBSBlockScope.__init__(self)
		match = re.match(self.startPattern, blockStartLine)

		if match == None:
			raise Exception()

		groups = match.groupdict()

		if 'scope' in groups:
			self.scope = groups['scope']
		else:
			self.scope = 'Public'

		self.name = groups['name']

		if 'params' in groups:
			self.params = groups['params']
		else:
			self.params = []

	@classmethod
	def setStartPattern(cls, blockIdentifierPattern):
		cls.startPattern = ( '^(?P<scope>%s)?\\s*(?P<name>%s)\\s+%s\\s*(?P<params>%s)?$' % (VBSBlockScope.SCOPE_MODIFIERS_PATTERN, \
			blockIdentifierPattern, VBSCRIPT_VAR_NAME_PATTERN, VBSBlockScopeMethod.METHOD_PARAMS_PATTERN) )

	@classmethod
	def setEndPattern(cls, blockIdentifierPattern):
		cls.endPattern = ( '^\\bEnd\\b\\s+%s$' % blockIdentifierPattern )

class VBSBlockScopeFunction(VBSBlockScopeMethod):
	def __init__(self, blockStartLine):
		VBSBlockScopeMethod.__init__(self, blockStartLine)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bFunction\\b")
		cls.setEndPattern("\\bFunction\\b")

class VBSBlockScopeSub(VBSBlockScopeMethod):
	def __init__(self, blockStartLine):
		VBSBlockScopeMethod.__init__(self, blockStartLine)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bSub\\b")
		cls.setEndPattern("\\bSub\\b")

class VBSBlockScopePropertyGet(VBSBlockScopeMethod):
	def __init__(self, blockStartLine):
		VBSBlockScopeMethod.__init__(self, blockStartLine)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bProperty\\b\\s+\\bGet\\b")
		cls.setEndPattern("\\bProperty\\b")

class VBSBlockScopePropertyLet(VBSBlockScopeMethod):
	def __init__(self, blockStartLine):
		VBSBlockScopeMethod.__init__(self, blockStartLine)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bProperty\\b\\s+\\bLet\\b")
		cls.setEndPattern("\\bProperty\\b")

class VBSBlockScopePropertySet(VBSBlockScopeMethod):
	def __init__(self, blockStartLine):
		VBSBlockScopeMethod.__init__(self, blockStartLine)		
		
	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bProperty\\b\\s+\\bSet\\b")
		cls.setEndPattern("\\bProperty\\b")

# sets up the classes start and end patterns for the different method classes
VBSBlockScopeFunction.setupPatterns()
VBSBlockScopeSub.setupPatterns()
VBSBlockScopePropertyGet.setupPatterns()
VBSBlockScopePropertyLet.setupPatterns()
VBSBlockScopePropertySet.setupPatterns()

VBSCRIPT_NON_GLOBAL_SCOPE_CLASSES = [VBSBlockScopeClass, VBSBlockScopeFunction, VBSBlockScopeSub, \
	VBSBlockScopePropertyGet, VBSBlockScopePropertyLet, VBSBlockScopePropertySet]

# returns formatted string for file with comments removed long with leading and trailing whitespaces
def getLibraryScopesFormatted(path):
	outputContents = ''
	# holds the scope for the current line (if empty then in global scope)
	# array used becase of possibility of methods inside a class
	currentScopeStack =[]
	globalScope = VBSGlobalScope()
	scopes = [globalScope]

	# builds formatted lines array
	lines = []
	with openTryEncodings(path) as f:
		buildLine = ''
		for line in f:
			# removes comments and white spaces
			formattedLine = removeComments(line).strip(' \t')

			# skip lines that are just white spaces and comments
			if len(formattedLine) == 0:
				continue

			# to remove line continuation
			if (formattedLine[-1] == '_'):
				buildLine += formattedLine
			else:
				buildLine += '\n'

				# checks if is end of current scope
				if len(currentScopeStack) > 0:
					if currentScopeStack[-1].isEnd(buildLine):
						print('end of scope="%s"' % currentScopeStack[-1].__name__)
						currentScope = currentScopeStack.pop(-1)
						scopes.append(currentScope)
						continue

				# checks to see if is the start of a new scope
				# returns None if is not the start of a new scope
				newScope = getNewScope(line)
				if newScope != None:
					# add code to build scope class properly
					currentScope.append(newScope)
					continue

				# adds the line to the current scope
				if len(currentScopeStack) > 0:
					currentScopeStack[-1].appendContent(buildLine)
				else:
					globalScope.appendContent(buildLine)
				# clears the line that has been built from memory
				buildLine = ''

	return outputContents

def getNewScope(line):
	for scope in VBSCRIPT_NON_GLOBAL_SCOPE_CLASSES:
		if scope.isStart(line):
			print(scope.__name__)
			return scope(line)

	return None

def removeComments(line):
	inStr = False
	pos = 0
	for char in line:
		if char == '"':
			inStr = not inStr
		# exit loop when comment starts
		elif (not inStr) and (char == "'"):
			break

		pos += 1

	return line[:pos]

def isVbScriptFile(path):
	return (os.path.splitext(path)[1].lower() in ('.vbs', '.qfl'))

def openTryEncodings(path, encType='r'):
	# can be found at 'https://docs.python.org/2.4/lib/standard-encodings.html'
	# doesn't have correct encoding for .qfl files (must be saved again as UTF-8)
	encodings = ['utf-8', 'utf-16', 'ascii']
	for e in encodings:
		try:
			codecs.open(path, encType, e).read()
			return codecs.open(path, encType, e)
		except UnicodeDecodeError as e:
			continue

	raise FileEncodingNotFoundException('no encoding found for the file %s' % path)