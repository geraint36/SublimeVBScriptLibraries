import os, codecs, re
#import time

# \w is equivalent to [a-zA-Z0-9_] in a regex
VBSCRIPT_VAR_NAME_PATTERN = '\\b[a-zA-Z]{1}\w{0,254}\\b'

class FileNotFoundException(Exception):
	def __init__(self,*args,**kwargs):
		Exception.__init__(self,*args,**kwargs)

class FileEncodingNotFoundException(Exception):
	def __init__(self,*args,**kwargs):
		Exception.__init__(self,*args,**kwargs)

class IncorrectFileExtensionException(Exception):
	def __init__(self,*args,**kwargs):
		Exception.__init__(self,*args,**kwargs)

class VBScriptElement(object):
	def __init__(self):
		pass

	def parseLine(self, line, comment):
		raise NotImplementedError(".parseLine() method not implemented by class='%s'" % self.__class__.__name__)

class VBScriptVariable(VBScriptElement):
	pattern = ( '^(?P<type>Set )?\\s*(?P<name>%s)\\s*=(?P<value>.+)$' % VBSCRIPT_VAR_NAME_PATTERN )

	def __init__(self, line, comment):
		VBScriptElement.__init__(self)
		match = self.getMatch(line)

		if (match == None):
			raise ValueError("'Could not construct class='%s' from line='%s'" % \
				(self.__class__.__name__, blockStartLine))

		self.comment = comment
		groups = match.groupdict()
		self.name = groups['name']
		self.value = groups['value']

		# if 'Set' keyword is used then is a reference to a variable, otherwise is a copy of value
		if ('type' in groups):
			self.type = 'Reference'
		else:
			self.type = 'Value'

	@classmethod
	def isVar(cls, line):
		return (None != cls.getMatch(line))

	@classmethod
	def getMatch(cls, line):
		return re.match(cls.pattern, line)

class VBScriptParameter(VBScriptElement):
	pattern = ( '^(?P<type>ByVal |ByRef )?\\s*(?P<name>%s)$' % VBSCRIPT_VAR_NAME_PATTERN)

	def __init__(self, line):
		VBScriptElement.__init__(self)
		match = self.getMatch(line)

		if (match == None):
			raise ValueError("'Could not construct class='%s' from line='%s'" % \
				(self.__class__.__name__, line))

		groups = match.groupdict()
		self.name = groups['name']

		# if 'Set' keyword is used then is a reference to a variable, otherwise is a copy of value
		if ('type' in groups):
			self.type = groups['type']
		else:
			self.type = 'ByRef'

	@classmethod
	def isParam(cls, line):
		return (None != cls.getMatch(cls, line))

	@classmethod
	def getMatch(cls, line):
		return re.match(cls.pattern, line)

class VBSScriptScope(VBScriptElement):
	def __init__(self):
		self.elements = []
		self.ended = False

	def addElement(self, element):
		self.elements.append(element)

	@classmethod
	def getNewScope(cls, line, comment):
		for scope in VBSCRIPT_NON_GLOBAL_SCOPE_CLASSES:
			if scope.isStart(line):
				return scope(line, comment)

		return None

	def parseLine(self, line, comment):
		if self.isEnd(line):
			self.ended = True
			return None

		# see if is the start of a new scope
		newScope = VBSScriptScope.getNewScope(line, comment)
		if None != newScope:
			self.addElement(newScope)
			return newScope

		# see if line is a variable
		if VBScriptVariable.isVar(line):
			var = VBScriptVariable(line, comment)
			self.addElement(var)
			return None
		
		# other cases
		return None

	@classmethod
	def isEnd(cls, line):
		raise NotImplementedError(".isEnd() method not implemented by class='%s'" % cls.__name__)

class VBScriptScopeGlobal(VBSScriptScope):
	def __init__(self):
		VBSScriptScope.__init__(self)

	@classmethod
	def isEnd(cls, line):
		# line is never the end of the global scope
		return False

# inherited by all scopes apart from the global scope
class VBScriptBlock(VBSScriptScope):
	SCOPE_MODIFIERS_PATTERN = '(\\bpublic\\b|\\bprivate\\b)?'
	startPattern = None
	endPattern = None

	def __init__(self, blockStartLine, comment):
		VBSScriptScope.__init__(self)

		# instance variable as other constructor may wish to do more with it
		self.match = self.matchStart(blockStartLine)

		if self.match == None:
			raise ValueError("'Could not construct class='%s' from line='%s'" % \
				(self.__class__.__name__, blockStartLine))

		self.comment = comment
		groups = self.match.groupdict()

		if 'scope' in groups:
			self.scope = groups['scope']
		else:
			# default scope is Public
			self.scope = 'Public'

		self.name = groups['name']

	@classmethod
	def matchStart(cls, line):
		return re.match(cls.startPattern, line, re.IGNORECASE)

	@classmethod
	def isStart(cls, line):
		return (None != cls.matchStart(line))

	@classmethod
	def isEnd(cls, line):
		return (None != re.match(cls.endPattern, line, re.IGNORECASE))

class VBScriptBlockClass(VBScriptBlock):
	# setup up the regexp pattern static class variables
	startPattern = ( '^(?P<scope>%s)\\s*\\bClass\\b\\s+(?P<name>%s)$' % \
		(VBScriptBlock.SCOPE_MODIFIERS_PATTERN, VBSCRIPT_VAR_NAME_PATTERN) )
	endPattern = ( '^\\bEnd\\b\\s+\\bClass\\b$' )

	def __init__(self, blockStartLine, comment):
		VBScriptBlock.__init__(self, blockStartLine, comment)

class VBScriptBlockMethod(VBScriptBlock):
	PARAMS_TYPE_PATTERN = '\\bByVal\\b|\\bByRef\\b'
	METHOD_SINGLE_PARAM_PATTERN = ( '\\s*(%s)?\\s*(%s)\\s*' % (PARAMS_TYPE_PATTERN, VBSCRIPT_VAR_NAME_PATTERN) )
	METHOD_PARAMS_PATTERN = ( '\\((%s,)*(%s)?\\)' % \
		(METHOD_SINGLE_PARAM_PATTERN, METHOD_SINGLE_PARAM_PATTERN) )

	def __init__(self, blockStartLine, comment):
		VBScriptBlock.__init__(self, blockStartLine, comment)

		groups = self.match.groupdict()

		if 'params' in groups:
			# TODO add code to parse the parameters
			self.params = self.getParams( groups['params'] )
		else:
			self.params = []

	def getParams(self, paramsStr):
		if paramsStr == None:
			return []
		paramsStr = paramsStr.strip('()')
		if len(paramsStr) == 0:
			return []

		paramsList = paramsStr.split(',')
		params = []
		for text in paramsList:
			param = VBScriptParameter(text)
			params.append(param)
		return params

	@classmethod
	def setStartPattern(cls, blockIdentifierPattern):
		cls.startPattern = ( '^(?P<scope>%s)?\\s*%s\\s+(?P<name>%s)\\s*(?P<params>%s)?$' % (VBScriptBlock.SCOPE_MODIFIERS_PATTERN, \
			blockIdentifierPattern, VBSCRIPT_VAR_NAME_PATTERN, VBScriptBlockMethod.METHOD_PARAMS_PATTERN) )

	@classmethod
	def setEndPattern(cls, blockIdentifierPattern):
		cls.endPattern = ( '^\\bEnd\\b\\s+%s$' % blockIdentifierPattern )

class VBScriptBlockFunction(VBScriptBlockMethod):
	def __init__(self, blockStartLine, comment):
		VBScriptBlockMethod.__init__(self, blockStartLine, comment)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bFunction\\b")
		cls.setEndPattern("\\bFunction\\b")

class VBScriptBlockSub(VBScriptBlockMethod):
	def __init__(self, blockStartLine, comment):
		VBScriptBlockMethod.__init__(self, blockStartLine, comment)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bSub\\b")
		cls.setEndPattern("\\bSub\\b")

class VBScriptBlockPropertyGet(VBScriptBlockMethod):
	def __init__(self, blockStartLine, comment):
		VBScriptBlockMethod.__init__(self, blockStartLine, comment)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bProperty\\b\\s+\\bGet\\b")
		cls.setEndPattern("\\bProperty\\b")

class VBScriptBlockPropertyLet(VBScriptBlockMethod):
	def __init__(self, blockStartLine, comment):
		VBScriptBlockMethod.__init__(self, blockStartLine, comment)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bProperty\\b\\s+\\bLet\\b")
		cls.setEndPattern("\\bProperty\\b")

class VBScriptBlockPropertySet(VBScriptBlockMethod):
	def __init__(self, blockStartLine, comment):
		VBScriptBlockMethod.__init__(self, blockStartLine, comment)		
		
	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bProperty\\b\\s+\\bSet\\b")
		cls.setEndPattern("\\bProperty\\b")

# sets up the classes start and end patterns for the different method classes
VBScriptBlockFunction.setupPatterns()
VBScriptBlockSub.setupPatterns()
VBScriptBlockPropertyGet.setupPatterns()
VBScriptBlockPropertyLet.setupPatterns()
VBScriptBlockPropertySet.setupPatterns()

VBSCRIPT_NON_GLOBAL_SCOPE_CLASSES = [VBScriptBlockClass, VBScriptBlockFunction, VBScriptBlockSub, \
	VBScriptBlockPropertyGet, VBScriptBlockPropertyLet, VBScriptBlockPropertySet]

# returns formatted string for file with comments removed long with leading and trailing whitespaces
def parseVbScriptLibrary(path):
	outputContents = ''
	# holds the scope for the current line (if empty then in global scope)
	# array used becase of possibility of methods inside a class	
	globalScope = VBScriptScopeGlobal()
	currentScopeStack =[globalScope]
	scopes = [globalScope]

	lines = getVBScriptLines(path)
	comment = None

	for line in lines:
		if len(line) == 0:
			# clear comment
			comment = None
			continue

		# builds comment
		if line[0] == "'":
			if comment == None:
				comment = line[1:]
			else:
				comment += '\n' + line[1:]

		currentScope = currentScopeStack[-1]
		newScope = currentScope.parseLine(line, comment)

		# if end of current scope
		if currentScope.ended:
			oldScope = currentScopeStack.pop(-1)
			scopes.append(oldScope)
			continue

		if None != newScope:
			currentScopeStack.append(newScope)

	# raises error if a non-global block has not been closed
	if len(currentScopeStack) > 1:
		raise ValueError('Unclosed VbScript blocks=%r' % currentScopeStack[1:])

	return scopes

def getVBScriptLines(path):
	lines = []
	lastCodeLinePos = None
	with openTryEncodings(path) as f:
		buildLine = ''
		for line in f:
			codeLines, comment = seperateLineIntoCodeAndComment(line)
			
			if comment != None:
				lines.append(comment)

			for code in codeLines:
				if (lastCodeLinePos != None) and (len(lines[lastCodeLinePos]) > 0) and (lines[lastCodeLinePos][-1] == '_'):
					# maybe need to add a space charater in betwine the two lines
					lines[lastCodeLinePos] = lines[lastCodeLinePos][:-1] + code.strip(' \t')
					continue

				lines.append(code)
				lastCodeLinePos = len(lines) - 1

	return lines

def seperateLineIntoCodeAndComment(line):
	inStr = False
	pos = 0
	for char in line:
		if char == '"':
			inStr = not inStr
		# exit loop when comment starts
		elif (not inStr) and (char == "'"):
			break

		pos += 1

	codeLines = splitMultiLineCode(line[:pos])
	if (len(line) == pos):
		comment = None
	else:
		comment = line[pos:].strip()
	return codeLines, comment

def splitMultiLineCode(line):
	lines = []
	inStr = False
	pos = 0
	lastPos = 0
	for char in line:
		if char == '"':
			inStr = not inStr
		# exit loop when comment starts
		elif (not inStr) and (char == ":"):
			lines.append(line[lastPos:pos].strip())
			lastPos = pos + 1

		pos += 1

	# incase of empty line
	if len(line) > lastPos:
		lines.append(line[lastPos:].strip())
	return lines

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