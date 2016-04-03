import os, codecs, re
#import time

# \w is equivalent to [a-zA-Z0-9_] in a regex
VBSCRIPT_VAR_NAME_PATTERN = '\\b[a-zA-Z]{1}\\w{0,254}\\b'

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

# extended bt VBScriptVariable, VBScriptBlockFunction and VBScriptBlockPropertyGet
class VBScriptCanReturnValue(object):
	def __init__(self):
		self.alreadyAskedForValue = False
		self.returnValue = None

	def getValue(self):
		# caching of value
		if self.alreadyAskedForValue:
			return self.returnValue
		self.alreadyAskedForValue = True

		# TODO - finish will work with methods return values too
		outputValue = self.getContents()
		# TODO - make seperate class for things with return values and make variable, function and property get extend it too
		prevVals = []
		while issubclass( outputValue.__class__, VBScriptCanReturnValue ):
			prevVals.append(outputValue)
			outputValue = outputValue.getContents()

			# to prevent infinite recursion which can occure for example when a=b and b=a
			if (outputValue in prevVals):
				return None

		self.returnValue = outputValue
		return outputValue

	def getContents(self):
		raise NotImplementedError('.getContents() not implemented for the class=%s' % self.__class__.__name__)

class VBScriptVariable(VBScriptElement, VBScriptCanReturnValue):
	pattern = ( '^(?P<type>Set )?\\s*(?P<name>%s)\\s*=(?P<value>.+)$' % VBSCRIPT_VAR_NAME_PATTERN )
	string_pattern = '"(""|[^"])*"' # needs more adding doesn't allow for escaped " chars ("""")
	number_pattern = '([1-9][0-9]*|0)(\\.[0-9])?'
	# \w is equivalent to [a-zA-Z0-9_]
	call_expression_pattern = '[\\w. \\(\\),]'

	def __init__(self, line, lineNo, comment, globalScope):
		VBScriptElement.__init__(self)
		VBScriptCanReturnValue.__init__(self)
		match = self.getMatch(line)

		if (match == None):
			raise ValueError("'Could not construct class='%s' from line='%s'" % \
				(self.__class__.__name__, line))

		self.lineNo = lineNo
		self.globalScopeRef = globalScope

		self.comment = comment
		groups = match.groupdict()
		self.name = groups['name']
		self.contentsCalculated = False
		self.value = groups['value']
		# used for debugging
		self.valueStr = groups['value']

		# if 'Set' keyword is used then is a reference to a variable, otherwise is a copy of value
		if ('type' in groups):
			self.type = 'Reference'
		else:
			self.type = 'Value'

	def getContents(self):
		if not self.contentsCalculated:
			self.contentsCalculated = True
			print('before=%s' % self.value)
			currentScope = self.globalScopeRef.getLineCombinedScope(self.lineNo)
			self.parseValue(currentScope)
			
			print('after=%s' % self.value)

		return self.value

	@classmethod
	def isVar(cls, line):
		return (None != cls.getMatch(line))

	@classmethod
	def getMatch(cls, line):
		return re.match(cls.pattern, line)

	@classmethod
	def formatMethodVariableName(cls, methodStr):
		# to allow for cases like '( varname ) '
		methodStr = methodStr.strip(' ()')
		pos = methodStr.find('(')
		# trims of parameters (incase of methods)
		if pos >= 0:
			methodStr = methodStr[:pos]

		return methodStr

	def getName(self):
		return self.name

	@classmethod
	def isString(cls, inputValue):
		return None != re.match( cls.string_pattern, inputValue, re.IGNORECASE )

	@classmethod
	def isNumber(cls, inputValue):
		return None != re.match( cls.number_pattern, inputValue, re.IGNORECASE )

	@classmethod
	def isCallExpression(cls, inputValue):
		return None != re.match( cls.call_expression_pattern, inputValue, re.IGNORECASE )

	# TODO - addin code for things of the form 'var.meth(param1, param2).var2' etc. (allow for dots)
	@classmethod
	def parseSingleExpression(cls, inputValue, scope):
		formattedValue = cls.formatMethodVariableName(inputValue)
		# if is a string
		if cls.isString(formattedValue):
			return formattedValue[1:-1]
		# if is a number
		elif cls.isNumber(formattedValue):
			return float(formattedValue)
		# case for variable/methods and classes (with possible calls to sub methods etc.)
		elif cls.isCallExpression(formattedValue):
			callStack = formattedValue.split('.')
			print(callStack)
			#########################################################################################
			# TODO - finish this code too tired at the moment (parseSingleExpression and parseExpression)
			#########################################################################################

			for i in range(len(callStack)):
				call = cls.formatMethodVariableName(callStack[i])
				output = scope

				# only classes can have sub properties and methods
				if ( i < (len(callStack)-1) ):
					# if value is variable that is defined in the current scope
					if output.containsVariable(call):
						output = output.getVariable(call).getValue()
					# if value is code block (class or method) that is defined in the current scope
					elif output.containsSubBlock(call):
						output = output.getSubBlock(call)
				# if last element in the call stack then doesn't need to be a class
				elif (i == (len(callStack) - 1)):
					# if value is variable that is defined in the current scope
					if output.containsVariable(call):
						output = output.getVariable(call)
					# if value is code block (class or method) that is defined in the current scope
					elif output.containsSubBlock(call):
						output = output.getSubBlock(call)
					else:
						return None
					return output
				else:
					return None
		
		# if unknown return None
		return None

	@classmethod
	def parseExpression(cls, inputValue, scope):
		formattedValue = cls.formatMethodVariableName(inputValue)
		callStack = formattedValue.split('.')
		print("callstack=%s" % callStack)
		# TODO - finish this code too tired at the moment
		
		for i in range(len(callStack)):
			call = cls.formatMethodVariableName(callStack[i])
			output = scope

			output = cls.parseSingleExpression(call, output)
		
		return output

	def parseValue(self, scope):
		self.value = self.parseExpression(self.value, scope)

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

class VBScriptScope(VBScriptElement):
	def __init__(self):
		self.scopeRange = None
		# of the form {name:varClassInstance, ...}
		self.variables = {}
		# of the form {identifier:scriptBlockClassInstance, ...}
		self.blocks = {}
		self.ended = False

	@classmethod
	def formatKey(cls, key):
		return key.lower()

	def addVariable(self, var):
		name = self.formatKey(var.getName())
		# overwrites previous keys if they exist
		self.variables[name] = var

	def addSubBlock(self, block):
		name = self.formatKey(block.getName())
		# raises error if the key already exists
		if (name in self.blocks):
			raise ValueError('Method already found with the name=%s' % name)
		else:
			self.blocks[name] = block

	# used when build a combined scope
	# will overwrite variables and methods with the ones in the new scope if there are clashes
	def addScope(self, scope):
		# adds the variables
		for name, var in scope.variables.iteritems():
			self.variables[name] = var
		# adds the sub-blocks
		for identifier, block in scope.blocks.iteritems():
			self.blocks[identifier] = block

	def containsVariable(self, name):
		return self.formatKey(name) in self.variables

	def getVariable(self, name):
		return self.variables[ self.formatKey(name) ]

	def getVariables(self):
		return self.variables.values()

	def containsSubBlock(self, name):
		return self.formatKey(name) in self.blocks

	def getSubBlock(self, name):
		return self.blocks[ self.formatKey(name) ]

	def getSubBlocks(self):
		return self.blocks.values()

	def getLineCombinedScope(self, line):
		scopes = [self]
		
		for scope in scopes:
			for block in scope.getSubBlocks():		
				if issubclass(VBScriptScopeGlobal, block.__class__):
					scopes.append(block)
				elif block.lineInScope(line):
					scopes.append(block)
		
		
		print('line=%s scopes=%r' % (line, map(lambda x:x.__class__, scopes)))
		outputScope = VBScriptScope()
		for scope in scopes:
			outputScope.addScope(scope)

		return outputScope

	@classmethod
	def getNewScope(cls, line, comment, lineNo):
		for scope in VBSCRIPT_NON_GLOBAL_SCOPE_CLASSES:
			if scope.isStart(line):
				return scope(line, comment, lineNo)

		return None

	def hasEnded(self):
		return (None != self.scopeRange)

	def getScopeRange(self):
		return self.scopeRange

	def lineInScope(self, line):
		scopeRange = self.getScopeRange()
		# should never be none when called (TODO - maybe raise error instead)
		return (scopeRange != None) and (line in scopeRange)

	def parseLine(self, line, comment, lineNo, globalScope):
		if self.isEnd(line):
			self.scopeRange = range(self.startLineNumber + 1, lineNo)
			return None

		# see if is the start of a new scope
		newScope = VBScriptScope.getNewScope(line, comment, lineNo)
		if None != newScope:
			self.addSubBlock(newScope)
			return newScope

		# see if line is a variable
		if VBScriptVariable.isVar(line):
			var = VBScriptVariable(line, lineNo, comment, globalScope)
			self.addVariable(var)
			return None
		
		# other cases
		return None

	@classmethod
	def isEnd(cls, line):
		raise NotImplementedError(".isEnd() method not implemented by class='%s'" % cls.__name__)

	def getName(self):
		raise NotImplementedError('.getName() methods has not been implemented for the class=%s' % self.__class__.__name__)

class VBScriptScopeGlobal(VBScriptScope):
	def __init__(self):
		VBScriptScope.__init__(self)

	@classmethod
	def isEnd(cls, line):
		# line is never the end of the global scope
		return False

# inherited by all scopes apart from the global scope
class VBScriptBlock(VBScriptScope):
	SCOPE_MODIFIERS_PATTERN = '(\\bpublic\\b|\\bprivate\\b)?'
	startPattern = None
	endPattern = None

	def __init__(self, blockStartLine, comment, lineNo):
		VBScriptScope.__init__(self)

		# instance variable as other constructor may wish to do more with it
		self.match = self.matchStart(blockStartLine)

		if self.match == None:
			raise ValueError("'Could not construct class='%s' from line='%s'" % \
				(self.__class__.__name__, blockStartLine))

		self.startLineNumber = lineNo
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

	def getName(self):
		return self.name

class VBScriptBlockClass(VBScriptBlock):
	# setup up the regexp pattern static class variables
	startPattern = ( '^(?P<scope>%s)\\s*\\bClass\\b\\s+(?P<name>%s)$' % \
		(VBScriptBlock.SCOPE_MODIFIERS_PATTERN, VBSCRIPT_VAR_NAME_PATTERN) )
	endPattern = ( '^\\bEnd\\b\\s+\\bClass\\b$' )

	def __init__(self, blockStartLine, comment, lineNo):
		VBScriptBlock.__init__(self, blockStartLine, comment, lineNo)

class VBScriptBlockMethod(VBScriptBlock):
	PARAMS_TYPE_PATTERN = '\\bByVal\\b|\\bByRef\\b'
	METHOD_SINGLE_PARAM_PATTERN = ( '\\s*(%s)?\\s*(%s)\\s*' % (PARAMS_TYPE_PATTERN, VBSCRIPT_VAR_NAME_PATTERN) )
	METHOD_PARAMS_PATTERN = ( '\\((%s,)*(%s)?\\)' % \
		(METHOD_SINGLE_PARAM_PATTERN, METHOD_SINGLE_PARAM_PATTERN) )

	def __init__(self, blockStartLine, comment, lineNo):
		VBScriptBlock.__init__(self, blockStartLine, comment, lineNo)

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

	def getName(self):
		return self.name

class VBScriptBlockFunction(VBScriptBlockMethod, VBScriptCanReturnValue):
	def __init__(self, blockStartLine, comment, lineNo):
		VBScriptBlockMethod.__init__(self, blockStartLine, comment, lineNo)
		VBScriptCanReturnValue.__init__(self)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bFunction\\b")
		cls.setEndPattern("\\bFunction\\b")

	def getContents(self):
		if self.containsVariable(self.name):
			return self.getVariable(self.name)
		else:
			return None

class VBScriptBlockSub(VBScriptBlockMethod):
	def __init__(self, blockStartLine, comment, lineNo):
		VBScriptBlockMethod.__init__(self, blockStartLine, comment, lineNo)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bSub\\b")
		cls.setEndPattern("\\bSub\\b")

class VBScriptBlockPropertyGet(VBScriptBlockMethod, VBScriptCanReturnValue):
	def __init__(self, blockStartLine, comment, lineNo):
		VBScriptBlockMethod.__init__(self, blockStartLine, comment, lineNo)
		VBScriptCanReturnValue.__init__(self)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bProperty\\b\\s+\\bGet\\b")
		cls.setEndPattern("\\bProperty\\b")

	def getContents(self):
		if self.containsVariable(self.name):
			return self.getVariable(self.name)
		else:
			return None

class VBScriptBlockPropertyLet(VBScriptBlockMethod):
	def __init__(self, blockStartLine, comment, lineNo):
		VBScriptBlockMethod.__init__(self, blockStartLine, comment, lineNo)

	# needs to be called before the class pattern parameters are used (could find nice way to run static init block of code)
	@classmethod
	def setupPatterns(cls):
		cls.setStartPattern("\\bProperty\\b\\s+\\bLet\\b")
		cls.setEndPattern("\\bProperty\\b")

class VBScriptBlockPropertySet(VBScriptBlockMethod):
	def __init__(self, blockStartLine, comment, lineNo):
		VBScriptBlockMethod.__init__(self, blockStartLine, comment, lineNo)		
		
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

# classes used to store and extract library details
class LibraryDetails(object):
	def __init__(self, path, useRelativePath=False):
		if useRelativePath:
			path = LibraryDetailsCache.getLibraryPath(path)

		self.path = path
		self.file_last_modified = os.path.getmtime(path)
		self.contents = parseVBScriptLibrary(path)

	def getLastModified(self):
		return self.file_last_modified

	def getContents(self):
		return self.contents

class LibraryDetailsCache(object):
	LIBRARY_PARENT_FOLDER = '\\TestLibrary\\'
	POSSIBLE_SCRIPT_PARENT_FOLDERS = ['\\TestLibrary\\', '\\RegressionControl\\']
	librariesDirPath = None
	libraries = {}

	def __init__(self):
		pass

	@classmethod
	def formatPath(cls, path):
		return path.lower()

	@classmethod
	def getDetails(cls, path):
		path = cls.formatPath(path)
		# see if library is already stored
		if (path in cls.libraries):
			libDetails = cls.libraries[path]
			# if the file has not been modified return it's details
			if (libDetails.getLastModified() == os.path.getmtime(path)):
				return libDetails.getContents()

		# if not added or outdated version extract details then return them
		cls.addLibrary(path)
		return cls.libraries[path].getContents()

	@classmethod
	def addLibrary(cls, path):
		path = cls.formatPath(path)
		cls.libraries[path] = LibraryDetails(path)

	@classmethod
	def getLibraryPath(cls, relativePath):
		return os.path.join( cls.librariesDirPath, relativePath )

	# needs to be called at least once before the class is used (call each time you need to 
	# the libraries folder path)
	@classmethod
	def findAndSetLibrariesFolderPath(cls, filePath):
		for parentDir in POSSIBLE_SCRIPT_PARENT_FOLDERS:
			pos = filePath.lower().find( parentDir.lower() )

			if pos >=0:
				# stores the direcory path
				cls.librariesDirPath = os.path.join( filePath[:pos], LIBRARY_PARENT_FOLDER )

		raise ValueError("Could not find the script in one of the parent folders \
			POSSIBLE_SCRIPT_PARENT_FOLDERS=%r" % POSSIBLE_SCRIPT_PARENT_FOLDERS)

# returns formatted string for file with comments removed long with leading and trailing whitespaces
def parseVBScriptLibrary(path):
	outputContents = ''
	# holds the scope for the current line (if empty then in global scope)
	# array used becase of possibility of methods inside a class	
	globalScope = VBScriptScopeGlobal()
	currentScopeStack =[globalScope]
	scopes = [globalScope]

	lines = getVBScriptLines(path)
	comment = None

	for line, pos in lines:
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
		newScope = currentScope.parseLine(line, comment, pos, globalScope)

		# if end of current scope
		if currentScope.hasEnded():
			oldScope = currentScopeStack.pop(-1)
			scopes.append(oldScope)
			continue

		if None != newScope:
			currentScopeStack.append(newScope)

	# raises error if a non-global block has not been closed
	if len(currentScopeStack) > 1:
		raise ValueError('Unclosed VBScript blocks=%r' % map(lambda x:[x.__class__, x.name],currentScopeStack[1:]) )

	return scopes

def getVBScriptLines(path):
	# list of the from [[line, pos], ...]
	lines = []
	# line indexes start a 1 to match sublime's line numbering
	pos = 1
	lastCodeLinePos = None
	with openTryEncodings(path) as f:
		for line in f:
			codeLines, comment = seperateLineIntoCodeAndComment(line)
			
			if comment != None:
				lines.append([comment, pos])

			for code in codeLines:
				if (lastCodeLinePos != None) and (len(lines[lastCodeLinePos]) > 0) and (lines[lastCodeLinePos][-1] == '_'):
					# maybe need to add a space charater in betwine the two lines
					lines[lastCodeLinePos] = lines[lastCodeLinePos][:-1] + code.strip(' \t')
					continue

				# if code line is split over multiple will line number of the last one
				# this will help with the ranges for the scopes (used inside the VBScriptScope.parseLine() method)
				lines.append([code, pos])
				lastCodeLinePos = len(lines) - 1
			pos += 1

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

def isVBScriptFile(path):
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