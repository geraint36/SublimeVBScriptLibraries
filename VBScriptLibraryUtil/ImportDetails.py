#import re
import os
import codecs
#import time

class FileNotFoundException(Exception):
	def __init__(self,*args,**kwargs):
		Exception.__init__(self,*args,**kwargs)

class FileEncodingNotFoundException(Exception):
	def __init__(self,*args,**kwargs):
		Exception.__init__(self,*args,**kwargs)

class IncorrectFileExtensionException(Exception):
	def __init__(self,*args,**kwargs):
		Exception.__init__(self,*args,**kwargs)

def getLibraryContents(path, removeComments=True, removeWhiteSpaces=True):
	outputContents = ''
	with openTryEncodings(path) as f:
		for line in f:
			if removeWhiteSpaces:
				line = line.strip()

			outputContents += line + '\n'

	return outputContents

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