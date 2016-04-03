# assumes that the files that this is being run on contain valid VB syntax
# also requires the current file to be in a directory called '\testlibrary\'

# WARNINGS: 
# - possible errors if functions are commented out (generally not allowed for comments)

import sublime_plugin
import sublime
#import re
#import os
#import sys
#import codecs
#import time

# commented for now will need to add this (or something similar) when run as sublime package
"""
# adds the path of this module to the sys.path variable
# this will allow custom imports to work no matter where this file is executed from
path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if not path in sys.path:
	sys.path.insert(1, path)
# cleans up the path variable
del path
"""

import ImportDetails

FILE_FOLDER_NAME_REGEX = 'a-zA-Z0-9_\\-'
LIBRARY_PARENT_FOLDER = '\\TestLibrary\\'
POSSIBLE_SCRIPT_PARENT_FOLDERS = ['\\TestLibrary\\', '\\RegressionControl\\']


class ImportedClassesMethods(sublime_plugin.EventListener):
	# should be of the form {path:LibraryDetailsClassInstance, ... }
	libraryDetails = {}
	def __init__(self):
		pass

	# the return value of this function is what will appear in the auto complete (unless is is an empty list
	# in which case the standard autocompletions options will  be used)
	def on_query_completions(self, view, prefix, locations):
		words = []
		matches = []

		filePath = view.file_name()
		# exits if not in a vbscript file
		if not (ImportDetails.isVBScriptFile(filePath)):
			return []

		# to get the preceeding word (which could be a varaible storing a library)
		words = getVariableTreeBeforeCursor(view)

		# TODO - change so uses actual view text (including edits) instead of the saved file
		libDetails = ImportDetails.LibraryDetailsCache.getDetails(filePath)

		#print(libDetails[0])
		#print(libDetails[0].getVariable('aman').valueStr)
		print('fnc1 - %s' % libDetails[0].getSubBlock("fnc1").getValue())
		print('aman - %s' % libDetails[0].getVariable('aman').getValue())
		print('user - %s' % libDetails[0].getVariable('user').getValue())

		# if an empty list is returned from this method then the standard sublime suggestions will be used
		# this means that after any keyword that stores a library none of the standard suggestions will be 
		# available but everywhere else it'll just display the standard auto-complete options
		return matches

# gets the word preceeding the word that the cursor is curently at
def getVariableTreeBeforeCursor(view):
	# [0] is used because of the posiblility of multiple cursors
	region = view.sel()[0]
	# gets the start positions of the word that the cursor is currently at
	wordStart = view.word(region).begin()

	return getVariableTree(view, wordStart-1)

def getVariableTree(view, pos, inputWords=None):
	if None == inputWords:
		words = []
	else:
		# clones the list
		words = list(inputWords)

	# get the word region before the inputted position
	wordRegion = view.word( sublime.Region(pos, pos) )
	
	# gets the character before the word
	charBefore = view.substr( sublime.Region(wordRegion.begin()-1, wordRegion.begin()) )
	# gets the character after the word
	charAfter = view.substr( sublime.Region(wordRegion.end(), wordRegion.end()+1) )

	if charAfter != '.':
		return []

	# adds the string for the word from the region
	words.insert(0, view.substr(wordRegion))

	if charBefore == '.':
		words = getVariableTree(view, wordRegion.begin()-1, words)

	return words

if __name__ == '__main__':
	#a = ImportDetails.VBSBlockScopeClass("private", "test")
	a = ImportDetails.LibraryDetailsCache

	path = 'C:\\Users\\User\\Documents\\Geraint\\Programming\\QTP\\TestLibrary\\lib\\MethodsTest.qfl'
	import os
	print(os.path.isfile(path))
	b = ImportDetails.parseVBScriptLibrary(path)
	print('---------------')
	for scope in b:
		print('name=%s subblocks=%r variables=%r' % (scope.__class__.__name__, scope.blocks, scope.variables) )