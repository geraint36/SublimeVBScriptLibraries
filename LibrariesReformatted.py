# assumes that the files that this is being run on contain valid VB syntax
# also requires the current file to be in a directory called '\testlibrary\'

# WARNINGS: 
# - possible errors if functions are commented out (generally not allowed for comments)

#import sublime_plugin
#import sublime
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
# cleans up the path variables
del path
"""

from VBScriptLibraryUtil import ImportDetails

VBSCRIPT_ALLOW_VAR_NAME_REGEX = '\\b[a-zA-Z]{1}[a-zA-Z0-9_]{,254}\\b'
VBSCRIPT_VARIABLE_BASIC_REGEX = 'a-zA-Z0-9_'
FILE_FOLDER_NAME_REGEX = 'a-zA-Z0-9_\\-'
LIBRARY_PARENT_FOLDER = '\\TestLibrary\\'
POSSIBLE_SCRIPT_PARENT_FOLDERS = ['\\TestLibrary\\', '\\RegressionControl\\']

class VBSParameterDetails:
	def __init__(self, paramType, name):
		# 'ByVal' or 'ByRef'
		self.paramType = paramType
		self.name = name

	# just for testing
	def toString(self):
		return (self.paramType, self.name)

class VBSVariableDetails:
	def __init__(self, name, value):
		self.name = name
		# what the variable contains
		# only used if it is an object and has other details inside it like class with methods etc.
		# set to None otherwise
		self.value = value

	# just for testing
	def toString(self):
		return (self.name, self.value)

class VBSPropertyDetails(VBSVariableDetails):
	def __init__(self, scope, name, value):
		# probably should use .super() but differece between Python 2 and 3 plus don't really know how it works
		VBSVariableDetails.__init__(self, name, value)
		# 'Private' or 'Public'
		self.scope = scope

	# just for testing
	def toString(self):
		return (self.scope,) + VBSVariableDetails.toString(self)

class VBSMethodDetails:
	def __init__(self, scope, methType, name, params):
		# 'Private' or 'Public'
		self.scope = scope
		# 'Function', 'Sub', 'Property Let' or 'Property Set'
		self.methType = methType
		self.name = name
		# a list of VBSParameterDetails classes
		self.params = params

	# just for testing
	def toString(self):
		return (self.scope, self.methType, self.name, self.params)

class VBSClassDetails:
	def __init__(self, properties, methods):
		# a list of VBSPropertyDetails classes
		self.properties = properties
		# a list of VBSMethodDetails classes
		self.methods = methods

	# just for testing
	def toString(self):
		return (self.properties, self.methods)

if __name__ == '__main__':
	a = VBSPropertyDetails('public', 'IR', 'CD235434')
	print(a.toString())
	print(__file__)