#!/usr/bin/python3

import argparse, datetime, logging, os, sys, yaml
from openpyxl import Workbook, styles
from neo4j import GraphDatabase
from timeit import default_timer as timer

def checkOutputPath(path):
	if os.path.exists(path):
		return True
	else:
		return False

def query_yes_no(question, default="no"):
	valid = {'yes': True, 'y': False, 'ye': True, 'no': False, 'n': False}

	if default == None:
		prompt = ' [y/n] '
	elif default.lower() == 'yes':
		prompt = ' [Y/n] '
	elif default.lower() == 'no':
		prompt = ' [y/N] '
	else:
		raise ValueError('Invalid default answer: {}'.format(default))

	while True:
		sys.stdout.write(question + prompt)
		choice = input().lower()
		if default is not None and choice == '':
			return valid[default]
		elif choice in valid:
			return valid[choice]
		else:
			sys.stdout.write("Please respond with 'yes' or 'no' (or 'y' or 'n').\n")

def newFileName(domain):
	return '{0:%Y%m%d%H%M%s}-{1}'.format(datetime.datetime.now(), domain)

def input_default(self, prompt, default):
	return raw_input("%s [%s] " % (prompt, default)) or default

class bhAnalytics(object):
	def __init__(self, logging, domain, outputFile, queryFile=None):
		# Default Database Values
		self.neo4jDB = 'bolt://localhost:7687'
		self.username = 'neo4j'
		self.password = 'neo4jj'
		self.domain = domain
		self.outputFile = outputFile
		self.getQueryFile(queryFile)
		self.driver = None
		self.db_validated = False
		self.logging = logging

	def getQueryFile(self, queryFile):
		if queryFile is None:
			queryFile = 'queries.yaml'

		with open(queryFile, 'r') as configFile:
			self.queryData = yaml.load(configFile, Loader=yaml.Loader)

	def updateDomain(self):
		self.domain = input_default('Enter the domain name.', self.domain)
		self.domain = self.domain.upper()

	def showDBInfo(self):
		print('The current database settings are:')
		print('Neo4j URL: {}'.format(self.neo4jDB))
		print('Username: {}'.format(self.username))
		print('Password: {}\n'.format(self.password))
	
	def updateDBInfo(self):
		self.neo4jDB = input_default('Enter the Neo4j Bolt URL:', self.neo4jDB)
		self.username = input_default('Enter the username:', self.username)
		self.password = input_default('Enter the password:', self.password)
	
	def connectDB(self):
		# Close the database if it is connected
		if self.driver is not None:
			self.driver.close()

		# Connect to the database
		while self.db_validated == False:
			self.driver = GraphDatabase.driver(self.neo4jDB, auth=(self.username, self.password))
			self.db_validated = self.validateDB()
			if not self.db_validated:
				print('Unable to validate provided domain and database settings. Please verify provided values:')
				self.updateDomain()
				self.updateDBInfo()

	def validateDB(self):
		print('Validating Selected Domain')
		session = self.driver.session()
		try:
			result = session.run("MATCH (n {domain:$domain}) RETURN COUNT(n)", domain=self.domain).value()
		except Exception as e:
			self.logging.debug('Unable to connect to the neo4j database')
			self.logging.error(e)

		if (int(result[0]) > 0):
			return True
		else:
			return False

def main(logging, domain, outputFile):
	# Instantiate a new bhAnalytics Object
	bhMain = bhAnalytics(logging, domain, outputFile)

	# Show the current DB settings and prompt to change
	bhMain.showDBInfo()
	if query_yes_no('Would you like to change the neo4j DB settings?'):
		bhMain.updateDBInfo()

	bhMain.connectDB()

if __name__ == "__main__":
	# Setup logging
	logLvl = logging.DEBUG
	logging.basicConfig(level=logLvl, format='%(asctime)s - %(levelname)s - %(message)s')
	logging.debug('Debugging logging is on.')

	# Parse the command line arguments and display a help message in incorrec tparameters are provided
	parser = argparse.ArgumentParser(description = 'Generate a report in spreadsheet format from a neo4j database containing BloodHound data.')
	parser.add_argument('-d', '--domain', type=str, help='The name of the domain that you wish to generate reporting for.')
	parser.add_argument('-o', '--outputfile', type=str, help='The path and name of the file that you wish the report to be written to.')
	args = parser.parse_args()
	
	# Check if the domain name was specified, print help if not
	if args.domain:
		domain = args.domain.upper()
		# If a output file name was specified, check to see if it exists
		
		if args.outputfile:
			if checkOutputPath(args.outputfile):
				answer = query_yes_no('The specified output file already exists. Do you wish to overwrite it?')
				if answer == True:
					outputFile = args.outputFile
				else:
					print('Please specify another output file name.')
					parser.print_help(sys.stderr)
		else:
			outputFile = newFileName(domain)
		# Call the main function
		main(logging, domain, outputFile)
	else:
		parser.print_help(sys.stderr)
