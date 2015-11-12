import xlrd

workbook = xlrd.open_workbook('ETData2015.xlsx.xlsx')
worksheet = workbook.sheet_by_index(0)

# iteration = 0
test_name = []
test_values = [[] for thing in xrange(0,353)]


for x in range(0,353):
	test_name.append(worksheet.row(x)[0].value)
	iterrow = iter(worksheet.row(x))
	next(iterrow)
	for value in iterrow:
		test_values[x].append(value.value)



# outliers = {[] for thing in xrange(0,95)}
outliers = {thing:[] for thing in xrange(0,95)}



def find_median(sorted_list):
	return sorted_list[int(len(sorted_list)/2)]

def find_upper_quartile(sorted_list):
	return sorted_list[int(len(sorted_list)*3/4)]

def find_lower_quartile(sorted_list):
	return sorted_list[int(len(sorted_list)/4)]


def find_outliers(row):
	sorted_list = sorted(test_values[row])
	sorted_list = [x for x in sorted_list if x != '']
	lowerQ = find_lower_quartile(sorted_list)
	upperQ = find_upper_quartile(sorted_list)
	qRange = (upperQ - lowerQ)*3
	for value in test_values[row]:		
		if (value < (lowerQ - qRange)) or (value > (upperQ + qRange)):
			if value != '':
				outliers[test_values[row].index(value)+1].append(row)



def sort_similar_outliers(outlier_list):
	storage = []
	for test in outlier_list:
		storage.append(test_name[test])
		print test_name[test]
	storage.sort(key=lambda x: x[6:])
	print "________________________________"
	for test in storage:
		print test

for x in range(2, 353):
	find_outliers(x)

# from random import shuffle

# testing = outliers[66]
# print testing
# shuffle(testing)
# print testing
# sort_similar_outliers(testing)
# sort_similar_outliers(outliers[66])

class DataSheet(object):


	def initTestValues(self):
		workbook = xlrd.open_workbook('ETData2015.xlsx.xlsx')
		worksheet = workbook.sheet_by_index(0)
		self.test_name = []
		self.test_values = [[] for thing in xrange(0,353)]
		self.outliers = {thing:[] for thing in xrange(0,95)}
		for x in range(0,353):
			self.test_name.append(worksheet.row(x)[0].value)
			iterrow = iter(worksheet.row(x))
			next(iterrow)
			for value in iterrow:
				self.test_values[x].append(value.value)	

	def find_median(self, sorted_list):
		return sorted_list[int(len(sorted_list)/2)]

	def find_upper_quartile(self, sorted_list):
		return sorted_list[int(len(sorted_list)*3/4)]

	def find_lower_quartile(self, sorted_list):
		return sorted_list[int(len(sorted_list)/4)]

	def find_outliers(self, row):
		sorted_list = sorted(self.test_values[row])
		sorted_list = [x for x in sorted_list if x != '']
		lowerQ = self.find_lower_quartile(sorted_list)
		upperQ = self.find_upper_quartile(sorted_list)
		qRange = (upperQ - lowerQ)*3
		for value in self.test_values[row]:	
			if (value < (lowerQ - qRange)) or (value > (upperQ + qRange)):
				if value != '':
					self.outliers[self.test_values[row].index(value)+1].append(row)

	def sort_similar_outliers(self, outlier_list):
		storage = []
		for test in outlier_list:
			storage.append(self.test_name[test])
			print test_name[test]
		storage.sort(key=lambda x: x[6:])
		print "________________________________"
		for test in storage:
			print test

	def number_outliers(self, lot):
		print len(self.outliers[lot])

	def mainLoop(self):
		while 1:
			first_row = raw_input("First Row: ")
			if first_row == "End":
				break
			second_row = raw_input("Second Row: ")
			for x in range(int(first_row)+1, int(second_row)+1):
				self.find_outliers(x)
			print self.outliers
			lot = raw_input("Lot: ")
			print self.number_outliers(int(lot))
			self.outliers = {thing:[] for thing in xrange(0,95)}
SheetTest = DataSheet()

SheetTest.initTestValues()
SheetTest.mainLoop()