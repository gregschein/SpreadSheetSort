import xlrd

workbook = xlrd.open_workbook('ETData2015.xlsx.xlsx')
worksheet = workbook.sheet_by_index(0)

iteration = 0
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