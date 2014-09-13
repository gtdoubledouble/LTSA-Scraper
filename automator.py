from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from xlrd import open_workbook
from xlwt import Workbook

''' 
This script scrapes the owner names given the PID number from the LTSA website
You will need your username and login
All the PID numbers in the first column in a file called pids.xlsx
All the output numbers+names will be in names.xlsx

You will need selenium, xlrd, xlwt, and also ChromeDriver (brew install ChromeDriver)


Helpful links:
http://selenium-python.readthedocs.org/en/latest/locating-elements.html
http://www.qaautomation.net/?p=188
http://www.simplistix.co.uk/presentations/python-excel.pdf

Notes:
If you cannot get the element and it throws an exception, then its because the Javascript app or page has not fully loaded
So you need to force a manual wait with time.sleep()
For this, implement "try until no error" using a while-loop + try + except

Dummies guide:
Selenium is a framework that launches a browser and automates the browsing based on code, as if you were controlling it
It uses a select browser, so Chrome, Firefox, IE are all good to go but you need the driver
You open websites using driver.get('url'), and there are commands like driver.back(), driver.forward(), etc.
The challenge in this script is just finding the proper element (text fields, form submit) then manipulating it to your liking
So you can locate elements in many different ways, such as find_element_by_name, or find_elements_by_name (returns a list). 
IDs, tags, classes are also game.
The returned WebElement will have a .get_attribute() method where you could do .get_attribute('value') or just go webElement.text

One of the biggest challenges I had was when you can't find a certain element because the internet I was using was too slow.
So you have to keep trying again with manual time delays in between.



''' 

class PIDGetter(object):

	def __init__(self):
		self.driver = webdriver.Chrome()
		self.driver.get('https://myltsa.ltsa.ca/myltsalogin')

	def login(self, user, pw):
		username = self.driver.find_element_by_name("username")
		password = self.driver.find_element_by_name("password")
		username.send_keys(user)
		password.send_keys(pw)
		loginForm = self.driver.find_element_by_name("loginForm")
		loginForm.submit()

	def get_pid(self, pid_num):
		pid_entered = False
		while not pid_entered:	
			try:
				pid = self.driver.find_element_by_id("titleSearchNumberId")
				print type(pid_num), pid_num
				pid.send_keys(pid_num)
				pid_entered = True
			except Exception as e:
				time.sleep(1) # sleep or else the angular app won;t load

		search_clicked = False
		while not search_clicked:
			try:
				searchBtn = self.driver.find_element_by_tag_name('button')
				searchBtn.click()
				search_clicked = True
			except:
				time.sleep(1)

		name_located = False
		owner_name = ''
		while not name_located:
			try:
				# the owner name exists in the 4th icp-center table column
				name = self.driver.find_elements_by_class_name('icp-center')
				owner_name = name[3].text
				name_located = True
			except Exception as e:
				time.sleep(1)
				# print 'NAME LOCATOR', e

		# return the pid

		# selenium should go back a page
		self.driver.back()
		return owner_name


class ExcelWriter(object):


	def __init__(self):
		# grab excel file and columns
		self.wb_read = open_workbook('hudson.xlsx')
		self.wb_write = Workbook()
		self.output = self.wb_write.add_sheet('PID and Names')

	def go(self):
		row_counter = 1
		for sheet in self.wb_read.sheets():
			for row in range(sheet.nrows):
				pid_to_query = sheet.cell(row,3).value.replace('-','')
				pid_to_query = str(pid_to_query)
				owner_name = pid_getter.get_pid(pid_to_query)

				print row_counter, "Query of PID", sheet.cell(row,0).value, "Owner name = ", owner_name

				# write to new file
				# self.output.write(row_counter,0,pid_to_query)
				# self.output.write(row_counter,1,owner_name)

				# writing to a new file by copying everything else
				for col in range(sheet.ncols):
					self.output.write(row_counter,col,sheet.cell(row,col).value)
				self.output.write(row_counter,sheet.ncols,owner_name)

				self.wb_write.save('names.xlsx')

				row_counter += 1 

pid_getter = PIDGetter()
pid_getter.login("___YOUR USERNAME HERE___", "__YOUR PASSWORD HERE__")
excel_writer = ExcelWriter()
excel_writer.go()







