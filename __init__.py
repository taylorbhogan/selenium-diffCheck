import os
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

def executeDiffChecking():
	def handleFailure(driver, file_name, sheet_index=0):
		print("start handleFailure")
		screenshot_path = f'./screenshots/{time.time_ns()}-{os.path.splitext(file_name)[0]}-{sheet_index}.png'
		result = driver.get_screenshot_as_file(screenshot_path)

	def handleSuccess():
		print("match")

	def findDifference(driver, file_name):
		try:
			submitButton = WebDriverWait(driver, 15).until(
				EC.presence_of_element_located((By.NAME, "Find difference"))
			)

			submitButton.click()
		except:
			print("couldn't find the submit button")
			driver.quit()

		try:
			driver.switch_to.alert.accept()
			handleSuccess()
		except:
			handleFailure(driver, file_name)


	driver = webdriver.Chrome()
	driver.get('https://www.diffchecker.com/excel-compare/')

	originalDirectoryPath = os.fsencode('/Users/chefables_imac/Desktop/automation/test_files/dev')
	changedDirectoryPath = os.fsencode('/Users/chefables_imac/Desktop/automation/test_files/prod')

	originalExcelDocuments = sorted(list(filter(lambda val: os.fsdecode(os.path.splitext(val)[1]) == '.xlsx', os.listdir(originalDirectoryPath))))  #originalDirectoryContents
	changedExcelDocuments = sorted(list(filter(lambda val: os.fsdecode(os.path.splitext(val)[1]) == '.xlsx', os.listdir(changedDirectoryPath))))  #changedDirectoryContents

	for idx in range(len(originalExcelDocuments)): #each Excel doc
		inputs = WebDriverWait(driver, 4).until(
			EC.presence_of_all_elements_located((By.XPATH, "//input[@class='diff-input-header_fileInput__6v6Mq']"))
		)

		originalPath = f'{os.fsdecode(originalDirectoryPath)}/{os.fsdecode(originalExcelDocuments[idx])}'
		changedPath = f'{os.fsdecode(changedDirectoryPath)}/{os.fsdecode(changedExcelDocuments[idx])}'

		inputs[0].send_keys(originalPath)
		inputs[1].send_keys(changedPath)

		findDifference(driver, os.fsdecode(changedExcelDocuments[idx]))

		sheetNameSelects = WebDriverWait(driver, 15).until(
			EC.presence_of_all_elements_located((By.XPATH, "//select[@class='excel-input_sheetSelect__p7MmN diffResult']"))
		)
		sheet_idx = 1
		options = sheetNameSelects[0].find_elements(By.TAG_NAME, "option")
		while sheet_idx < len(options):  #each sheet after the first
			sheetNameSelects[0].click()
			options[sheet_idx].click()

			sheetNameSelects[1].click()
			optionsRight = sheetNameSelects[1].find_elements(By.TAG_NAME, "option")
			optionsRight[sheet_idx].click()

			findDifference(driver, os.fsdecode(changedExcelDocuments[idx]))

			sheet_idx += 1

		print("iteration complete")

	print("script finished executing")

executeDiffChecking()