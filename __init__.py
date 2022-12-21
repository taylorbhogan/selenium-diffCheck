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
		time.sleep(1)
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
			driver.quit()

		try:
			driver.switch_to.alert.accept()
			handleSuccess()
		except:
			handleFailure(driver, file_name)


	driver = webdriver.Chrome()
	driver.get('https://www.diffchecker.com/excel-compare/')
	time.sleep(1)

	originalInput = driver.find_element(By.XPATH, "//input[@id='fileOriginal-Spreadsheet']")
	changedInput = driver.find_element(By.XPATH, "//input[@id='fileChanged-Spreadsheet']")

	originalDirectoryPath = os.fsencode('/Users/chefables_imac/Desktop/automation/test_files/dev')
	changedDirectoryPath = os.fsencode('/Users/chefables_imac/Desktop/automation/test_files/prod')

	originalDirectoryContentsUnfiltered = os.listdir(originalDirectoryPath)
	changedDirectoryContentsUnfiltered = os.listdir(changedDirectoryPath)

	originalDirectoryContents = sorted(list(filter(lambda val: os.fsdecode(os.path.splitext(val)[1]) == '.xlsx', originalDirectoryContentsUnfiltered)))
	changedDirectoryContents = sorted(list(filter(lambda val: os.fsdecode(os.path.splitext(val)[1]) == '.xlsx', changedDirectoryContentsUnfiltered)))
	time.sleep(1)

	for idx in range(len(originalDirectoryContents)):
		originalPath = f'{os.fsdecode(originalDirectoryPath)}/{os.fsdecode(originalDirectoryContents[idx])}'
		changedPath = f'{os.fsdecode(changedDirectoryPath)}/{os.fsdecode(changedDirectoryContents[idx])}'

		originalInput.send_keys(originalPath)
		changedInput.send_keys(changedPath)

		time.sleep(2)

		findDifference(driver, os.fsdecode(changedDirectoryContents[idx]))

		time.sleep(1)

		sheetNameSelects = driver.find_elements(By.XPATH, "//select[@class='excel-input_sheetSelect__p7MmN diffResult']")
		sheet_idx = 1
		options = sheetNameSelects[0].find_elements(By.TAG_NAME, "option")
		while sheet_idx < len(options):
			time.sleep(1)
			sheetNameSelects[0].click()
			options[sheet_idx].click()

			sheetNameSelects[1].click()
			optionsRight = sheetNameSelects[1].find_elements(By.TAG_NAME, "option")
			optionsRight[sheet_idx].click()
			
			time.sleep(1)

			findDifference(driver, os.fsdecode(changedDirectoryContents[idx]))

			sheet_idx += 1
			time.sleep(1)


		print("iteration complete")
		time.sleep(1)

	print("script finished executing")

executeDiffChecking()