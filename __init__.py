import os
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup

def handleFailure(driver, file_name, sheet_index=0):
	time.sleep(1)
	screenshot_path = f'./screenshots/{time.time_ns()}-{os.path.splitext(file_name)[0]}-{sheet_index}.png'
	result = driver.get_screenshot_as_file(screenshot_path)

def handleSuccess():
	print("match")

def findDifference(driver, file_name):
	submit_button = driver.find_element(By.NAME, "Find difference")
	submit_button.click()

	try:
		driver.switch_to.alert.accept()
		handleSuccess()
	except:
		handleFailure(driver, file_name)


def run():
	driver = webdriver.Chrome()
	driver.get('https://www.diffchecker.com/excel-compare/')
	time.sleep(1)

	original_input = driver.find_element(By.XPATH, "//input[@id='fileOriginal-Spreadsheet']")
	changed_input = driver.find_element(By.XPATH, "//input[@id='fileChanged-Spreadsheet']")

	original_directory_path = os.fsencode('/Users/chefables_imac/Desktop/automation/test_files/dev')
	changed_directory_path = os.fsencode('/Users/chefables_imac/Desktop/automation/test_files/prod')

	original_directory_contents_unfiltered = os.listdir(original_directory_path)
	changed_directory_contents_unfiltered = os.listdir(changed_directory_path)

	original_directory_contents = sorted(list(filter(lambda val: os.fsdecode(os.path.splitext(val)[1]) == '.xlsx', original_directory_contents_unfiltered)))
	changed_directory_contents = sorted(list(filter(lambda val: os.fsdecode(os.path.splitext(val)[1]) == '.xlsx', changed_directory_contents_unfiltered)))
	time.sleep(1)

	for idx in range(len(original_directory_contents)):
		original_path = f'{os.fsdecode(original_directory_path)}/{os.fsdecode(original_directory_contents[idx])}'
		changed_path = f'{os.fsdecode(changed_directory_path)}/{os.fsdecode(changed_directory_contents[idx])}'

		original_input.send_keys(original_path)
		changed_input.send_keys(changed_path)

		time.sleep(1)

		findDifference(driver, os.fsdecode(changed_directory_contents[idx]))

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

			findDifference(driver, os.fsdecode(changed_directory_contents[idx]))

			sheet_idx += 1
			time.sleep(1)


		print("iteration complete")
		time.sleep(1)

	print("script finished executing")

run()