from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import xlsxwriter
from datetime import date
from test_lists import *


def click_xpath(driver, xpath):
    try:
        element = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, xpath)))
        element.click()
    except Exception as e:
        print(f'error clicking element XPath {xpath}: {e}')


def send_form_fill_xpath(driver, xpath, payload):
    driver.find_element(By.XPATH, f'{xpath}').send_keys(payload)


def send_form_fill_id(driver, id_path, payload):
    driver.find_element(By.ID, f'{id_path}').send_keys(payload)


def current_url(driver):
    return driver.current_url


def write_workbook(test_results):

    full_date = date.today()
    today = full_date.strftime("%d%m%y")
    headers = ["Category", "Element Tested", "Expected", "Result"]
    workbook_name = f'Olympia Automated Test: Thomas OLeary 1 - {today}.xlsx'
    workbook = xlsxwriter.Workbook(workbook_name)

    worksheet = workbook.add_worksheet()

    class TableData:
        def __init__(self, i, test_results):
            self.i = i
            self.test_results = test_results

        def write_section(self):
            # headers
            if self.i == 0:
                title_cell_format = workbook.add_format(
                    {'bold': True, 'font_size': 14, 'font_color': 'black', 'align': 'center'})
                title_cell_format.set_bottom(2)
                worksheet.merge_range(
                    'A1:D1', "Olympia Financial Automated Tests by Thomas O'Leary", title_cell_format)

            if self.i == 0:
                cell_format = workbook.add_format(
                    {'bold': True, 'font_size': 14, 'font_color': 'black'})
                cell_format.set_bottom(2)
                worksheet.write_row(f"A2", headers, cell_format)

            # results
            if self.i != 0:

                cell_format = workbook.add_format(
                    {'bold': False, 'font_size': 12, 'font_color': 'black'})

                # Seperate format for result based on value.
                result_cell_format = workbook.add_format(
                    {'bold': False, 'font_size': 12, 'font_color': 'black', 'align': 'center'})
                result_cell_format.set_bottom(1)
                result = self.test_results[self.i - 1][3]

                if result == "Passed":
                    result_cell_format.set_bg_color('green')
                if result == "Failed":
                    result_cell_format.set_bg_color('red')

                worksheet.write(
                    f"A{self.i + 2}", self.test_results[self.i - 1][0], cell_format)
                worksheet.write(
                    f"B{self.i + 2}", self.test_results[self.i - 1][1], cell_format)
                worksheet.write(
                    f"C{self.i + 2}", self.test_results[self.i - 1][2], cell_format)
                worksheet.write(
                    f"D{self.i + 2}", result, result_cell_format)

    for i in range(0, len(test_results) + 1):
        t = TableData(i, test_results)
        t.write_section()

    # Set column widths based on string lengths so all values visible
    max_test_length = max([len(str(test)) for test in nav_test_list])
    max_expected_length = max([len(str(expected))
                              for expected in expected_urls])

    # Set the column width to the maximum length of the text in each column
    worksheet.set_column('A:A', 12)
    worksheet.set_column('B:B', max_test_length + 1)
    worksheet.set_column('C:C', max_expected_length + 1)
    worksheet.set_column('D:D', 8)
    # worksheet.autofit()
    workbook.close()
    print(f"Workbook Done saved as {workbook_name}")
