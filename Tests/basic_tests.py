'''
I have taken the liberty of writing some very basic tests of your website’s publicly facing frontend.
In the interest of time I used a combination of Python/Selenium. I have only separated the helper
functions, paths and constants while combining the tests to reduce the number of files in the
directory.

Test one checks navigation functionality in the dropdown menu located beside the Olympia logo. It
confirms that when each link is clicked it does in fact lead to the correct URL.

Test two confirms that required fields in the "Contact Us" form throw an error if submitted while
empty and that the error displayed is the correct message for that specific field.

Currently the results of the test are written to a XLSX file, which can be viewed in the README
file in my GitHub repository.

Thank you very much for your time,

Thomas O’Leary
All contact information is available on my resumé and cover letter.
'''

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from test_helpers import *
from test_paths import *
from test_lists import *

# Collect results
test_results = []

# Driver Variables
c_options = Options()
c_options.add_argument("disable-infobars")
c_options.add_argument("--disable-extensions")
driver = webdriver.Chrome(service=Service(
    ChromeDriverManager().install()), options=c_options)

# _________________________________  Navigation Tests _____________________________________

# Open home
olympia_home = "https://www.olympiafinancial.com/"
driver.get(olympia_home)

# Open the menu.
click_xpath(driver, menu_button_XPath)

# Sections of dropdown.
quick_links_id_path = driver.find_element(
    By.XPATH, '//*[@id="w-node-_16738520-6de8-1fcc-6cbc-d48e45d42b91-336b9d43"]/div[2]/div')
resources_id_path = driver.find_element(
    By.XPATH, '//*[@id="w-node-_17c43587-c790-b4d8-1466-5f40336b9d61-336b9d43"]/div[2]/div')
company_id_path = driver.find_element(
    By.XPATH, '//*[@id="w-node-e16c4357-004d-8209-e100-36bb04545bd4-336b9d43"]/div[2]/div')

# Collect the anchor tags.
menu_sections = [quick_links_id_path, resources_id_path, company_id_path]
anchor_tags = []
hrefs = []
for section in menu_sections:
    anchor_tags.extend(section.find_elements(By.TAG_NAME, "a"))

# Get actual hrefs from dom.
for anchor_tag in anchor_tags:
    href = anchor_tag.get_attribute("href")
    hrefs.append(href)

# Get handle of (main window)
main_window_handle = driver.current_window_handle

# Iterate through hrefs
for i in range(0, len(hrefs)):
    # Open in new tab to keep dropdown open in main window
    driver.execute_script("window.open('', '_blank');")

    # Switch tabs
    driver.switch_to.window(driver.window_handles[1])

    driver.get(hrefs[i])

    # Get the actual URL of the new tab
    actual_url = driver.current_url

    # Compare the actual and expected
    if actual_url == expected_urls[i]:
        test_results.append(
            ["Nav Link", f'{nav_test_list[i]}', f'{expected_urls[i]}', "Passed"])
    else:
        test_results.append(
            ["Nav Link", f'{nav_test_list[i]}', f'{expected_urls[i]}', "Failed"])

    # Close the second tab
    driver.close()

    # Switch back to the main window
    driver.switch_to.window(main_window_handle)

# _________________________________ 404 Test ______________________________________________

# check 404 with bogus route
driver.get("https://www.olympiafinancial.com/thisshouldntwork")
element = driver.find_element(
    By.XPATH, f'{error_404_message_XPath}')

if element.text == "404 PAGE NOT FOUND":
    test_results.append(["Nav Link", "404 Error Page",
                        "404 PAGE NOT FOUND", "Passed"])
else:
    test_results.append(["Nav Link", "404 Error Page",
                        "404 PAGE NOT FOUND", "Failed"])

# _________________________________  Form Tests ___________________________________________

form_url = "https://www.olympiabenefits.com/contact-us"
driver.get(form_url)

# Submit the form while empty.
click_xpath(driver, form_submit_XPath)

# Paths for expected error messages
form_input_fields = [
    error_message_first,
    error_message_email,
    error_message_phone,
    error_message_province,
    error_message_I_am,
    error_message_comment,
]

# Iterate through fields and verify the message not only appears but displays the correct text.
for i in range(0, len(form_input_fields)):
    error_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, form_input_fields[i]))
    )

    if error_element.text == expected_err_msg[i]:
        test_results.append(
            ["Form Field", f'{form_test_list[i]}', f'{expected_err_msg[i]}', "Passed"])

    else:
        test_results.append(
            ["Form Field", f'{form_test_list[i]}', f'{expected_err_msg[i]}', "Failed"])

# Close browser.
driver.quit()

# Create xlsx report.
write_workbook(test_results)
