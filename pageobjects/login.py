from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class Login:

    def __init__(self, driver):
        self.driver = driver
        self.username_textbox = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.ID, "id_auth-username")))
        self.password_textbox = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.ID, "id_auth-password")))
        self.submit_button = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.XPATH, '//button[@type="submit"]')))
        self.cookies_button = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.ID, "hs-eu-confirmation-button")))


    def enter_username(self, username):
        self.username_textbox.clear()
        self.username_textbox.send_keys(username)


    def enter_password(self, password):
        self.password_textbox.clear()
        self.password_textbox.send_keys(password)


    def click_submit(self):
        self.submit_button.click()


    def accept_cookies(self):
        self.cookies_button.click()

