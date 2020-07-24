# Import the libraries and packages we will need and create a Unit Test Class that will have
# a setup, a teardown and our first test method (or test case).
import unittest
from pageobjects.homescreen import HomeScreen
from pageobjects.login import Login
from webdriver import Driver
from values import strings

class TestCCHQ(unittest.TestCase):

    def setUp(self):
        self.driver = Driver()
        #self.driver.maximize_window()
        self.driver.navigate(strings.base_url)
        login_page= Login(self.driver)
        login_page.enter_username(strings.login_username)
        login_page.enter_password(strings.login_password)
        login_page.click_submit()
        login_page.accept_alert()

    print("Done")

    #def test_home_screen_components(self):
    #   home_screen = Homescreen(self.driver)
    #   home_screen.validate_title_is_present()
    #   home_screen.validate_icon_is_present()
    #   home_screen.validate_menu_options_are_present()
    #   home_screen.validate_posts_are_visible()
    #   home_screen.validate_social_media_links()

    def tearDown(self):
        self.driver.instance.quit()


if __name__ == '__main__':
    unittest.main()
