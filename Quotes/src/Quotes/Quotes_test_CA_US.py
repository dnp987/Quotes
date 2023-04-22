'''
Created on 19May2020

@author: DNP Enterprises Inc.
'''
from Quotes.Excel_utils2 import Excel_utils2
import unittest
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class Log_Quotes(unittest.TestCase):
    @classmethod
    def setUpClass(inst):
        ''' setUpClass sets up all variables and the browser config '''
        inst.file_in = 'C:/temp/test-parameters.xlsx'
        inst.file_out = 'C:/temp/quotes-test-unittest.xlsx'
        inst.data_in = Excel_utils2(inst.file_in, 'test_parameters', 'in')
        inst.data_out = Excel_utils2(' ', 'Quotes', 'out')
        # Get account # and password
        inst.target_url = inst.data_in.sht.cell(2, 2).value
        inst.acc_num = inst.data_in.sht.cell(3, 2).value
        inst.passwd = inst.data_in.sht.cell(4, 2).value
        browser = "C:\\Selenium\\chromedriver.exe"
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--incognito")
        inst.driver = webdriver.Chrome(browser) # Open Chrome
        inst.wait = WebDriverWait(inst.driver, 10)

    def test_CheckHomePageTitle(self):
        ''' Load the test url '''
        self.driver.get(self.target_url) # Navigate to the test website
        page_title = "BMO InvestorLine - Account Access" 
        self.assertEqual(page_title, self.driver.title)
        
    def test_CheckLogin(self):
        ''' Verify that log in works '''
        page_title = "BMO InvestorLine - Home Page"
        # Enter account #
        log_in = self.driver.find_element_by_id("loginText")
        log_in.send_keys(self.acc_num)
        # Enter password
        password = self.driver.find_element_by_name("password")
        password.send_keys(self.passwd)
        log_in_button = self.driver.find_element_by_id("sasi_btn").click() # Click log in button to log in
        self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".home")))
        self.assertEqual(page_title, self.driver.title)
       
      
    def test_Check_Menu(self):
        ''' Vertify menu items '''
        # Verify the main menu
        expected_menu_items = ["Trading", "My Portfolio", "Quotes & Tools", "Markets & News"]
        actual_menu_items = ["#nav_trading_main > a:nth-child(1)", "#nav_portfolio_main > a:nth-child(1)", "#nav_quotes_main > a:nth-child(1)","#nav_markets_main > a:nth-child(1)"]
        
        j =0 
        for expected in expected_menu_items:
            actual_menu_found =  self.driver.find_element_by_css_selector(actual_menu_items[j]).text
            #self.assertEqual(actual_menu_found, expected_menu_items[j])
            self.assertEqual(actual_menu_found, expected)
            j+=1
       
    def test_Get_Quotes(self):
        ''' Get quotes using symbols from the excel sheet '''
        self.data_in = Excel_utils2(self.file_in, 'symbols', 'in') # Set the work sheet to the test stock symbols
        date_time = datetime.now().strftime('%Y-%B-%d %I:%M:%S %p') # get the date and time
        self.data_out.set_cell(1, 1, "Quotes as of: ", "Arial", True, 12) # first two rows contain the date and headinds
        self.data_out.set_cell(1, 2, date_time, "Arial", False, 10)                 
        self.data_out.set_cell(2, 1, "Symbol", "Arial", True, 12)
        self.data_out.set_cell(2, 2, "Exchange", "Arial", True, 12)
        self.data_out.set_cell(2, 3, "Last", "Arial", True, 12)
        self.data_out.set_cell(2, 4, "Bid", "Arial", True, 12)
        self.data_out.set_cell(2, 5, "Ask", "Arial", True, 12)
        self.data_out.set_cell(2, 6, "High", "Arial", True, 12)
        self.data_out.set_cell(2, 7, "Low", "Arial", True, 12)
        self.data_out.set_cell(2, 8, "Open", "Arial", True, 12)
        self.data_out.set_cell(2, 9, "Prev. Close", "Arial", True, 12)

        i=0
        for row in self.data_in.sht.rows:
            i+=1
            if (i == 1): # skip the 1st row of the symbol input spreadsheet, it contains headings
                continue
 
            test_symbol = self.data_in.sht.cell(i, 1).value
            test_exchange = self.data_in.sht.cell(i, 2).value       
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#nav_quotes_main > a:nth-child(1)"))) # Wait for the Quotes & Tools menu item to appear
            self.driver.find_element_by_css_selector("#nav_quotes_main > a:nth-child(1)").click() # click on Quotes & Tools
            self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.paddingLeft25:nth-child(2) > input:nth-child(2)"))) # Wait for the Symbol text box to appear
            symbol = self.driver.find_element_by_css_selector("div.paddingLeft25:nth-child(2) > input:nth-child(2)")
            # Select the Country drop down list &  set the exchange according to the symbol read in
            Select(self.driver.find_element_by_css_selector("div.paddingLeft25:nth-child(3) > select:nth-child(2)")).select_by_visible_text(test_exchange)
            symbol.send_keys(test_symbol, Keys.RETURN) # enter a symbol and get a quote
                   
            self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".page_navigation > div:nth-child(1) > a:nth-child(1)"))) # get a quote but wait for the "Summary" tab to appear
            last_price = self.driver.find_element_by_css_selector(".brownContainer > div:nth-child(1) > span:nth-child(1)").get_attribute("innerText")
            bid_price = self.driver.find_element_by_css_selector("#table1 > table > tbody > tr.tableRowBG > td:nth-child(2)").get_attribute("innerText")
            ask_price = self.driver.find_element_by_css_selector("#table1 > table > tbody > tr:nth-child(2) > td:nth-child(2)").get_attribute("innerText")
            high = self.driver.find_element_by_css_selector("#table1 > table > tbody > tr.tableRowBG > td:nth-child(4)").get_attribute("innerText")
            low =  self.driver.find_element_by_css_selector("#table1 > table > tbody > tr:nth-child(2) > td:nth-child(4)").get_attribute("innerText")
            open_price = self.driver.find_element_by_css_selector("#table1 > table > tbody > tr.tableRowBG > td:nth-child(6)").get_attribute("innerText")
            prev_close = self.driver.find_element_by_css_selector("#table1 > table > tbody > tr:nth-child(2) > td:nth-child(6)").get_attribute("innerText")
        
            self.data_out.set_cell(i+1, 1, test_symbol, "Arial", False, 10) # setup the quote in a spreadsheet
            self.data_out.set_cell(i+1, 2, test_exchange, "Arial", False, 10)
            self.data_out.set_cell(i+1, 3, last_price, "Arial", False, 10)
            self.data_out.set_cell(i+1, 4, bid_price, "Arial", False, 10)
            self.data_out.set_cell(i+1, 5, ask_price, "Arial", False, 10)
            self.data_out.set_cell(i+1, 6, high, "Arial", False, 10)
            self.data_out.set_cell(i+1, 7, low, "Arial", False, 10)
            self.data_out.set_cell(i+1, 8, open_price, "Arial", False, 10)
            self.data_out.set_cell(i+1, 9, prev_close, "Arial", False, 10)
        
    @classmethod       
    def  tearDownClass(inst):
        ''' Save the quotes in a spreadsheet and close the browser '''
        inst.data_out.save_file(inst.file_out)
        inst.driver.quit()
            
if __name__ == '__main__':
    unittest.main()