import logging
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
import configparser
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import sys
import datetime
import traceback
import os
import csv

class UPDATE_INSPECTION:

    def __init__(self):

        self.path = os.getcwd()

        # create logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)

        # create console handler
        ch = logging.StreamHandler()

        #create file handler 
        date = str(datetime.datetime.now().strftime('%d-%b-%Y %H_%M_%S %p'))

        fh = logging.FileHandler(f'{self.path}\\debug\\{date}.log',encoding='utf-8')

        # create formatter
        formatter = logging.Formatter('%(asctime)s - %(funcName)s - %(lineno)d - %(levelname)s - %(message)s',datefmt='%d/%b/%Y %I:%M:%S %p')

        # add formatter to ch
        ch.setFormatter(formatter)

        #add formatter to fh
        fh.setFormatter(formatter)

        # add ch to logger
        self.logger.addHandler(ch)

        #add fh to logger
        self.logger.addHandler(fh)


        #config.init file
        self.my_config_parser = configparser.ConfigParser()

        self.my_config_parser.read(f'{self.path}\\config\\config.ini')

        self.config = { 

        'email': self.my_config_parser.get('config','email'),
        'password': self.my_config_parser.get('config','password'),


        }
        
        
        self.can_not_update_dict = {}

        self.incorrect_fa_status = {}
        
        self.passed_csv = 'passed.csv'
        self.failed_csv = 'failed.csv'
    
    def login(self):

        self.driver=webdriver.Chrome()

        self.driver.get('https://www-plmprd.cisco.com/Agile/')

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="userInput"]'))).send_keys(self.config["email"])

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="login-button"]'))).click()

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="passwordInput"]'))).send_keys(self.config["password"])

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="login-button"]'))).click()

        count_render_2fa = 0

        while (self.driver.title != "Universal Prompt"):

            sleep(1)

            count_render_2fa+=1

            self.logger.info("Wait for Universal Prompt render:"+str(count_render_2fa))
            
        
        while True:
            
            try:

                WebDriverWait(self.driver, 60).until(ec.visibility_of_element_located((By.XPATH, '//button[@id="trust-browser-button"]'))).click()
                
                break
                
            except:
                
                self.logger.warning('Not found trust browser button')

        two_fa_url=self.driver.current_url

        count_duo_pass = 0

        while(two_fa_url==self.driver.current_url):

            sleep(1)

            count_duo_pass+=1

            self.logger.info("Wait for count_duo_pass:"+str(count_duo_pass))

            if count_duo_pass == 30:

                self.logger.warning("!!! LOGIN TIMEOUT !!!")

                self.driver.quit()

                sys.exit()

        self.logger.info("Login to QIS is success!!!")

        sleep(1)

        self.driver.get('https://www-plmprd.cisco.com/Agile/')

        self.main_page = self.driver.current_window_handle

        self.logger.debug("Main Page:"+self.main_page)

        handles = self.driver.window_handles

        for handle in handles:

            sleep(1)

            self.driver.switch_to.window(handle)

            if self.main_page != self.driver.current_window_handle:

                self.driver.close()

        self.driver.switch_to.window(self.main_page)

        self.driver.maximize_window()

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Collapse Left Navigation"]'))).click()

        return True


    def extract_data_excel(self):                                                                                                                                            

        self.fa_dict = {}

        self.df = pd.read_excel("qit_disposition_template.xlsm")
        
        self.logger.info(self.df)

        index = self.df.index

        number_of_rows = len(index)
        
        self.logger.info('Number of rows >>> '+str(number_of_rows))

        for i in range(number_of_rows):

            if self.df['FA#'][i] in self.fa_dict:

                self.fa_dict[self.df['FA#'][i]].update({self.df['Site Received Serial#'][i]:[self.df['QIT Disposition'][i], str(self.df['Problem Description'][i]).replace('_x000D_', ''),self.df['Case Owner'][i],self.df['PID'][i]]})
            
            else:

                self.fa_dict[self.df['FA#'][i]] = {self.df['Site Received Serial#'][i]: [self.df['QIT Disposition'][i], str(self.df['Problem Description'][i]).replace('_x000D_', ''),self.df['Case Owner'][i],self.df['PID'][i]]}

        
                #remove FA case that already done
        with open(self.passed_csv, newline='', encoding='UTF8') as f:
            all_fa_done = csv.reader(f)
            for fa in all_fa_done:
                
                try:
                
                    self.fa_dict.pop(fa[0])
                    
                except:
                    
                    pass
        
        self.logger.debug('After Filter FA case that already done.')
        
        for i,j in self.fa_dict.items():

            self.logger.debug(i + str(j))

        self.logger.info('FA Case number >>> '+str(self.fa_dict.__len__()) + ' Cases')


    def update_qms_data(self,fa_case):
        
        try:
            
            if not self.search_case(fa_case,'Inspection & Review'):
                
                return False
            
            WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()
        
            WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//select[@name="R1_2023_7"]//option[text() = "FA Canceled"]'))).click()
            
            WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2000003227_7"]'))).send_keys('The PID "End of FA" date has passed and the FA site no longer has the test infrastructure to support FA.')
            
            WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()
            
            
            self.search_case(fa_case,'Inspection & Review')
            
            WebDriverWait(self.driver, 20).until(ec.visibility_of_all_elements_located((By.XPATH, '//div[@id="tabsDiv"]//li')))[-5].click()

            WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//select[@name="TABLE_VIEWS_LIST_1"]//option[@title="QIT_Disposition"]'))).click()

            sleep(1)

            rows_exact = int(WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//strong[@id="totalCount_PSRTABLE_AFFECTEDITEMS"]'))).text)

            self.logger.debug(str(rows_exact))

            # get html page source

            # self.htmlSource = self.driver.page_source

            # with open("htmlSource.txt",'w',encoding='utf-8') as f:

            #     f.write(self.htmlSource)
            

            #handle to case which has more than 30 units.
            exit_loop_more_than_30_cases = False
            while True:
                
                while True:

                    WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//tr[@class="GMHeaderRow"]//span[@title="QIT  Disposition PID/Comp"]'))).click()
                    
                    sleep(1)
                    
                    try:
                    
                        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//img[@title="Ascending"]')))
                        
                        self.logger.debug('Check Sort is pass.')
                        
                        break
                    
                    except:
                        
                        self.logger.warning('Chech Sort is fail.')
            
                sleep(1)

                rows = WebDriverWait(self.driver, 20).until(ec.visibility_of_all_elements_located((By.XPATH, '//tr[@class="GMDataRow"]')))

                rows_len = len(rows)

                self.logger.info("Exact rows >>> "+str(rows_exact)+" Rows number >>> "+str(rows_len))

                if int(rows_exact)*2 != rows_len:

                    self.logger.critical(fa_case + ': Rows number does not match !!!')

                    self.err.update({fa_case: 'Rows number does not match'})

                row_start = int(rows_len/2)
                
                for i in range(row_start, rows_len):

                    row = rows[i]

                    self.logger.debug(row)
                    
                    entries = row.find_elements(By.TAG_NAME,'td')
                    
                    if int(rows_exact) >= 30:
                        
                        bulkUnit = True
                        
                        dispose = entries[10]
                        
                        if dispose.text.strip() != '':
                
                            WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="MSG_Save_1span"]'))).click()
                            
                            self.logger.debug(fa_case+ ' is bulk units that complete.')
                            
                            exit_loop_more_than_30_cases = True
                                    
                            break
                        
                    else:
                        
                        bulkUnit = False
                        

                    fa_flag = entries[2]

                    self.logger.info("FA FLAG >>> "+fa_flag.text)

                    entries[1].click()

                    if fa_flag.text.strip() == 'Yes':

                        sn = entries[4].text.strip()

                        if sn == "":

                            self.err.update({fa_case:' >>> Not found Serial Number'})

                            self.logger.critical(fa_case+' >>> Not found Serial Number')

                            self.press_down()

                            continue

                        self.logger.info('Serial Number >>> '+sn)

                        try:

                            self.logger.info('Dispose to >>> '+self.fa_dict[fa_case][sn][0])
                            
                            if self.fa_dict[fa_case][sn][0].lower() != 'scrap':

                                self.err.update({fa_case:' >>> Incorrect Disposition >>> '+sn})

                                self.logger.critical(fa_case+' >>> Incorrect Disposition >>> '+sn)

                                self.press_down()

                                continue
                                
                        except:

                            self.err.update({fa_case:' >>> Not found Disposition >>> '+sn})

                            self.logger.critical(fa_case+' >>> Not found Disposition >>> '+sn)

                            self.press_down()

                            continue

                    else:

                        self.press_down()

                        continue

                    dispose = entries[10]

                    if dispose.text.strip() == '':

                        ActionChains(self.driver).double_click(dispose).perform()

                        self.press_down(time =4)

                    self.press_enter()
                    
                    
                    ai = entries[12]
                    
                    ActionChains(self.driver).double_click(ai).perform()
                    
                    WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//a[@title="Search to add"]'))).click()
                    
                    checkRender = WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//a[@title="PID Result/Failure Mode/Failure Code : PID Result/Failure Mode/Failure Code"]'))).text
                    
                    self.logger.debug('Check Render >>> '+checkRender)
                    
                    self.press_down(2)
                    
                    self.press_right()
                    
                    self.press_down()
                    
                    self.press_right()
                    
                    self.press_down(2)
                    
                    self.press_enter()
                    
                    WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="close_floater_R1_2000008011_0_display"]'))).click()

                    self.press_enter()
                    
                    self.press_down()

                WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="MSG_Save_1span"]'))).click()
                
                if exit_loop_more_than_30_cases:
                    
                    break
                
                #repeat for bulk units
                if bulkUnit:
                    
                    self.search_case(fa_case,'Inspection & Review')

                    WebDriverWait(self.driver, 20).until(ec.visibility_of_all_elements_located((By.XPATH, '//div[@id="tabsDiv"]//li')))[-5].click()
                    
                    WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//select[@name="TABLE_VIEWS_LIST_1"]//option[@title="QIT_Disposition"]'))).click()
                    
                else:
                    
                    #exit loop more than 30 cases
                    break

            self.search_case(fa_case,'Inspection & Review')
            
            if self.move_case(fa_case,'PCA'):
                
                if self.auto_close_case(fa_case):
            
                    return True
        
        finally:
            
            return False
        

    def move_case(self,case,target):
            
        not_reset_audit = True

        while(not_reset_audit):

            try :   

                WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//em[@id="MSG_NextStatus_em"]'))).click()

                if target == 'PCA':

                    WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//*[ text() = "Pending Closure Approval" ]'))).click()
                    
                    WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="ewfinish"]'))).click()


                elif target == 'RMA':

                    WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="MSG_NextStatusspan"]'))).click()
                    
                    WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="ewfinish"]'))).click()


                elif target == 'FI':
                    
                    WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//*[ text() = "Fault Isolation" ]'))).click()
                    
                not_reset_audit = False
                
            except:

                reset_handle = False

                handles = self.driver.window_handles

                self.logger.debug("No. Of window handles1: "+str(len(handles))+", "+str(handles))

                for handle in handles:

                    self.driver.switch_to.window(handle)

                    window_title = self.driver.title

                    if window_title == 'Application Error':

                        self.logger.error("Application Error Window: "+window_title+" , " +handle)

                        reset_handle = True

                        self.driver.close()

                if reset_handle:
                    
                    self.driver.switch_to.window(self.main_page)

                    while True:

                        try:

                            WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="top_refreshspan"]'))).click()

                            break
                        
                        except:

                            self.logger.warning("Wait for against click intercepted !!!")

                            sleep(1)


        #window handle
        found_change_status_window = False
        count_open_change_status_window = 0
        while(not found_change_status_window):

            reset_handle = False
            sleep(1)#9-Dec-2021  add delay
            handles = self.driver.window_handles
            size = len(handles)
            self.logger.debug("No. Of window handles2: "+str(size)+' >>>  '+str(handles))

            for handle in handles:
                
                self.driver.switch_to.window(handle)
                window_title = self.driver.title
                if window_title == 'Change Status':
                    self.logger.debug("Change Status Window: "+window_title+' >>> '+str(handles))
                    found_change_status_window = True
                    break
                elif window_title == 'Application Error':
                    self.logger.error("Application Error Window: "+window_title+' >>> '+str(handles))

                    sleep(1)#9-Dec-2021  add delay
                    reset_handle = True
                    sleep(1)#9-Dec-2021  add delay
                    self.driver.close()

            if reset_handle:
                
                self.driver.switch_to.window(self.main_page)
                
                WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="MSG_NextStatusspan"]'))).click()  #9-Dec-2021  visible to clickable
                
                WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="ewfinish"]'))).click()    #9-Dec-2021
                
            count_open_change_status_window+=3
            
            if count_open_change_status_window > 10:
            
                self.logger.critical("Can't open change_status Window: "+str(count_open_change_status_window)+" second")
                self.can_not_update_state[case] = "Can't Open Change Status Window"
                break

        

        # if target == 'PCA':

        #     WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//a[@class="delete_button"]'))).click()  

        #     sleep(1)

        #     WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="search_query_approvers_display"]'))).send_keys('wlongsid'+Keys.ENTER)
            
        # elif target == 'FI':

        #     WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//a[@class="delete_button"]'))).click()  

        #     sleep(1)

        #     WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="search_query_notify_display"]'))).send_keys('wlongsid'+Keys.ENTER)        

        # sleep(5)              

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="save"]'))).click()  


        count_close_change_status_window = 0
        while (len(self.driver.window_handles) > 1):
            sleep(1)
            count_close_change_status_window+=1
            self.logger.warning("Wait for Close Change Status Window: "+ str(count_close_change_status_window)+" second")
            if count_close_change_status_window > 10:
                self.logger.critical("Can't Close Change Status Window: "+str(count_close_change_status_window))
                self.can_not_update_state[case] = "Can't Close Change Status Window"
                sleep(1)  #9-Dec-2021  add delay
                self.driver.close()
                sleep(1)  #9-Dec-2021  add delay
                break       
        sleep(1)
        
        self.driver.switch_to.window(self.main_page)
        
        return True      


    def search_case(self,fa_case,expect_fa_status):

            WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(Keys.CONTROL+'a',Keys.BACKSPACE)

            WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(fa_case)

            self.logger.info('FA Case >>> '+fa_case)

            while True:

                try:

                    WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()

                    break
                
                except:

                    self.logger.warning("Wait for against click intercepted !!!")

                    sleep(1)
            
            #fix bug in case that change priority 
            while True:

                try:
                    
                    self.fa_status = WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//h2[@style="color:Blue;"]'))).text
                    
                    self.logger.info('Status of '+fa_case+' >>> '+self.fa_status)
                    
                    #check case status
                    if self.fa_status != expect_fa_status:
                        
                        self.incorrect_fa_status[fa_case] = f'FA Status Value : {self.fa_status}  ,  Expected Status : {expect_fa_status}'
                        
                        self.logger.warning('Skip '+fa_case+ ' >>> The status is not accept.')
                        
                        return False

                    return True

                except:

                    self.logger.warning("This case wase changed the priority !!!")

                    if fa_case in WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//h4[@id="searchResultHeader"]'))).text:

                        rows_exact = int(WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//strong[@id="totalCount_QUICKSEARCH_TABLE"]'))).text)

                        self.logger.debug(str(rows_exact))

                        rows = WebDriverWait(self.driver, 20).until(ec.visibility_of_all_elements_located((By.XPATH, '//tr[@class="GMDataRow"]')))

                        rows_len = len(rows)

                        self.logger.info("Exact rows >>> "+str(rows_exact)+" Rows number >>> "+str(rows_len))

                        if int(rows_exact)*2 != rows_len:

                            self.logger.critical(fa_case + ': Rows number does not match !!!')

                            self.err.update({fa_case: 'Rows number does not match'})

                        row_start = int(rows_len/2)
                        
                        for i in range(row_start, rows_len):

                            row = rows[i]
                            
                            self.logger.debug(row)

                            entries = row.find_elements(By.TAG_NAME,'td')

                            fa_link = entries[3]

                            self.logger.info("FA Link : " + fa_link.text)

                            if fa_case == fa_link.text.strip():

                                fa_link.click()

                                sleep(1)
                                
                                break
                            
                    else:

                        self.logger.critical(fa_case+" : Page is not rendering !!!")
                        
                        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(Keys.CONTROL+'a',Keys.BACKSPACE)

                        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(fa_case)

                        WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()   


    def auto_close_case(self,fa_case):
            
        try:

            if not self.search_case(fa_case,'Pending Closure Approval'):
                
                return False
        
            WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Approvespan"]'))).click()

            found_approve_window = False
            count_open_approve_window = 0
            while(not found_approve_window):

                handles = self.driver.window_handles
                size = len(handles)
                print("No. Of window handles:",size,handles)

                for handle in handles:
                    sleep(1)
                    self.driver.switch_to.window(handle)
                    window_title = self.driver.title
                    if window_title == 'Approve':
                        print("Approve Window:",window_title,handle)
                        found_approve_window = True
                        break

                count_open_approve_window+=3
                if count_open_approve_window > 10:
                    print("Can't open Approve Window:",count_open_approve_window,"second")
                    self.can_not_update_dict[fa_case] = "Can't open Approve Window"
                    break


            # html = driver.page_source

            # with open('html.txt', 'w', encoding='utf-8-sig') as file:
            #     file.write((html))
            if count_open_approve_window > 10:
                
                self.driver.switch_to.window(self.main_page)
                
                return False

            else:

                WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="savespan"]'))).click()

                count_close_approve_window = 0
                while (len(self.driver.window_handles) > 1):
                    sleep(1)
                    count_close_approve_window+=1
                    print("Wait for close Approve Window:",count_close_approve_window,"second")
                    if count_close_approve_window > 10:
                        print("Can't Close Approve Window:",count_close_approve_window)
                        self.can_not_update_dict[fa_case] = "Can't Close Approve Window"
                        self.driver.close()
                        break                        
          
        except:

            self.can_not_update_dict[fa_case] = "Can't Access Case"


        self.driver.switch_to.window(self.main_page)
        
        return True


    def press_down(self,time =1):

        for i in range(time):

            ActionChains(self.driver).send_keys(Keys.DOWN).perform()


    def press_right(self,time =1):

        for i in range(time):

            ActionChains(self.driver).send_keys(Keys.RIGHT).perform()


    def press_enter(self,time =1):

        for i in range(time):

            ActionChains(self.driver).send_keys(Keys.ENTER).perform()


    def main(self):

        self.login()

        self.extract_data_excel()
        
        for case in self.fa_dict:
            
            if self.update_qms_data(case):
                
                with open(self.passed_csv, 'a',newline = '', encoding='UTF8') as f:

                    # create the csv writer
                    writer = csv.writer(f)

                    # write a row to the csv file
                    writer.writerow([case])
            else:
                
                with open(self.failed_csv, 'a',newline = '', encoding='UTF8') as f:

                    # create the csv writer
                    writer = csv.writer(f)

                    # write a row to the csv file
                    writer.writerow([case])
                    
                

        self.logger.info("INCORECT FA STATUS DICT: "+str(self.incorrect_fa_status))

        self.logger.info("CAN NOT UPDATE DICT: "+str(self.can_not_update_dict))
    
    
    
    
    
if __name__ == '__main__':

    try:

        run = UPDATE_INSPECTION()
        
        run.main()
        
        
    finally:

        run.logger.critical("Traceback Error: "+traceback.format_exc())

        while input('Please enter \'e\' for exit!') != 'e':

            pass

        run.driver.quit()

        sys.exit()

