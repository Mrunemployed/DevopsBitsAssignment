import json

import requests

import smtplib

from datetime import datetime,date

import re

import string

import time

import random

import smtplib

from email.message import EmailMessage

from requests.exceptions import ConnectionError

from selenium import webdriver

from selenium.webdriver.common.keys import Keys

from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.by import By

from selenium.common.exceptions import TimeoutException

from selenium.webdriver.support.ui import Select

from selenium.webdriver.chrome.options import Options

 

class utilities:

 

    #run headless

    options = Options()

    options.headless = True

    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(r"C:\Users\RUDRA\Documents\Py files\chromedriver\chromedriver.exe", options = options)

    #driver = webdriver.Chrome(r"C:\Users\RUDRA\Documents\Py files\chromedriver\chromedriver.exe")

    keyval = ""

 

    def pwd_gen(self):

        passwd = ''

        while passwd != None:

            reobj = re.compile('[1-9]')

            lowercse = string.ascii_lowercase

            uppercse = string.ascii_uppercase

            nums = string.digits

            symbs = '@#*!'

            combine = uppercse + symbs + nums + lowercse

            temp = random.sample(combine,12)

            random.shuffle(temp)

            passwd = "".join(temp)

            if passwd.find('@') > 0 or passwd.find('!') > 0 or passwd.find('#') > 0 or passwd.find('*') > 0:

                find_num = reobj.search(passwd)

                if find_num != None :

                    return passwd

                    break

                else:

                    continue

            else:

                continue

    

    def send_email(self,mailbox,o_passwd,usr_passwd,to_email,subj,caption,instance_url):

        msg = EmailMessage()

        msg['From'] = mailbox

        msg['To'] = to_email

        msg['Subject'] = subj

        msg.set_content("id: {1} \n New Password: {2} \n Go to:{3} to login".format(to_email,usr_passwd,instance_url))

        msg.add_alternative(f"""

        <table style="border: 1px solid #ddd;">

            <tr>

                <th style="padding: 1rem; background-color: rgb(0, 132, 255); text-align: center;">

                    <h2 aria-details="{caption}" style="color: white; font-family: Arial, Helvetica, sans-serif;">

                        Your Summit Account has been unlocked!

                    </h2>

                </th>

            </tr>

            <tr>

                <td style="font-family: Arial, Helvetica, sans-serif; padding: 8px;">Please use the following credentials to login:</td>

            </tr>

            <tr>

                <td style="font-family: Arial, Helvetica, sans-serif; padding: 8px;">

                    ID: <span style="color: rgba(0, 132, 255, 0.952);">{to_email}</span> <br>

                    New Password: 

                    <span style="color: rgba(0, 132, 255, 0.952); font-family: Arial, Helvetica, sans-serif;">{usr_passwd}</span>

                </td>

            </tr>

            <tr>

                <td style="font-family: Arial, Helvetica, sans-serif; margin: 5px; border: 1px solid #ddd; padding: 8px; background-color: #ddd;">

                    Click <a href="{instance_url}">here</a> to Login into Summit Or, simply go to: <br>

                    <span style="color: rgba(0, 132, 255, 0.952); font-family: Arial, Helvetica, sans-serif;"> {instance_url} </span> to Login into Symphony Summit

                </td>

            </tr>

            <tr>

                <td style="font-family: Arial, Helvetica, sans-serif; margin: 5px; border: 1px solid #ddd; padding: 4px; background-color: rgb(129, 123, 123); color: #ddd;"> Your Ticket will now be closed</td>

            </tr>

        </table>

            """, subtype = 'html')

        try:

            with smtplib.SMTP('smtp-mail.outlook.com', 587) as smtp:

                smtp.ehlo()

                smtp.starttls()

                smtp.ehlo()

 

                smtp.login(mailbox, o_passwd)

                smtp.send_message(msg)

            return 0

        except Exception as err:

            return str(err)

 

    def login_sel_fn(self,inputs):

        idbox = self.driver.find_element_by_name('txtLogin')

        idbox.send_keys(inputs['bot_email_id'])

 

        passwordbox = self.driver.find_element_by_name('txtPassword')

        passwordbox.send_keys(inputs['bot_password'])

 

        loginbutton = self.driver.find_element_by_name('butSubmit')

        loginbutton.click()

        #Duplicate login issue contingency:

        try :

            if WebDriverWait(self.driver,5).until(EC.frame_to_be_available_and_switch_to_it((By.ID,'SPopUp-frame'))):

                self.driver.find_element_by_xpath('//*[@id="ContentPanel_btnContinue"]').click()

        except TimeoutException:

            print(TimeoutException)

 

    def nav_analyst(self):        

        adminoption = self.driver.find_element_by_xpath('//*[@id="divMenu"]/nav/ul/li[16]')

        adminoption.click()

 

        analystoption = self.driver.find_element_by_xpath('//*[@id="SUMMIT_SD_EXEC"]')

        analystoption.click()

 

        tenantoption = self.driver.find_element_by_xpath('//*[@id="BodyContentPlaceHolder_UCInstanceList_lstInstance"]/option[2]')

        #tenantoption = self.driver.find_element_by_name('ctl00$BodyContentPlaceHolder$UCInstanceList$lstInstance')

        tenantoption.click()

    

    def workgroup_alloc(self,workgroup,email):

        #Select workgroup

        try:

                

            ddlwgelement = self.driver.find_element_by_name('ctl00$BodyContentPlaceHolder$lstWorkGroup')

            ddlwg = Select(ddlwgelement)

            ddlwg.select_by_visible_text(workgroup)

 

            self.driver.implicitly_wait(5)

 

            #Select Enterprise

            ddlenterpriseelement = self.driver.find_element_by_name('ctl00$BodyContentPlaceHolder$ddlEditionTypeByWorkgroup')

            ddlenterprise = Select(ddlenterpriseelement)

            ddlenterprise.select_by_visible_text('Enterprise')

 

            #Select Attribute License

            ddlattrelement = self.driver.find_element_by_name('ctl00$BodyContentPlaceHolder$ddlAttributeTypeByWorkgroup')

            ddlattr = Select(ddlattrelement)

            ddlattr.select_by_visible_text('Concurrent Analyst')

            #enter analyst name and select

            analystname = self.driver.find_element_by_name('ctl00$BodyContentPlaceHolder$txtExecutive')

            analystname.send_keys(email)

            selanalyst = self.driver.find_element_by_class_name('ui-menu-item')

            selanalyst.click()

 

            time.sleep(2)

            #submit or confirmation pop-up

            self.driver.find_element_by_name('ctl00$BodyContentPlaceHolder$btnSave').click()

 

            #Dialogsaved: Saved successfully dialogue Dialogconfirm: Terminate user sessions confirmation Dialogue

            dialogsaved = self.driver.find_elements_by_xpath('/html/body/div[8]/div/div/div[2]/button')

            dialogconfirm = self.driver.find_elements_by_xpath('/html/body/div[10]/div/div/div[3]/button[2]')

            time.sleep(2)

            

            if len(dialogconfirm) > 0:

                dialogconfirm[0].click()

                dialogsaved = self.driver.find_elements_by_xpath('/html/body/div[8]/div/div/div[2]/button')

                if len(dialogsaved) > 0:

                    return "Success"

                else:

                    err_message = WebDriverWait(self.driver,5).until(EC.visibility_of_element_located((By.CLASS_NAME,'noty_text'))).text

                    return (str(err_message))

            elif len(dialogsaved) > 0:

                dialogsaved[0].click()

                return ("Success")

 

        except Exception as err:

            return(str(err))

 

    def search_usr_lst(self,email):

        adminoption = self.driver.find_element_by_xpath('//*[@id="divMenu"]/nav/ul/li[16]')

        adminoption.click()

 

        usrlstoption = self.driver.find_element_by_id('SUMMIT_USER_LIST')

        usrlstoption.click()

 

        filteropt = self.driver.find_element_by_id('filter')

        filteropt.click()

        

        usrsearch = self.driver.find_element_by_id('BodyContentPlaceHolder_txtSearchUser')

        usrsearch.send_keys(email)

        usrsearchbtn = self.driver.find_element_by_id('BodyContentPlaceHolder_btnFilter')

        usrsearchbtn.click()

        

    def give_req_access(self,req_access_level):

 

        usrclick = self.driver.find_element_by_xpath('//*[@id="BodyContentPlaceHolder_gvMasterList"]/tbody/tr[2]/td[1]/a')

        usrclick.click()

        

        usraccesstab = self.driver.find_element_by_id('ACCESS')

        usraccesstab.click()

        

        tableid = self.driver.find_elements_by_class_name('mstGV')[1]

        tbody = tableid.find_element_by_tag_name('tbody')

        rows = tbody.find_elements_by_tag_name('tr')

 

        try:

                

            for row in rows:

                #print("Row: ",row.text,"\n")

                

                if row.text == req_access_level:

 

                    cols = row.find_elements_by_tag_name('td')

                    count = 0

 

                    for col in cols:

                        count = count+1

                        

                        if count == 3:

                            #print("Coll:",col.text,"\n")

                            checkboxopt = col.find_element_by_tag_name('input')

                            

                            if checkboxopt.is_selected() == False:

                                checkboxopt.click()

                            else:

                                #print("Continue")

                                continue

                else:

                    pass

 

        except Exception as err:

            return (str(err))

 

        saveusrdetails = self.driver.find_element_by_id('BodyContentPlaceHolder_btnSave')

        saveusrdetails.click()

        saveconfirmbtn = self.driver.find_element_by_xpath('/html/body/div[9]/div/div/div[2]/button')

        saveconfirmbtn.click()

 

    #Selenium Functions till Here

 

    #api request

    def post_req_fn(self,api_url,payload):

        try:

                

            post_reqst = requests.post(url=api_url, json=payload, verify=False)

            post_req_json = post_reqst.json()

            return post_req_json

        except Exception as err:

            err_dict = dict()

            err_dict ={'Errors': str(err)}

            print(err_dict)

            return err_dict

 

    #match role to catalog ddl with json file helper

    def load_key_val(self,fileptr):

        self.keyval = json.load(fileptr)

    

    def fetch_key_val(self,check_item):

        #print(self.keyval)

        for items in self.keyval['keys']:

            if check_item == items['catalog-ddl']:

                return(items['role-template'])

 

    #fetch catalog details

    def fetch_catalog_attr(self,item):

        catalog = dict()

        for inneritem in item['CustomAttribute']:

            if inneritem['SR_CtAttribute_Name'].find('Email') > 0 or inneritem['SR_CtAttribute_Name'].find('email') > 0:

                catalog['usr_email'] = inneritem['AttributeText']

            elif inneritem['SR_CtAttribute_Name'].find('workgroup') >= 0 or inneritem['SR_CtAttribute_Name'].find('Workgroup') >= 0 :

                catalog['workgroup'] = inneritem['AttributeText']

            elif inneritem['SR_CtAttribute_Name'].find('Access') >= 0 or inneritem['SR_CtAttribute_Name'].find('access') >= 0 :

                catalog['access'] = inneritem['AttributeText']

        return catalog

 

    #check API exists or incorrect API or not

    def chk_api_fn(self,api_url):

        try:

            reqst = requests.post(url=api_url, verify=False)

            return (reqst.status_code)

        except ConnectionError:

            return "Incorrect Api"

    

    def check_url_exists(self,url):

        try:

            reqs = requests.get(url, verify=False)

            return (reqs.status_code)

        except ConnectionError:

            return ("Invalid URL")

 

    def verify_inputs_fn(self,input_data):

        for item in input_data:

            if input_data[item] == '':

                return False

        return True

    

 