from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP

import pandas as pd
r = HTTP()


r.download(url="https://www.rpachallenge.com/assets/downloadFiles/challenge.xlsx",target_file="output/")
df = pd.read_excel("output/challenge.xlsx")
df['Phone Number'] = df['Phone Number'].astype(str)
browser = Selenium(auto_close=False)


def rpa():
    print('Bot started')
    browser.open_chrome_browser("https://www.rpachallenge.com/")
    
    try:
        for i in range(len(df)):
            firstname = df['First Name'][i]
            lastname = df['Last Name '][i]
            companyname = df['Company Name'][i]
            role = df['Role in Company'][i]
            address = df['Address'][i]
            email = df['Email'][i]
            phone = df['Phone Number'][i]
            browser.input_text(locator='//*[@ng-reflect-name="labelFirstName"]', text=firstname)
            browser.input_text(locator='//*[@ng-reflect-name="labelLastName"]', text=lastname)
            browser.input_text(locator='//*[@ng-reflect-name="labelCompanyName"]', text=companyname)
            browser.input_text(locator='//*[@ng-reflect-name="labelRole"]', text=role)
            browser.input_text(locator='//*[@ng-reflect-name="labelAddress"]', text=address)
            browser.input_text(locator='//*[@ng-reflect-name="labelEmail"]', text=email)
            browser.input_text(locator='//*[@ng-reflect-name="labelPhone"]', text=phone)
            browser.click_button(locator='//*[@value="Submit"]')
            

    finally:
        print('bot finished successfully')
        browser.close_window()
        pass


if __name__ == "__main__":
    rpa()
    

