import os
os.environ["NODE_SKIP_PLATFORM_CHECK"] = "1" #為了防止新版本playwright無法運行於Win7
from playwright.sync_api import sync_playwright
# from playwright.async_api import async_playwright

import requests
from bs4 import BeautifulSoup

def is_notebook() -> bool:
    '''
    判斷是不是處於notebook環境，因為jupyter notebook和sync_playwright不相容，所以改成使用selenium
    '''
    try:
        shell = get_ipython().__class__.__name__
        if shell == 'ZMQInteractiveShell':
            return True   # Jupyter notebook or qtconsole
        elif shell == 'TerminalInteractiveShell':
            return False  # Terminal running IPython
        else:
            return False  # Other type (?)
    except NameError:
        return False      # Probably standard Python interpreter
    

if is_notebook():
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.wait import WebDriverWait


class Client:
    def __init__(self, login_id=None, login_psw=None, TEST_MODE=False):
        # 建立request session
        s = requests.session()
        headers = {
            'user-agent': 'Mozilla/5.0 (Linux; U; Android 2.3.3; zh-tw; HTC_Pyramid Build/GRI40) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari',
            'referer': 'http://web9.vghtpe.gov.tw/'
        }
        s.headers.update(headers)
        self.session = s  # 將建立的物件存在instance屬性中
        self.headers = headers
        self.TEST_MODE = TEST_MODE
        self.login_id = login_id
        self.login_psw = login_psw


    def acquire_id_psw(self):    
        while(1):
            login_id = input("Enter your ID: ").upper().strip()
            if len(login_id) != 0:
                break
        while(1):
            login_psw = input("Enter your PASSWORD: ")
            if len(login_psw) != 0:
                break
        return login_id, login_psw


    def eip_login_selenium(self, login_id=None, login_psw=None):
        '''
        使用selenium連線EIP，會判斷成功、需要更改密碼、帳密錯誤
        '''
        if login_id is None or login_psw is None:
            if self.login_id is None or self.login_psw is None:
                login_id, login_psw = self.acquire_id_psw()
            else:
                login_id = self.login_id
                login_psw = self.login_psw
        
        # Provide the path to the Microsoft Edge WebDriver executable
        # edge_driver_path = './msedgedriver' # FIXME 這需要設定嗎?
        edge_options = webdriver.EdgeOptions()
        if not self.TEST_MODE:
            edge_options.add_argument('--headless')  # Optional: Run in headless mode

        driver = webdriver.Edge(options=edge_options)

        # Set the user agent
        driver.execute_cdp_cmd(
            "Network.setUserAgentOverride",
            {
                "userAgent": self.headers['user-agent']
            },
        )

        # Create a new page
        driver.get("https://eip.vghtpe.gov.tw/login.php")

        login_id_element = driver.find_element(By.CSS_SELECTOR, '#login_name')
        login_id_element.send_keys(login_id)

        login_psw_element = driver.find_element(By.CSS_SELECTOR, '#password')
        login_psw_element.send_keys(login_psw)

        login_btn_element = driver.find_element(By.CSS_SELECTOR, '#loginBtn')
        login_btn_element.click()

        # Wait for the page to load
        WebDriverWait(driver, timeout=10).until(lambda d: d.find_element(By.CSS_SELECTOR,".username"))

        # Check the URL for successful login
        print(driver.current_url)
        if "module_page.php" in driver.current_url or "vghtpe_dashboard.php" in driver.current_url:
            print("EIP: Login succeeded!")
            cookies = driver.get_cookies()
            for cookie in cookies:
                self.session.cookies.update({cookie['name']: cookie['value']})
            self.login_id = login_id
            self.login_psw = login_psw
            self.webmode = 'selenium'
            self.webbrowser = driver
            return True
        elif "https://eip.vghtpe.gov.tw/login_check.php" in driver.current_url:
            try:
                driver.find_element(By.LINK_TEXT,"暫不變更").click()
                WebDriverWait(driver, timeout=10).until(lambda d: d.find_element(By.CSS_SELECTOR,".username"))
                print("EIP: Login succeeded!")
                cookies = driver.get_cookies()
                for cookie in cookies:
                    self.session.cookies.update({cookie['name']: cookie['value']})
                self.login_id = login_id
                self.login_psw = login_psw
                self.webmode = 'selenium'
                self.webbrowser = driver
                return True
            except:
                print("EIP: Login failed! (need new PSW)")
                self.login_id = None
                self.login_psw = None
                return False
        elif "https://eip.vghtpe.gov.tw/login.php" in driver.current_url:
            print("EIP: Login failed! (wrong ID/PSW)")
            self.login_id = None
            self.login_psw = None
            return False
        else:
            print("EIP: Login failed! (others)")
            print(driver.current_url)
            self.login_id = None
            self.login_psw = None
            return False          

    
    def eip_login_playwright(self, login_id=None, login_psw=None):
        '''
        使用playwright連線EIP，會判斷成功、需要更改密碼、帳密錯誤
        '''
        if login_id is None or login_psw is None:
            if self.login_id is None or self.login_psw is None:
                login_id, login_psw = self.acquire_id_psw()
            else:
                login_id = self.login_id
                login_psw = self.login_psw
        
        with sync_playwright() as p:  
            if self.TEST_MODE:
                browser =p.chromium.launch(headless=False, channel="msedge") # channel="msedge"，使用該電腦原生的browser
            else:
                browser =p.chromium.launch(headless=True, channel="msedge") # channel="msedge"，使用該電腦原生的browser    

		    #設定user_agent
            context =browser.new_context(
                user_agent=self.headers['user-agent'],
                ignore_https_errors=True,
            )
            context.set_default_timeout(10000) #設定全域timeout為10秒 #TODO 這有效果嗎?
            page = context.new_page()
            
            page.goto("https://eip.vghtpe.gov.tw/login.php")
            page.locator('#login_name').fill(login_id)
            page.locator('#password').fill(login_psw)
            page.locator('#loginBtn').click()
            page.wait_for_load_state("networkidle")
            if "module_page.php" in page.url or "vghtpe_dashboard.php" in page.url:
                print("EIP: Login succeeded!")
                cookies_list = context.cookies()
                for i in cookies_list:
                    self.session.cookies.update({i['name']:i['value']})
                self.login_id = login_id
                self.login_psw = login_psw
                self.webmode = 'playwright'
                self.webbrowser = page
                return True
            elif "https://eip.vghtpe.gov.tw/login_check.php" in page.url:
                try:
                    # page.get_by_role('link', name="暫不變更")
                    # with page.expect_navigation():
                    #     page.get_by_text("暫不變更").click()
                    page.get_by_text("暫不變更").click()
                    page.wait_for_load_state("networkidle")
                    print("EIP: Login succeeded!")
                    cookies_list = context.cookies()
                    for i in cookies_list:
                        self.session.cookies.update({i['name']:i['value']})
                    self.login_id = login_id
                    self.login_psw = login_psw
                    self.webmode = 'playwright'
                    self.webbrowser = page
                    return True  
                except:
                    print("EIP: Login failed!(need new PSW)")
                    self.login_id = None
                    self.login_psw = None
                    return False
            elif "https://eip.vghtpe.gov.tw/login.php" in page.url:
                print("EIP: Login failed!(wrong ID/PSW)")
                self.login_id = None
                self.login_psw = None
                return False
            else:
                print("EIP: Login failed!(others)")
                print(page.url)
                self.login_id = None
                self.login_psw = None
                return False          


    def eip_login_webbrowser(self, login_id=None, login_psw=None):
        '''
        整合使用selenium和playwright登入EIP
        '''
        if is_notebook():
            res = self.eip_login_selenium(login_id=login_id, login_psw=login_psw)
        else:
            res = self.eip_login_playwright(login_id=login_id, login_psw=login_psw)
        return res


    def eip_app(self, app_name=None):
        # TODO
        '''
        未來應考慮可能要從web9_app遷移到eip_app
        '''
        pass


    def web9_login_requests(self, login_id=None, login_psw=None):
        '''
        連線web9系統
        '''
        if login_id is None or login_psw is None:
            if self.login_id is None or self.login_psw is None:
                login_id, login_psw = self.acquire_id_psw()
            else:
                login_id = self.login_id
                login_psw = self.login_psw
        
        tURL = "https://web9.vghtpe.gov.tw/Signon/lockaccount"
        login_payload = {
            "j_username": login_id,
            "j_password": login_psw,
            "Submit": "確認登入",
            "j_pin": "",
            "j_pin2": ""
        }
        r = self.session.post(tURL, data=login_payload)

        # 取得myFunctions的內容 => 不同的帳號相同的app name會有不同的seqNo
        r = self.session.get("https://web9.vghtpe.gov.tw/Signon/myFunctions.jsp")
        self.soup_myfunction = BeautifulSoup(r.text, "lxml")
        target = self.soup_myfunction.select("title")[0]

        if target.string.find('[Signon Main Function Screen]') == -1:  # 登入成功的判斷方式
            print("WEB9: Login failed!")
            return False
        else:
            print("WEB9: Login succeeded!")
            return True


    def web9_app_requests(self, app_name=None):
        '''
        web9登入後各服務的選擇
        '''
        # 列出parse出的所有app name
        self.app_dict = dict()
        app_list = self.soup_myfunction.select('input[name="FunBn"]')  # 把共同主螢幕的每個app列出來
        for i in app_list:
            t = i['onclick'].replace("VupFunc(", "").replace(")", "").replace('"', '').replace(",",
                                                                                               "").split()  # 找出他們的重要資訊
            if len(t) >= 3:  # 為了防止'PACS'沒有連結 => 需要再研究
                item, index, link = t
            # self.soup.select('#hideVuoperId')[0]['value']
            self.app_dict[item] = dict(index=index, link=link)

        # 如果沒有選擇app名稱就印出全部讓使用者選
        if not app_name:
            print(list(self.app_dict.keys()))
            self.app_name = input("\nEnter the app name: ").upper()  # 轉成大寫格式
        else:
            self.app_name = app_name

        # 啟動特定app
        selected_app = self.app_dict.get(self.app_name)
        if selected_app is None:
            print(f'APP: {self.app_name} is not found!')
            return None
        else:
            if selected_app['link'][0] != '/':  # TODO 之後處理網址不是相對路徑的
                print("Not a relative URL path")
                return None
            payload = {
                'srnId': self.app_name,
                'seqNo': selected_app['index']
            }
            try:
                r = self.session.get('https://web9.vghtpe.gov.tw' + selected_app['link'],
                                    params=payload)  # 取得每個app的特殊cookies或是SMSGEN需要的authorization
                print(f"WEB9: {app_name} login succeeded!")
                return r  # 返回進入app的response，讓後續程式操作
            except:
                print(f"WEB9: {app_name} login failed!")
                return False


    def scheduler_login(self, login_id=None, login_psw=None):
        '''
        登入排程系統
        '''
        if login_id is None or login_psw is None:
            if self.login_id is None or self.login_psw is None:
                login_id, login_psw = self.acquire_id_psw()
            else:
                login_id = self.login_id
                login_psw = self.login_psw
        
        tURL = "http://10.97.235.122/Exm/HISLogin/CheckUserByID"
        login_payload = {
            'signOnID': login_id,
            'signOnPassword': login_psw
        }
        r = self.session.post(tURL, data=login_payload)
        if r.status_code == 200:
            print("SCHEDULER: Login succeeded!")
            self.login_id = login_id
            self.login_psw = login_psw
            return True
        else:
            print("SCHEDULER: Login failed!")
            self.login_id = None
            self.login_psw = None
            return False
        


    def login_drweb(self):
        '''
        完成登入EIP+WEB9
        '''
        while True:
            res = self.eip_login_webbrowser()
            if res:
                break
        self.web9_login_requests(login_id=self.login_id, login_psw=self.login_psw)
        self.web9_app_requests('DRWEBAPP')
    

    # TODO
    def note_surgery_web(self,):
        if self.webmode == 'selenium':
            self.webbrowser.get('URL')
        elif self.webmode == 'playwright':
            self.webbrowser.goto('URL')

    # TODO
    def note_admission_web(self,):
        pass
    
    # TODO
    def note_discharge_web(self,):
        pass
    
    # TODO
    def note_progress_web(self,):
        pass


if __name__=='__main__':
    c=Client()
    c.TEST_MODE = True
    c.login_drweb()