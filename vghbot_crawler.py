import vghbot_login
import pandas as pd
import time, datetime
from bs4 import BeautifulSoup
from pathlib import Path
import urllib.parse

class VghCrawler(vghbot_login.Client):
    def __init__(self, login_id = None, login_psw = None, TEST_MODE=True):
        super().__init__(login_id=login_id, login_psw=login_psw, TEST_MODE=TEST_MODE)
        self.login_drweb()
    

    def patient_search(self, name='', drid='', hisno='', id='', ward='0'):
        '''
        用病歷號/身分證號/姓名/門診就診科別去找病人     
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPatient'
        payload = {
            'wd': ward,
            'histno': hisno,
            'pidno': id,
            'namec': name,
            'drid': drid,
            'er': '0',
            'bilqrta': '0', 
            'bilqrtdt': '',
            'bildurdt': '0',
            'other': '0',
            'nametype': ''
        }
        response = self.session.post(url, data=payload)
        table = pd.read_html(response.text, attrs={'id':'patlist'}, flavor='lxml')[0]
        return table # TODO 如果沒有找到會回傳甚麼?


    def patient_info(self, hisno):
        '''
        病人:用病歷號取得病人資訊
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findPba',
            'histno': str(hisno),
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        table = pd.read_html(response.text)[0]
        # TODO parsing 出來會是series嗎?


    def salary(self, year, month):
        '''
        醫師:薪水資料
        '''
        # Step 1: GET request
        url = 'https://web9.vghtpe.gov.tw/psrp/index.jsp'
        response = self.session.get(url)

        # Step 2: Extract values from hidden input tags using BeautifulSoup
        rimkey = None
        idno = None

        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            rimkey_input = soup.find('input', {'name': 'rimkey'})
            idno_input = soup.find('input', {'name': 'idno'})

            if rimkey_input:
                rimkey = rimkey_input['value']
            
            if idno_input:
                idno = idno_input['value']

        # Step 3: POST request with payload (unchanged)
        if rimkey is not None and idno is not None:
            payload = {
                'rimkey': rimkey,
                'idno': idno,
                'year': year,
                'month': month
            }
            post_url = 'https://web9.vghtpe.gov.tw/psrp/genpdf1'
            response = self.session.post(post_url, data=payload)

            # Step 4: Analyze response headers (unchanged)
            directory_name = f'salary_{self.login_id}'
            filename = f'{idno}_{year}_{month}.pdf'
            file = Path(directory_name, filename)
            file.parent.mkdir(parents=True, exist_ok=True)
            if response.status_code == 200:
                content_type = response.headers.get('Content-Type')
                if content_type == 'application/pdf':
                    # Download the file (unchanged)
                    with file.open('wb') as f:
                        f.write(response.content)
                    print(f"File downloaded: {filename}")
                elif 'text/plain' in content_type:
                    print(f"Content-Type is text/plain:{response.text}")
                else:
                    print("Unknown Content-Type")
            else:
                print("POST request failed.")
        else:
            print("Values not found in the first GET response.")


    def opd_patient_list_previous(self, date:str|list[str], to_now = False) -> list[pd.DataFrame]:
        '''
        醫師:過去看診名單
        '''
        # TODO 實作從指定日期以後的清單
        if type(date) == str and to_now == False:
            baseURL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm"
            payload = {
                'action':'findOpdRotQ8',
                'dtpdate': date,
                '_': int(time.time()*1000)
            }
            res = self.session.get(baseURL, params=payload)
            table_list = pd.read_html(res.text, parse_dates=['門診日期'], flavor='lxml')
            if len(table_list[0]) == 1 and table_list[0].iloc[0,0] == '看診清單':
                return None
            return {date:table_list[0]}
        elif type(date) == str and to_now == True:
            table_dict = {}
            yesterday = datetime.datetime.today() - datetime.timedelta(days=1)
            yesterday_string = yesterday.strftime("%Y%m%d")
            if date > yesterday_string:
                print("超過昨日日期")
                return None
            while(1):
                table_dict[yesterday_string] = self.opd_patient_list_previous(date=yesterday_string)[yesterday_string]
                yesterday = yesterday - datetime.timedelta(days=1)
                yesterday_string = yesterday.strftime("%Y%m%d")
                if yesterday_string < date:
                    break
            return table_dict
        elif type(date) == list:
            table_dict = {}
            for date_string in date:
                table_dict[date_string] = self.opd_patient_list_previous(date=date_string)[date_string]
            return table_dict


    def opd_patient_list_appointment(self, date:str = None):
        '''
        醫師:未來掛號名單
        date: 若為None印出後續每天預約掛號人數，若有指定date則是該天的掛號病人清單
        # TODO: 未來需要做一個時間區間的掛號名單嗎?
        '''
        if date is None:
            baseURL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm"
            payload = {
                'action':'findReg',
                '_': int(time.time()*1000)
            }
            res = self.session.get(baseURL, params=payload)
            table = pd.read_html(res.text, parse_dates=['掛號日期'], attrs={'id':'reglist'}, flavor='lxml')[0]
            return table
        else:
            table_list = [] # TODO 考慮轉成datafram將每個dataframe合併
            baseURL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm"
            payload = {
                'action':'findReg',
                '_': int(time.time()*1000)
            }
            res = self.session.get(baseURL, params=payload)
            table = pd.read_html(res.text, attrs={'id':'reglist'})[0]
            table_dict = table[table['掛號日期'] == date].loc[:,['掛號日期','科代碼','診間代碼']].to_dict('records') 
            for row in table_dict:
                payload2 = {
                    'action':'findReg',
                    'dt': row['掛號日期'], 
                    'ect': row['科代碼'],
                    'room': row['診間代碼'],
                    '_': int(time.time()*1000)
                }
                res = self.session.get(baseURL, params=payload2)
                table = pd.read_html(res.text, attrs={'id':'regdetail'}, flavor='lxml')[0]
                table_list.append(table)
            return table_list


    def opd_list(self, hisno):
        '''
        病人:門診就診清單(門診+>4門診都抓並合併)
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findOpd',
            'histno': str(hisno),
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        table = pd.read_html(response.text, attrs={'id':'opdlist'}, flavor='lxml')[0]
        if '無門診' in table.iloc[0,0]:
            table = None


        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findOpd01',
            'histno': str(hisno),
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        table2 = pd.read_html(response.text, attrs={'id':'opdlist01'}, flavor='lxml')[0]
        if '無門診' in table2.iloc[0,0]:
            table2 = None

        if table is not None and table2 is not None:
            return table.append(table2)
        elif table is None:
            return table2
        elif table2 is None:
            return table
        else:
            return None
        
        # TODO 尚未測試


    def opd_note(self, hisno, date, doctor, department):
        '''
        病人:依照篩選條件(doctor, department, 診次?, 時間) 取得特定病人門診病歷，回覆劃分成SOAP分類dict
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findOpd',
            'histno': str(hisno),
            'dt': date,
            'dept': 103, # FIXME
            'doc': urllib.parse.quote(doctor, encoding='big5'),
            'deptnm': urllib.parse.quote(department, encoding='big5'),
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        text = response.text()
        soup = BeautifulSoup(text, "html.parser")
        pre_tags = soup.find_all('pre')
        
        # Find the section containing the table after "用藥記錄"
        section = soup.find('legend', text='[用藥記錄]').find_parent('fieldset')
        drugs = pd.read_html(section.text, flavor='lxml')

        # Find the section containing the table after "用藥記錄"
        section = soup.find('legend', text='[門診醫囑]').find_parent('fieldset')
        orders = pd.read_html(section.text, flavor='lxml')

        data = {
            'S':pre_tags[0].get_text(),
            'O':pre_tags[1].get_text(),
            'A':'', # TODO 要怎麼parse??
            'P':pre_tags[2].get_text(),
            'drugs': drugs,
            'orders': orders,
        }

        return data


    def op_list(self, hisno):
        '''
        病人:手術清單
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findOpn',
            'histno': str(hisno),
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        table = pd.read_html(response.text, attrs={'id':'opnlist'}, flavor='lxml')[0]
        return table
    

    def op_note(self, hisno, date):
        '''
        病人:手術note
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findOpn',
            'histno': str(hisno),
            'dt': date, # TODO 需要從西元紀年改成民國紀年
            'tm': 1,
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        text = response.text()
        return text


    def ad_list(self, hisno):
        '''
        病人:住院清單
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findAdm',
            'histno': str(hisno),
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        table = pd.read_html(response.text, attrs={'id':'admlist'}, flavor='lxml')[0]
        return table
    

    def ad_note(self, hisno, caseno, adidate):
        '''
        病人:住院note
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findAdm',
            'histno': str(hisno),
            'caseno': str(caseno),
            'adidate': adidate, # TODO 需要從西元紀年改成民國紀年
            'tm': 1,
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        text = response.text()
        return text


    def drug_list(self, hisno):
        '''
        病人:藥物清單
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findUd',
            'histno': str(hisno),
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        table = pd.read_html(response.text, attrs={'id':'caselist'}, flavor='lxml')[0]
        return table


    def drug_content(self, hisno, caseno, dt, type, dept, dt1):
        '''
        病人:藥物內容
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findUd',
            'histno': str(hisno),
            'caseno': str(caseno),
            'dt': dt, # TODO 需要從西元紀年改成民國紀年
            'type': type,
            'dept': dept,
            'dt1': dt1,
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        table = pd.read_html(response.text, attrs={'id':'udorder'}, flavor='lxml')[0]
        return table


    def scaned_note(self):
        '''
        病人:掃描清單
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findScan',
            'tdept': 'OPH',
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        table = pd.read_html(response.text, flavor='lxml')[0]
        return table
    

    def consult_list(self, hisno):
        '''
        病人:會診清單
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findCps',
            'histno': str(hisno),
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        table = pd.read_html(response.text, attrs={'id':'cpslist'}, flavor='lxml')[0]
        return table


    def consult_note(self, hisno, caseno, oseq):
        '''
        病人:會診內容
        '''
        url = 'https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm'
        payload = {
            'action': 'findCps',
            'histno': str(hisno),
            'caseno': str(caseno),
            'oseq': str(oseq),
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload)
        text = response.text()
        return text

    # def opd_search(self):
    #     '''
    #     搜尋過去住院內容
    #     '''
    #     pass

    # def ad_search(self):
    #     '''
    #     搜尋過去住院內容
    #     '''
    #     pass

    # def op_search(self):
    #     '''
    #     搜尋過去住院內容
    #     '''
    #     pass

    # def drug_search(self, hisno, drug_name):
    #     '''
    #     搜尋過去用藥
    #     '''
    #     pass

    # def consult_search(self):
    #     '''
    #     搜尋過去會診內容
    #     '''
    #     pass

    # def certification(self, hisno, print=False):
    #     '''
    #     暫存/列印診斷證明
    #     '''
    #     pass

    # def patient_info_for_verification(self, hisno):
    #     '''
    #     病人:用病歷號取得病人事審資訊
    #     '''
    #     pass

    def op_schedule_list_doc(self, date, doc):
        '''
        回傳手術排程，透過日期(民國格式:1120702)+醫師燈號(4102)
        '''
        url = 'https://web9.vghtpe.gov.tw/ops/opb.cfm'
        payload_doc = {
            'action': 'findOpblist',
            'type': 'opbmain',
            'qry': doc, # '4102',
            'bgndt': date, # '1120703',
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload_doc)
        df = pd.read_html(response.text, flavor='lxml')[0]
        df = df.astype('string')
        soup = BeautifulSoup(response.text, "html.parser")
        link_list = soup.find_all('button', attrs={'data-target':"#myModal"})
        df['link'] = [l['data-url'] for l in link_list]
        return df

    def op_schedule_list_section(self, date, section):
        '''
        回傳手術排程，透過日期(民國格式:1120702)+部門(OPH)
        '''
        url = 'https://web9.vghtpe.gov.tw/ops/opb.cfm'
        payload_sect = {
            'action': 'findOpblist',
            'type': 'opbsect',
            'qry': section, # 'oph'
            'bgndt': date, # '1120702'
            '_': int(time.time()*1000)
        }
        response = self.session.get(url, params=payload_sect)
        df = pd.read_html(response.text, flavor='lxml')[0]
        df = df.astype('string')
        soup = BeautifulSoup(response.text, "html.parser")
        link_list = soup.find_all('button', attrs={'data-target':"#myModal"})
        df['link'] = [l['data-url'] for l in link_list]
        return df
    
    def op_schedule_detail(self, schedule_df: pd.DataFrame, hisno: str):
        df_dict = schedule_df.loc[ (schedule_df.loc[:,'病歷號']==hisno), ['病歷號', '姓名', '手術日期', '手術時間', '病歷號', 'link']].to_dict('records')[0]
        name =  df_dict['姓名']
        op_date = df_dict['手術日期']
        op_time = df_dict['手術時間']
        link_url = df_dict['link']

        base_url = 'https://web9.vghtpe.gov.tw'
        response = self.session.get(base_url+link_url)
        soup = BeautifulSoup(response.text, "html.parser")

        side = soup.select_one('table > tbody > tr:nth-child(12) > td:nth-child(2)').string # TODO 改成部位:的下一個sibling?
        if side == '右側':
            side = 'R'
        elif side == '左側':
            side = 'L'
        elif side == '雙側':
            side = 'B'

        result = {
            'hisno': hisno,
            'name': name,
            'op_room': soup.select_one('table > tbody > tr:nth-child(4) > td:nth-child(6)').string,
            'op_date': op_date,
            'op_time': op_time,
            'op_sect': soup.select_one('#OPBSECT')['value'].strip(),
            'op_bed': soup.select_one('table > tbody > tr:nth-child(1) > td:nth-child(6)').string.strip(' -'), # TODO 改成病房床號:的下一個sibling?
            'op_anesthesia': soup.select_one('#opbantyp')['value'], 
            'op_side': side,  
        }

        return result
        
    
    # def op_schedule_patient(self):
    #     pass

    # def download_scanned_note():
    #     '''
    #     下載掃描病歷
    #     '''
    #     pass

    # def upload_scanned_note():
    #     '''
    #     上傳掃描病歷
    #     '''
    #     pass

if __name__=='__main__':
    c = VghCrawler('DOC4123J','S00000000')