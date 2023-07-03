from bs4 import BeautifulSoup
from string import Template
from datetime import datetime, timedelta
import time
import pandas
import re
import logging
import pprint
import warnings

# vghbot library
import gsheet
import vghbot_login
import updater_cmd

class OPNote():
    def __init__(self, webclient, config):
        ## 將session轉移進來
        self.session = webclient.session

        ## Input1: config設定檔，因為config內容是所有病人共用，所以安排在此處理，如果是個別病人的資訊調整會在fill_data內
        self.config = config

        ################ ################
        #這是為了新增VS_CODE和R_CODE在googlesheet欄位而做的hack，感覺應該重新設計才比較有結構
        #目前設計是當'BOT_MODE' == "IVI" 就會啟用google sheet取得內部的VS_CODE
        ################ ################
        self.config['R_NAME'] = get_name_from_code(self.config['R_CODE'], self.session)
        if self.config['BOT_MODE'] == "IVI": #處理特殊狀況
            print(f"IVI模式，VS與R登號會以google sheet設定")
        else: #正常狀況在這就取得VS、R的CODE和NAME，但可能需要重新設計會比較乾淨
            self.config['VS_NAME'] = get_name_from_code(self.config['VS_CODE'], self.session)
            print(f"目前使用手術Config: VS{self.config['VS_NAME']}({self.config['VS_CODE']}) || R{self.config['R_NAME']}({self.config['R_CODE']})")

        ## Input2: google sheet資料取得: 多人Dataframe
        self.gc = gsheet.GsheetClient()
        self.extracted_df = self.gc.get_df_select(self.config['SPREADSHEET'], self.config['WORKSHEET']) # 取得資料並完成選擇

        # 創建一個存放取得的op_schedule
        self.op_schedule_df = None

        size_of_df = len(self.extracted_df)
        if size_of_df == 0:
            print("完成Dataframe擷取: 無匹配資料")
            return False
        else:
            print(f"完成Dataframe擷取: {size_of_df}筆資料")
            self.data = {}  # 存所有的資料 = post_data(patient_info) + google sheet + config資料
            '''
            self.data[hisno] = {
                'data_web9op': data_web9op,
                'data_gsheet': data_gsheet,
                'post_data': post_data
            }
            '''
            
            self.date = check_opdate() # 取得時間
            
            return True


    def start(self):
        num = 0
        # 針對每一人的資料處理歸檔(轉成dictionary)
        for hisno in self.extracted_df[self.config['COL_HISNO']]:  # 使用病歷號作為識別資料
            if hisno.isnumeric(): # 跳過病歷號不全為數字的row + 空白row
                data_web9op = self.get_data_web9op(hisno)  # 取得特定病人web9基本資料
                data_opschedule = self.get_data_opschedule(hisno) # 取得特定病人手術排程資料
                data_gsheet = self.get_data_gsheet(hisno)  # 取得特定病人google表單資料

                post_data = self.fill_data(**{
                    'hisno':hisno, 
                    'num':num, 
                    'data_web9op': data_web9op,
                    'data_gsheet': data_gsheet,
                    'data_opschedule': data_opschedule
                })  # 不同的Note會有客製化填資料 
                
                self.data[hisno] = { # TODO 未來可以檢討有沒有需要保留這麼多資訊(除了post_data以外)
                    'data_web9op': data_web9op,
                    'data_gsheet': data_gsheet,
                    'data_opschedule': data_opschedule,
                    'post_data': post_data
                }
                num = num + 1

        if self.recheck_print():
            for hisno in self.data.keys():
                res = self.post(self.data[hisno]["post_data"])
                logger.info(f"新增記錄: {hisno} || {self.data[hisno]['data_web9op']['name']}")
        logger.info(f"完成本次自動化作業，共{num}筆資料")


    def post(self, data):  # 送出資料
        if TEST_MODE == True:
            target_url = "https://httpbin.org/post"
        else:
            target_url = "https://web9.vghtpe.gov.tw/emr/OPAController"
        response = self.session.post(target_url, data=data)
        if TEST_MODE == True:
            pprint.pprint(response.text)
        return response


    def get_data_gsheet(self, hisno):
        '''
        取得指定hisno的google sheet該列資料
        config中會以COL開頭的欄位名標示醫師刀表中對應的欄位，本函數會將對應的欄位擷取存入dict
        '''
        config = self.config
        df = self.extracted_df.loc[self.extracted_df[config['COL_HISNO']] == hisno, :]  # 取得對應hisno的該筆googlesheet資料(row)
        
        if len(df)>1: # 防止重複的hisno資料
            logger.error(f"指定範圍內有重覆的病歷號[{hisno}]，只會取得最上面的資料列")
            df = df.iloc[[0],:] 
                
        df_col = {}  # 依據有特別標記的col_name(以COL為開頭)去取得df該行的資料
        for key in config.keys():
            if key[:3].upper() == 'COL':
                if config[key].strip()!='' and config[key] in df.columns: # 要先確定該config欄位內是有資料的，有資料才去找，沒資料就跳過
                    df_col[key] = df.iloc[0].at[config[key]]
                else:
                    df_col[key] = None  # 如果config內有此變數但google sheet上沒有對應的column

        return df_col


    def get_data_web9op(self, hisno):
        '''
        抓取病人基本資料和處理手術時間問題(google sheet和手術護理時間的匹配)
        Input: hisno
        Output: dict()
        '''
        target_url = "https://web9.vghtpe.gov.tw/emr/OPAController?action=NewOpnForm01Action"
        payload = {
            "hisno": hisno,
            "pidno": "",
            "drid": "",
            "b1": "新增"
        }
        r_new = self.session.post(target_url, data=payload)
        soup_r_new = BeautifulSoup(r_new.text, "html.parser")
        pt_name = soup_r_new.find(attrs={"type": "hidden", "name": "name"}).get('value')
        # 擷取護理紀錄時間
        sel_opck = soup_r_new.select_one("select#sel_opck > option").get('value')  # 擷取護理師記錄
        if self.config['BOT_MODE'] != "IVI":  # IVI的note不用擷取護理紀錄時間
            if len(sel_opck.strip()) == 0 or sel_opck.strip() == "0" or (
                    sel_opck.split('|')[0][-11:-4] != sel_opck.split('|')[1][-11:-4]) or (
                    sel_opck.split('|')[1][-11:-4] != self.date):  # 沒有護理紀錄或是和gsheet時間不同就給使用者input??
                logger.error(f"病人({pt_name}):目前預定日期與護理紀錄不相同 或 沒有護理紀錄時間\n目前預定日期:{self.date}，時間請再手動輸入: \n")
                bgntm = input(f"病人({pt_name})手術開始時間(ex:0830): ")
                endtm = input(f"病人({pt_name})手術結束時間(ex:0830): ")
            else:
                bgntm = sel_opck.split('|')[0][-4:]
                endtm = sel_opck.split('|')[1][-4:]
        else:
            bgntm = endtm = ""  # 先給個預設的，IVI的note fill_data()會修正
        patient_info = {
            "sect1": soup_r_new.find(attrs={"type": "hidden", "name": "sect1"}).get('value'),
            "name": pt_name, # 病人姓名可以讓後續recheck呈現時使用
            "sex": soup_r_new.find(attrs={"type": "hidden", "name": "sex"}).get('value'),
            "hisno": soup_r_new.find(attrs={"type": "hidden", "name": "hisno"}).get('value'),
            "age": soup_r_new.find(attrs={"type": "hidden", "name": "age"}).get('value'),
            "idno": soup_r_new.find(attrs={"type": "hidden", "name": "idno"}).get('value'),
            "birth": soup_r_new.find(attrs={"type": "hidden", "name": "birth"}).get('value'),
            "_antyp": soup_r_new.find(attrs={"type": "hidden", "name": "_antyp"}).get('value'),
            "opbbgndt": soup_r_new.find(attrs={"type": "hidden", "name": "opbbgndt"}).get('value'),
            "opbbgntm": soup_r_new.find(attrs={"type": "hidden", "name": "opbbgntm"}).get('value'),
            "diagn": soup_r_new.find(attrs={"type": "text", "name": "diagn"}).get('value'),
            "sel_opck": sel_opck,
            "bgntm": bgntm,
            "endtm": endtm
        }
        return patient_info


    def get_data_opschedule(self, hisno):
        '''
        取得手術排程中的資訊: 住院與否、麻醉、IOL?
        '''
        if self.op_schedule_df is None:
            # 查詢op_schedule_df
            url = 'https://web9.vghtpe.gov.tw/ops/opb.cfm'
            payload_doc = {
                'action': 'findOpblist',
                'type': 'opbmain',
                'qry': self.config['VS_CODE'], # '4102',
                'bgndt': self.date, # '1120703',
                '_': int(time.time()*1000)
            }
            response = self.session.get(url, params=payload_doc)
            df = pandas.read_html(response.text)[0]
            df = df.astype('string')
            soup = BeautifulSoup(response.text, "html.parser")
            link_list = soup.find_all('button', attrs={'data-target':"#myModal"})
            df['link'] = [l['data-url'] for l in link_list]
            self.op_schedule_df = df

        df_dict = self.op_schedule_df.loc[ (self.op_schedule_df.loc[:,'病歷號']==hisno), ['病歷號', '姓名', '手術日期', '手術時間', '病歷號', 'link']].to_dict('records')[0]
        name =  df_dict['姓名']
        op_date = df_dict['手術日期']
        op_time = df_dict['手術時間']
        link_url = df_dict['link']

        base_url = 'https://web9.vghtpe.gov.tw'
        response = self.session.get(base_url+link_url)
        soup = BeautifulSoup(response.text, "html.parser")

        side = soup.select_one('table > tbody > tr:nth-child(12) > td:nth-child(2)').string # TODO 改成部位:的下一個sibling?
        if side == '右側':
            side = 'OD'
        elif side == '左側':
            side = 'OS'
        elif side == '雙側':
            side = 'OU'

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


    def fill_data(self, **kwargs):
        data_web9op = kwargs.get('data_web9op')
        data_gsheet = kwargs.get('data_gsheet')
        data_opschedule = kwargs.get('data_opschedule')

        # 填手術紀錄內容:
        post_data = {
            "sect1": data_web9op.get("sect1"),
            "name": data_web9op.get("name"),
            "sex": data_web9op.get("sex"),
            "hisno": data_web9op.get("hisno"),
            "age": data_web9op.get("age"),
            "idno": data_web9op.get("idno"),
            "birth": data_web9op.get("birth"),
            "_antyp": data_web9op.get("_antyp"),
            "opbbgndt": data_web9op.get("opbbgndt"),
            "opbbgntm": data_web9op.get("opbbgntm"),
            "opscode_num": 1, "film": "N", "against": "N", "action": "NewOpa01Action", "signchk": "Y",
            "sect": data_opschedule.get("op_sect"),
            "ward": data_opschedule.get("op_bed"),
            "source": "O",  # source是表示病人來自門診
            "again": "N", "reason_aga": "0", "mirr": "N", "saw": "N", "hurt": "1", "posn1": "1", "posn2": "0",
            "cler1": "2", "cler2": "0", "item1": "3", "item2": "2",
            ##以下都是空字串想測試全部拿掉##
            "ass2n": "", "ass3n": "", "ant1n": "", "ant2n": "", "dirn": "", "trtncti": "", "babym": "", "babyd": "",
            "bed": "",
            "final": "", "ass2": "", "ass3": "", "ant1": "", "ant2": "", "dir": "", "antyp1": "", "antyp2": "",
            "side": "", "oper": "", "rout": "", "side01": "", "oper01": "", "rout01": "",
            "opanam3": "", "opacod3": "", "opanam4": "", "opacod4": "", "opanam5": "", "opacod5": "",
            "opaicd2": "", "opaicdnm2": "", "opaicd3": "", "opaicdnm3": "",
            "opaicd4": "", "opaicdnm4": "", "opaicd5": "", "opaicdnm5": "",
            "opaicd6": "", "opaicdnm6": "", "opaicd7": "", "opaicdnm7": "", 
            "opaicd8": "", "opaicdnm8": "", "opaicd9": "", "opaicdnm9": "",
            ##以下需要變更##
            "man": "####", "ass1": "####", # 特殊處理
            "mann": "####", "ass1n": "####", # 特殊處理
            "bgndt": self.date, "enddt": self.date,
            "bgntm": data_web9op.get('bgntm'),  # 由護理紀錄時間
            "endtm": data_web9op.get('endtm'),  # 由護理紀錄時間
            "sel_opck": data_web9op.get("sel_opck"),
            "diagn": "##########",  # 特殊處理
            "diaga": "##########",  # 特殊處理
            "antyp": data_opschedule.get("op_anesthesia"),
            "opanam1": "", # 特殊處理
            "opacod1": "", # 特殊處理
            "opanam2": "", # 特殊處理
            "opacod2": "", # 特殊處理
            "opaicd0": "", # 特殊處理
            "opaicdnm0": "", # 特殊處理
            "opaicd1": "", # 特殊處理
            "opaicdnm1": "", # 特殊處理
            "op2data": "##########",  # 特殊處理: 病歷內文
        }

        #### 特殊處理
        # 如果VS_CODE存在就使用新的CODE
        if existandnotnone(data_gsheet, 'COL_VS_CODE'): 
            code = data_gsheet['COL_VS_CODE']
            post_data['man'] = code
            post_data['mann'] = get_name_from_code(code, self.session)
        else:
            post_data['man'] = self.config['VS_CODE']
            post_data['mann'] = self.config['VS_NAME']

        # 如果R_CODE存在就使用新的CODE
        if existandnotnone(data_gsheet, 'COL_R_CODE'):
            code = data_gsheet['COL_R_CODE']
            post_data['ass1'] = code
            post_data['ass1n'] = get_name_from_code(code, self.session)
        else:
            post_data['ass1'] = self.config['R_CODE']
            post_data['ass1n'] = self.config['R_NAME']

        # 判斷側別 => 判斷後存入新的變數 data_gsheet['OP_SIDE']
        if check_op_side(data_opschedule.get('op_side')) is not None: # 手術排程
            data_gsheet['OP_SIDE'] = check_op_side(data_opschedule.get('op_side'))
        elif check_op_side(data_gsheet.get('COL_OP')) is not None: # 刀表術式
            data_gsheet['OP_SIDE'] = check_op_side(data_gsheet.get('COL_OP'))
        elif check_op_side(data_web9op.get("diagn")) is not None: # web9術前診斷
            data_gsheet['OP_SIDE'] = check_op_side(data_web9op.get("diagn"))
        elif check_op_side(data_gsheet.get('COL_DIAGNOSIS')) is not None: # 刀表診斷
            data_gsheet['OP_SIDE'] = check_op_side(data_gsheet.get('COL_DIAGNOSIS'))
        elif check_op_side(data_gsheet.get('COL_SIDE')) is not None: # 刀表側別
            data_gsheet['OP_SIDE'] = check_op_side(data_gsheet.get('COL_SIDE'))
        else:
            print('異常: 無法決定側別')
            return False
        # 處理OD/OS/OU轉換成template內的Right/Left/Both
        data_gsheet['TRANSFORMED_SIDE'] = NOTE_TRANSFORM_SIDE.get(data_gsheet['OP_SIDE'])

        # 判斷手術的組套種類 => 後續依此決定診斷碼/術後診斷/病歷文本
        data_gsheet['OP_TYPE'] = check_op_type(data_gsheet.get('COL_OP'))
        if existandnotnone(data_gsheet, 'COL_LENSX'): # 如果有COL_LENSX欄位，且內部有資料就要換成'LENSX'
            data_gsheet['OP_TYPE'] = 'LENSX'

        # 處理診斷碼
        if data_gsheet['OP_TYPE'] == 'PHACO' or data_gsheet['OP_TYPE'] == 'ECCE' or data_gsheet['OP_TYPE'] == 'LENSX':
            if data_gsheet['OP_SIDE'] == "OD":
                post_data['opanam1'] = "PHACOEMULSIFICATION + PC-IOL IMPLANTATION"
                post_data['opacod1'] = "OPH 1342"
                post_data['opaicd0'] = "08RJ3JZ"
                post_data['opaicdnm0'] = "Replacement of Right Lens with Synthetic Substitute, Percutaneous Approach"
            elif data_gsheet['OP_SIDE'] == "OS":
                post_data['opanam1'] = "PHACOEMULSIFICATION + PC-IOL IMPLANTATION"
                post_data['opacod1'] = "OPH 1342"
                post_data['opaicd0'] = "08RK3JZ"
                post_data['opaicdnm0'] = "Replacement of Left Lens with Synthetic Substitute, Percutaneous Approach"
            elif data_gsheet['OP_SIDE'] == "OU":
                post_data['opanam1'] = "PHACOEMULSIFICATION + PC-IOL IMPLANTATION"
                post_data['opacod1'] = "OPH 1342"
                post_data['opanam2'] = "PHACOEMULSIFICATION + PC-IOL IMPLANTATION"
                post_data['opacod2'] = "OPH 1342"
                post_data['opaicd0'] = "08RJ3JZ"
                post_data['opaicdnm0'] = "Replacement of Right Lens with Synthetic Substitute, Percutaneous Approach"
                post_data['opaicd1'] = "08RK3JZ"
                post_data['opaicdnm1'] = "Replacement of Left Lens with Synthetic Substitute, Percutaneous Approach"
        elif data_gsheet['OP_TYPE'] == 'VT': # TODO
            if data_gsheet['OP_SIDE'] == "OD":
                post_data['opanam1'] = "OCUTOME PARS PLANA VITRECTOMY ( V.T. ), COMPLICATED"
                post_data['opacod1'] = "OPH 14791"
                post_data['opaicd0'] = "08B43ZZ"
                post_data['opaicdnm0'] = "Excision of Right Vitreous, Percutaneous Approach"
            elif data_gsheet['OP_SIDE'] == "OS":
                post_data['opanam1'] = "OCUTOME PARS PLANA VITRECTOMY ( V.T. ), COMPLICATED"
                post_data['opacod1'] = "OPH 14791"
                post_data['opaicd0'] = "08B53ZZ"
                post_data['opaicdnm0'] = "Excision of Left Vitreous, Percutaneous Approach"
            elif data_gsheet['OP_SIDE'] == "OU":
                post_data['opanam1'] = "OCUTOME PARS PLANA VITRECTOMY ( V.T. ), COMPLICATED"
                post_data['opacod1'] = "OPH 14791"
                post_data['opanam2'] = "OCUTOME PARS PLANA VITRECTOMY ( V.T. ), COMPLICATED"
                post_data['opacod2'] = "OPH 14791"
                post_data['opaicd0'] = "08B43ZZ"
                post_data['opaicdnm0'] = "Excision of Right Vitreous, Percutaneous Approach"
                post_data['opaicd1'] = "08B53ZZ"
                post_data['opaicdnm1'] = "Excision of Left Vitreous, Percutaneous Approach"
        elif data_gsheet['OP_TYPE'] == 'TRABE': # TODO
            pass
        
        # 處理術前診斷
        post_data['diagn'] = data_web9op.get("diagn")  # 抓排程內容 # TODO 換成刀表的診斷? 刀表比較沒有時間差問題?
        
        # 處理術後診斷(應該以google sheet為主，因為有些會有加打IVIA)
        if existandnotnone(data_gsheet, 'COL_LENSX') and (data_gsheet['COL_OP'].upper().find('LEN') == -1):  # 判斷COL_LENSX有資料+COL_OP內沒有LEN(SX)關鍵字 => 術式要補上LENSX
            post_data['diaga'] = f"Ditto s/p LenSx+{data_gsheet['COL_OP']}"
        else:
            post_data['diaga'] = f"Ditto s/p {data_gsheet['COL_OP']}"
        
        # 確認術後診斷有側別
        if check_op_side(post_data['diaga']) is None:
            post_data['diaga'] = post_data['diaga'] + f" {data_gsheet['OP_SIDE']}"

        # 處理病歷內文:
        df_template = self.gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_TEMPLATE_OPNOTE)
        df_template_selected = df_template.loc[ 
            (df_template['OP_TYPE']==data_gsheet['OP_TYPE'])
            &((df_template['VS_CODE']==self.config['VS_CODE']) | (df_template['VS_CODE']=='0'))
            &((df_template['R_CODE']==self.config['R_CODE']) | (df_template['R_CODE']=='0')),:
            ].sort_values(by =['VS_CODE','R_CODE'], axis=0, ascending=False)
        template = df_template_selected.iloc[0,3]
    
        if data_gsheet['OP_TYPE'] == 'PHACO' or data_gsheet['OP_TYPE'] == 'ECCE' or data_gsheet['OP_TYPE'] == 'LENSX':
            
            # 處理complications紀錄，如果沒有Complications輸出Nil
            if not existandnotnone(data_gsheet, 'COL_COMPLICATIONS'):
                data_gsheet['COL_COMPLICATIONS'] = 'Nil'

            # 處理IOL細節(IOL種類+度數+Target+SN)
            data_gsheet['DETAILS_OF_IOL'] = f"Style of IOL: {data_gsheet['COL_IOL']} F+{data_gsheet['COL_FINAL']}"
            # 若有Target加上
            if existandnotnone(data_gsheet, 'COL_TARGET'):
                data_gsheet['DETAILS_OF_IOL'] = data_gsheet['DETAILS_OF_IOL'] + f" Target: {data_gsheet['COL_TARGET']}"
            # 若有SN加上
            if existandnotnone(data_gsheet, 'COL_SN'):
                data_gsheet['DETAILS_OF_IOL'] = data_gsheet['DETAILS_OF_IOL'] + f" SN: {data_gsheet['COL_SN']}"
            
            # 組裝病歷內文
            post_data['op2data'] = Template(template).substitute(data_gsheet)

        elif data_gsheet['OP_TYPE'] == 'VT': # TODO
            pass
            # post_data['op2data'] = Template(template).substitute(data_gsheet)
        elif data_gsheet['OP_TYPE'] == 'TRABE': # TODO
            pass
            # post_data['op2data'] = Template(template).substitute(data_gsheet)


        return post_data


    def recheck_print(self):  # 確認一次要處理的資料有沒有錯誤
        print(f"手術紀錄日期: {self.date}")  # TODO 可能跨越多個日期??
        print(f"VS:{self.config['VS_CODE']}||R:{self.config['R_CODE']}")
        data = self.data
        t = list()
        # TODO 因為CATA/VT/TRABE是否要有不同
        for hisno in data.keys(): 
            t.append(
                (
                    hisno,
                    data[hisno]['data_web9op']['name'],
                    data[hisno]['data_gsheet']['OP_TYPE'],
                    data[hisno]['data_gsheet']['COL_IOL'],
                    data[hisno]['data_gsheet']['COL_FINAL'],
                    data[hisno]['post_data']['diaga'],
                    data[hisno]['data_gsheet']['COL_SN'],
                    data[hisno]['data_gsheet']['COL_COMPLICATIONS'],
                )
            )
        printed_df = pandas.DataFrame(t, columns=['病歷號', '姓名', '手術種類', 'IOL', 'Final', '術後診斷', 'SN', '併發症'])
        print(f"紀錄清單:\n{printed_df}")

        mode = input("確認無誤(y/n) ").strip().lower()
        if mode == 'y':
            return True
        else:
            return False


class OPNote_SURGERY(OPNote):
    def __init__(self, webclient, config):
        df_available = super().__init__(webclient, config)
        if df_available:
            self.start()


class OPNote_IVI(OPNote):
    def __init__(self, webclient, config):
        df_available = super().__init__(webclient, config)
        if df_available:
            self.op_start = datetime.strptime(self.config['OP_START'], "%H%M")
            self.op_interval = timedelta(minutes=self.config['OP_INTERVAL'])
            self.start()

    def fill_data(self, **kwargs):
        data_web9op = kwargs.get('data_web9op')
        data_gsheet = kwargs.get('data_gsheet')
        data_opschedule = kwargs.get('data_opschedule')
        index = kwargs.get('num')

        # 填手術紀錄內容:
        post_data = {  # 應該可以透過一個需要擷取的list來取得這些需要的資訊
            "sect1": data_web9op.get("sect1"),
            "name": data_web9op.get("name"),
            "sex": data_web9op.get("sex"),
            "hisno": data_web9op.get("hisno"),
            "age": data_web9op.get("age"),
            "idno": data_web9op.get("idno"),
            "birth": data_web9op.get("birth"),
            "_antyp": data_web9op.get("_antyp"),
            "opbbgndt": data_web9op.get("opbbgndt"),
            "opbbgntm": data_web9op.get("opbbgntm"),
            "opscode_num": 1, "film": "N", "against": "N", "action": "NewOpa01Action", "signchk": "Y",
            "sect": "OPH", "ward": "OPD", "source": "O",  # source是表示門診
            "again": "N", "reason_aga": "0", "mirr": "N", "saw": "N", "hurt": "1", "posn1": "1", "posn2": "0",
            "cler1": "2", "cler2": "0", "item1": "3", "item2": "2",
            ##以下都是空字串想測試全部拿掉##
            "ass2n": "", "ass3n": "", "ant1n": "", "ant2n": "", "dirn": "", "trtncti": "", "babym": "", "babyd": "",
            "bed": "",
            "final": "", "ass2": "", "ass3": "", "ant1": "", "ant2": "", "dir": "", "antyp1": "", "antyp2": "",
            "side": "", "oper": "", "rout": "", "side01": "", "oper01": "", "rout01": "",
            "opanam2": "", "opacod2": "", "opanam3": "", "opacod3": "", "opanam4": "", "opacod4": "", "opanam5": "",
            "opacod5": "",
            "opaicd1": "", "opaicdnm1": "", "opaicd2": "", "opaicdnm2": "", "opaicd3": "", "opaicdnm3": "",
            "opaicd4": "", "opaicdnm4": "", "opaicd5": "", "opaicdnm5": "",
            "opaicd6": "", "opaicdnm6": "", "opaicd7": "", "opaicdnm7": "", "opaicd8": "", "opaicdnm8": "",
            "opaicd9": "", "opaicdnm9": "",
            ##以下需要變更##
            "man": "####", "ass1": "####", # 特殊處理
            "mann": "####", "ass1n": "####", # 特殊處理
            "bgndt": self.date, "enddt": self.date,
            "bgntm": (self.op_start + index * self.op_interval).strftime("%H%M"),
            "endtm": (self.op_start + index * self.op_interval + self.op_interval).strftime("%H%M"),
            "sel_opck": "",  # IVI 這欄位應該是空的
            "diagn": "##########",  # 特殊處理
            "diaga": "##########",  # 特殊處理
            "antyp": "LA",
            "opanam1": f"INTRAVITREAL INJECTION OF {data_gsheet['COL_DRUGTYPE'].upper()}",
            "opacod1": "OPH 1476",  # 使用通用碼
            "opaicd0": "3E0C3GC",
            "opaicdnm0": "Introduction of Other Therapeutic Substance into Eye, Percutaneous Approach",
            "op2data": "##########",  # 特殊處理
        }

        #### 特殊處理
        # 如果VS_CODE存在就使用新的CODE
        if existandnotnone(data_gsheet, 'COL_VS_CODE'):
            code = data_gsheet['COL_VS_CODE']
            post_data['man'] = code
            post_data['mann'] = get_name_from_code(code, self.session)
        else:
            post_data['man'] = self.config['VS_CODE']
            post_data['mann'] = self.config['VS_NAME']

        # 如果R_CODE存在就使用新的CODE
        if existandnotnone(data_gsheet, 'COL_R_CODE'):
            code = data_gsheet['COL_R_CODE']
            post_data['ass1'] = code
            post_data['ass1n'] = get_name_from_code(code, self.session)
        else:
            post_data['ass1'] = self.config['R_CODE']
            post_data['ass1n'] = self.config['R_NAME']


        # 術前診斷:
        if existandnotnone(data_gsheet, 'COL_DIAGNOSIS') and existandnotnone(data_gsheet, 'COL_SIDE'):
            post_data['diagn'] = f"{data_gsheet['COL_DIAGNOSIS']} {data_gsheet['COL_SIDE']}"
        else:
            raise Exception(f"資料輸入未完整!")
        
        # 術後診斷:
        post_data['diaga'] = (f"Ditto s/p IVI-{data_gsheet['COL_DRUGTYPE']} {data_gsheet['COL_SIDE']}")
        if existandnotnone(data_gsheet, 'COL_OTHER_TREATMENT'):
            post_data['diaga'] = post_data['diaga'] + f" + {data_gsheet['COL_OTHER_TREATMENT']}"

        # 想利用data_gsheet這個dictionary直接丟入template substitue，多傳入參數不會錯，少傳會報錯，除非使用safe_substitute
        data_gsheet['TRANSFORMED_SIDE'] = NOTE_TRANSFORM_SIDE.get(data_gsheet['COL_SIDE'], "")
        # data_gsheet['TRANSFORMED_DISTANCE'] = NOTE_TRANSFORM_IVIDISTANCE.get(data_gsheet['COL_PHAKIC'], "3.5")
        data_gsheet['TRANSFORMED_DISTANCE'] = '4.0'

        df_template = self.gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_TEMPLATE_OPNOTE)
        df_template_selected = df_template.loc[ 
            (df_template['OP_TYPE']=='IVI')
            &((df_template['VS_CODE']==self.config['VS_CODE']) | (df_template['VS_CODE']=='0'))
            &((df_template['R_CODE']==self.config['R_CODE']) | (df_template['R_CODE']=='0')), :
            ].sort_values(by =['VS_CODE','R_CODE'], axis=0, ascending=False)
        template = df_template_selected.iloc[0,3]
        post_data['op2data'] = Template(template).substitute(data_gsheet)

        return post_data
    

    def recheck_print(self):  # 確認一次要處理的資料有沒有錯誤
        print(f"=================\n手術紀錄日期: {self.date}")
        print(f"手術開始: {self.config['OP_START']} 間隔:{self.config['OP_INTERVAL']} 分鐘")
        data = self.data
        t = [(
            hisno,
            data[hisno]['post_data']['name'],
            data[hisno]['post_data']['diagn'],
            data[hisno]['post_data']['diaga'],
            data[hisno]['post_data']['man'],
            data[hisno]['post_data']['ass1']
        ) for hisno in data.keys()]
        printed_df = pandas.DataFrame(t, columns=['病歷號', '姓名', '診斷', '處置', "VS帳號", "R帳號"])
        print(f"紀錄清單:\n{printed_df}")

        mode = input("確認無誤(y/n) ").strip().lower()
        if mode == 'y':
            return True
        else:
            return False


def existandnotnone(dictionary: dict, key):
    if dictionary.get(key) is not None:  # 不是None
        if type(dictionary[key]) == str and len(dictionary[key].strip()) > 0:  # 字串的話不是空白
            return True
        elif type(dictionary[key]) == int:
            return True
    return False


def get_name_from_code(id_code, session): 
    # TODO 需要加上assertion來確認這個函數正常運行嗎?

    _url = "https://web9.vghtpe.gov.tw/emr/OPAController"
    payload = {
        "doc": str(id_code),
        "action": "CheckDocAction"
    }
    res = session.get(_url, params = payload)
    name = res.text.strip()
    if len(name)==0:
        print("取得燈號對應姓名異常")
        return input(f"請輸入 {id_code} 的姓名:")
    else:
        return name


def check_opdate(default=None): # TODO 手術(除IVI)時間應該可以用排程系統抓取
    '''
    預設為當日日期，允許使用者更改
    '''
    if default is None:
        default = str(datetime.today().year - 1911) + datetime.today().strftime("%m%d")  # 採用中華民國紀年，沒有指定就以當下時間

    date_final = input(f"目前預設手術紀錄日期為: {default}\n(正確請按enter，錯誤請輸入新日期[格式:{default}])\n==請問正確嗎? ")
    if len(date_final.strip()) == 0:
        return default
    else:
        while True:
            if re.match(r"^1\d{6}$", date_final) is None:
                print("格式有誤，請重新輸入")
                date_final = input(f"請更正手術日期[格式:{default}]: ")
            else:
                return date_final


def check_op_side(input_string:str):
    '''
    偵測字串中是否有側別資訊，回傳'OD', 'OS', 'OU', None
    '''
    if input_string is None: #如果傳入資料為None
        return None
    elif type(input_string) != str: #如果傳入資料不為字串
        return None
    
    if input_string.upper().find('OD') > -1: #右側
        return 'OD'
    elif input_string.upper().find('OS') > -1: #左側
        return 'OS'
    elif input_string.upper().find('OU') > -1: #雙側
        return 'OU'
    else:
        return None
    
def check_op_type(input_string:str):
    '''
    偵測字串中為何種手術類別，回傳'phaco', 'ecce', 'lensx', 'vt', None
    '''
    if input_string is None: #如果傳入資料為None
        return None
    elif type(input_string) != str: #如果傳入資料不為字串
        return None

    if (input_string.upper().find('LENSX') > -1) or (input_string.upper().find('LENS') > -1): # lensx
        return 'LENSX'
    elif (input_string.upper().find('ECCE') > -1): # ecce 
        return 'ECCE'
    elif input_string.upper().find('PHACO') > -1: # phaco
        return 'PHACO'
    elif input_string.upper().find('VT') > -1: # vt
        return 'VT'
    elif input_string.upper().find('TRABE') > -1: # trabe
        return 'TRABE'
    else:
        return None


# Logging 設定
logger = logging.getLogger()
logger.setLevel(logging.INFO)  # 這是logger的level
BASIC_FORMAT = '[%(asctime)s %(levelname)-8s] %(message)s'
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
formatter = logging.Formatter(BASIC_FORMAT, datefmt=DATE_FORMAT)
# 設定console handler的設定
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)  # 可以獨立設定console handler的level，如果不設就會預設使用logger的level
ch.setFormatter(formatter)
# 設定file handler的設定
log_filename = "vghbot_note.log"
fh = logging.FileHandler(log_filename)  # 預設mode='a'，持續寫入
fh.setLevel(logging.INFO)
fh.setFormatter(formatter)
# 將handler裝上
logger.addHandler(ch)
logger.addHandler(fh)

# 轉換side  => 'TRANSFORMED_SIDE'
NOTE_TRANSFORM_SIDE = {  # 如果都不符就輸出空白
    'OD': "RIGHT",
    'OS': "LEFT",
    'OU': "BOTH"
}

# 轉換打針距離  => 'TRANSFORMED_DISTANCE'
NOTE_TRANSFORM_IVIDISTANCE = {  # 如果都不符就輸出3.5-4mm
    'TRUE': "4",
    'FALSE': "3.5"
}


TEST_MODE = False
UPDATER_OWNER = 'zmh00'
UPDATER_REPO = 'vghbot'
UPDATER_FILENAME = 'op'
UPDATER_VERSION_TAG = 'v1.0'

if __name__ == '__main__':
    if TEST_MODE:
        print("##############測試模式##############\n")
    else:
        warnings.simplefilter("ignore")
        # Check if the version is the latest
        u = updater_cmd.Updater_github(UPDATER_OWNER, UPDATER_REPO, UPDATER_FILENAME, UPDATER_VERSION_TAG)
        if u.start() == False:
            exit()
    

    # 選擇CATA|IVI mode
    while True:
        mode = input("Choose the OP Note mode (1:SURGERY | 2:IVI | 0:EXIT): ")
        if mode not in ['1','2','0']:
            print("WRONG MODE INPUT")
        elif mode == '0':
            break
        else:
            config = dict()
            if mode == '1':
                #手術(白內障)紀錄模式
                config['BOT_MODE'] = 'SURGERY'
                gc = gsheet.GsheetClient()
                df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_SURGERY)
            elif mode == '2':
                #IVI紀錄模式
                config['BOT_MODE'] = 'IVI'
                #自動下載排程 # TODO

                #自動更新BOT # TODO

                #打開BOT讓使用者編輯 # TODO
                
                #重新讀取並送出
                gc = gsheet.GsheetClient()
                df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_IVI)
            while True:
                selected_col = ['INDEX','VS_CODE','SPREADSHEET','WORKSHEET']
                selected_df = df.loc[:, selected_col]
                selected_df.index +=1 # 讓index從1開始方便選擇
                selected_df.rename(columns={'INDEX':'組套名'}, inplace=True) # rename column
                # 印出現有組套讓使用者選擇
                print("\n=========================")
                print(selected_df) 
                print("=========================")
                selection = input("請選擇以上profile(0是退回): ")
                if selection != '0':
                    if int(selection) in selected_df.index:
                        # 與drweb連線
                        webclient = vghbot_login.Client(TEST_MODE=TEST_MODE)
                        webclient.login_drweb()

                        # R_CODE以webclient登入帳密轉換
                        r_code = webclient.login_id[3:7]
                        config['R_CODE'] = r_code

                        # 將選擇的組套匯入config
                        config.update( df.loc[int(selection)-1,:].to_dict() )
                        if config['BOT_MODE'] == 'IVI': 
                            OPNote_IVI(webclient, config)
                        elif config['BOT_MODE'] == 'SURGERY':
                            OPNote_SURGERY(webclient, config)
                    else:
                        print("!!選擇錯誤!!\n")
                else:  # 等於0 => 退到上一層
                    break
