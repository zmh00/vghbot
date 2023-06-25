from bs4 import BeautifulSoup
from string import Template
from datetime import datetime, timedelta
import pygsheets
import pandas
import re
import logging
import json
import pickle
import pprint
import warnings

# import bot_login
import gsheet
import vghbot_login

class OPNote():
    def __init__(self, webclient, config):
        ## 將session轉移進來
        self.session = webclient.session

        ## Input1: config設定檔，因為config內容是所有病人共用，所以安排在此處理，如果是個別病人的資訊調整會在fill_data內
        self.config = config
        ################ ################
        #這是為了新增VS_CODE和R_CODE在googlesheet欄位而做的hack，感覺應該重新設計才比較有結構
        #目前設計是當指到spreadsheet為"BOT"就會啟用google sheet取得內部的VS_CODE和R_CODE
        ################ ################
        self.config['VS_NAME'] = get_name_from_code(self.config['VS_CODE'], self.session)
        self.config['R_NAME'] = get_name_from_code(self.config['R_CODE'], self.session)
        if self.config['SPREADSHEET'] == "BOT": #處理特殊狀況
            print(f"目前使用共用BOT範本，VS與R登號會以google sheet設定")
        else: #正常狀況在這就取得VS、R的CODE和NAME，但可能需要重新設計會比較乾淨
            print(f"目前使用Config:{self.config['BOT']} || VS{self.config['VS_NAME']}({self.config['VS_CODE']}) || R{self.config['R_NAME']}({self.config['R_CODE']})")

        ## Input2: google sheet資料取得: 多人Dataframe
        self.df_of_worksheet = get_dataframe_from_gsheet(self.config['SERVICE_FILE'], self.config['SPREADSHEET'],
                                                         self.config['WORKSHEET'])
        self.date, self.extracted_df = extract(self.df_of_worksheet, self.config['COL_HISNO'], self.config['COL_DATE'],
                                               self.config['DATE_PATTERN'])

        size_of_df = len(self.extracted_df)
        if size_of_df == 0:
            print("完成Dataframe擷取: 無匹配資料")
            return False
        else:
            print(f"完成Dataframe擷取: {size_of_df}筆資料")
            self.data = {}  # 存所有的資料 = post_data(patient_info) + google sheet + config資料
            # self.data[hisno] = {
            #     'data_web9op': data_web9op,
            #     'data_gsheet': data_gsheet,
            #     'post_data': post_data
            # }
            return True

    def start(self):
        num = 0
        # 針對每一人的資料處理歸檔(變成dictionary)
        for hisno in self.extracted_df[self.config['COL_HISNO']]:  # 使用病歷號作為識別資料
            if type(hisno) != str:  # pandas這個欄位是混雜數字空格，因此是object格式，若非數字會被轉成str輸出，選擇病歷號要排除str物件
                data_web9op = self.get_data_web9op(hisno)  # 取得特定病人web9基本資料
                data_gsheet = self.get_data_gsheet(hisno)  # 取得特定病人googgle表單資料
                post_data = self.fill_data(hisno, num, data_web9op, data_gsheet, self.config)  # 不同的Note會有客製化填資料
                self.data[hisno] = {
                    'data_web9op': data_web9op,
                    'data_gsheet': data_gsheet,
                    'post_data': post_data
                }
                num = num + 1

        if self.recheck_print():
            for hisno in self.data.keys():
                res = self.post(self.data[hisno]["post_data"])
                self.data[hisno]["post_response"] = res  # TODO for debug用途未來考慮刪除
                logger.info(f"新增記錄: {hisno} || {self.data[hisno]['data_web9op']['name']}")
        logger.info(f"完成本次自動化作業，共{num}筆資料")

        if TEST_MODE:
            self.save()  # TODO for debug用途未來考慮刪除

    def save(self):  # 收集資料測試用
        with open(f"{datetime.today().strftime('%Y%m%d_%H%M%S')}_data", 'wb') as f:
            pickle.dump(self.data, file=f)

    def post(self, data):  # 送出資料
        if TEST_MODE == True:
            target_url = "https://httpbin.org/post"
        else:
            target_url = "https://web9.vghtpe.gov.tw/emr/OPAController"
        response = self.session.post(target_url, data=data)
        if TEST_MODE == True:
            pprint.pprint(response.json())
        return response

    def get_data_web9op(self, hisno):  # 抓取病人基本資料和處理手術時間問題(google sheet和手術護理時間的匹配)
        # Input: hisno # Output: 爬到的病人基本資料
        # 取得病人基本資料 => 新增手術紀錄需要
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
        if self.config['BOT'] != "IVI":  # IVI的note不用擷取護理紀錄時間
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
            "name": pt_name,
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

    def get_data_gsheet(self, hisno):
        # Input: hisno, config # Output: 指定hisno的google sheet資料
        config = self.config
        df = self.extracted_df.loc[self.extracted_df[config['COL_HISNO']] == hisno, :]  # 取得對應hisno的該筆googlesheet資料(row)
        df_col = {}  # 依據有特別標記的col_name(以COL為開頭)去取得df該行的資料，方便之後存取
        if len(df)>1:
            logger.error(f"指定範圍內有重覆的病歷號[{hisno}]，只會取得最上面的資料列")
            df = df.iloc[[0],:] 
        for key in config.keys():
            if key[:3] == 'COL':
                if config[key] in df.columns:
                    df_col[key] = df[config[key]].item()
                else:
                    df_col[key] = None  # 如果config內有此變數但google sheet上沒有對應的column
        return df_col

    def fill_data(self):  # return為一個要post的dict型態payload
        print("需要實做客製化的fill data")

    def recheck_print(self):  # return True or False
        print("需要實做客製化的recheck print")
        # print(f"手術紀錄日期: {self.date}")
        # data = self.data
        # printed_df = pandas.DataFrame()
        # print(f"紀錄清單:\n{printed_df}")
        #
        # mode = input("確認無誤(y/n) ").strip().lower()
        # if mode == 'y':
        #     return True
        # else:
        #     return False


class OPNote_CATA(OPNote):
    def __init__(self, webclient, config):
        df_available = super().__init__(webclient, config)  # 判斷有無抓取到資料
        if df_available:
            if self.date is None:
                self.date = check_opdate()

            self.start()

    def fill_data(self, hisno, index, data_web9op, data_gsheet, config):
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
            "man": "####", "ass1": "####",
            "mann": "####", "ass1n": "####",
            "bgndt": self.date, "enddt": self.date,
            "bgntm": data_web9op.get('bgntm'),  # 由CATA護理紀錄時間
            "endtm": data_web9op.get('endtm'),  # 由CATA護理紀錄時間
            "sel_opck": data_web9op.get("sel_opck"),
            "diagn": "##########",  # 特殊處理
            "diaga": "##########",  # 特殊處理
            "antyp": "RA",  # TODO 這要用擷取的?
            "opanam1": "", # 特殊處理
            "opacod1": "", # 特殊處理
            "opanam2": "", # 特殊處理
            "opacod2": "", # 特殊處理
            "opaicd0": "", # 特殊處理
            "opaicdnm0": "", # 特殊處理
            "opaicd1": "", # 特殊處理
            "opaicdnm1": "", # 特殊處理
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
        # 處理ODOS左右邊轉換
        data_gsheet['TRANSFORMED_SIDE'] = NOTE_TRANSFORM_SIDE.get(data_gsheet['COL_SIDE'], "")
        # 處理雙眼OU資料
        if data_gsheet['COL_SIDE'].upper()=="OD":
            post_data['opanam1'] = "PHACOEMULSIFICATION + PC-IOL IMPLANTATION"
            post_data['opacod1'] = "OPH 1342"
            post_data['opaicd0'] = "08RJ3JZ"
            post_data['opaicdnm0'] = "Replacement of Right Lens with Synthetic Substitute, Percutaneous Approach"
        elif data_gsheet['COL_SIDE'].upper()=="OS":
            post_data['opanam1'] = "PHACOEMULSIFICATION + PC-IOL IMPLANTATION"
            post_data['opacod1'] = "OPH 1342"
            post_data['opaicd0'] = "08RK3JZ"
            post_data['opaicdnm0'] = "Replacement of Left Lens with Synthetic Substitute, Percutaneous Approach"
        elif data_gsheet['COL_SIDE'].upper()=="OU":
            post_data['opanam1'] = "PHACOEMULSIFICATION + PC-IOL IMPLANTATION"
            post_data['opacod1'] = "OPH 1342"
            post_data['opanam2'] = "PHACOEMULSIFICATION + PC-IOL IMPLANTATION"
            post_data['opacod2'] = "OPH 1342"
            post_data['opaicd0'] = "08RJ3JZ"
            post_data['opaicdnm0'] = "Replacement of Right Lens with Synthetic Substitute, Percutaneous Approach"
            post_data['opaicd1'] = "08RK3JZ"
            post_data['opaicdnm1'] = "Replacement of Left Lens with Synthetic Substitute, Percutaneous Approach"
        # 處理complications紀錄
        if not existandnotnone(data_gsheet, 'COL_COMPLICATIONS'):
            data_gsheet['COL_COMPLICATIONS'] = 'Nil'
        # 處理IOL細節
        data_gsheet['DETAILS_OF_IOL'] = f"Style of IOL: {data_gsheet['COL_IOL']} F+{data_gsheet['COL_FINAL']}"
        if existandnotnone(data_gsheet, 'COL_SN'):
            data_gsheet['DETAILS_OF_IOL'] = data_gsheet['DETAILS_OF_IOL'] + f" SN: {data_gsheet['COL_SN']}"

        # 術前診斷
        post_data['diagn'] = data_web9op.get("diagn")  # 抓排程內容
        # TODO # post_data['diagn'] = data_gsheet['COL_DIAGNOSIS']  # 需要改成抓gsheet資料嗎?但也可以考慮多數仰賴排程系統資料比較快?

        # 術後診斷應該以google sheet為主 => 因為有些會有加打IVIA
        # 術後診斷 + 處理使用範本 => 分成三種LenSx/ECCE/Phaco
        # 如果術式資訊內含側別就會直接使用，沒有側別就會加入側別欄位資訊
        if existandnotnone(data_gsheet, 'COL_LENSX'):  # 單純判斷COL_LENSX有沒有資料
            post_data['diaga'] = f"Ditto s/p LenSx+{data_gsheet['COL_OP']}"
            if check_side_exist_in_string(post_data['diaga']) is None:
                post_data['diaga'] = post_data['diaga'] + f" {data_gsheet['COL_SIDE']}"
            post_data['op2data'] = Template(config['TEMPLATE_LENSX']).substitute(data_gsheet)
        elif data_gsheet['COL_OP'].lower().find('ecce') > -1:
            post_data['diaga'] = f"Ditto s/p {data_gsheet['COL_OP']}"
            if check_side_exist_in_string(post_data['diaga']) is None:
                post_data['diaga'] = post_data['diaga'] + f" {data_gsheet['COL_SIDE']}"
            post_data['op2data'] = Template(config['TEMPLATE_ECCE']).substitute(data_gsheet)
        else:
            post_data['diaga'] = f"Ditto s/p {data_gsheet['COL_OP']}"
            if check_side_exist_in_string(post_data['diaga']) is None:
                post_data['diaga'] = post_data['diaga'] + f" {data_gsheet['COL_SIDE']}"
            post_data['op2data'] = Template(config['TEMPLATE']).substitute(data_gsheet)

        return post_data

    def recheck_print(self):  # 確認一次要處理的資料有沒有錯誤
        print(f"手術紀錄日期: {self.date}")  # TODO 可能跨越多個日期??
        data = self.data
        t = list()
        for hisno in data.keys():
            if data[hisno]['post_data']['diaga'].lower().find('lensx') > -1:
                op_type = "LENSX"
            elif data[hisno]['post_data']['diaga'].lower().find('ecce') > -1:
                op_type = "ECCE"
            else:
                op_type = ""
            t.append(
                (
                    hisno,
                    data[hisno]['data_web9op']['name'],
                    op_type,
                    data[hisno]['data_gsheet']['COL_IOL'],
                    data[hisno]['data_gsheet']['COL_FINAL'],
                    data[hisno]['data_gsheet']['COL_SN'],
                    data[hisno]['data_gsheet']['COL_COMPLICATIONS'],
                    data[hisno]['post_data']['bgntm'],
                    data[hisno]['post_data']['endtm'],  # TODO 增加手術日期欄位
                    data[hisno]['post_data']['man'],
                    data[hisno]['post_data']['ass1']
                )
            )
        printed_df = pandas.DataFrame(t, columns=['病歷號', '姓名', '手術', 'IOL', 'Final', 'SN', '併發症', '開始', '結束', "VS帳號", "R帳號"])
        print(f"紀錄清單:\n{printed_df}")

        mode = input("確認無誤(y/n) ").strip().lower()
        if mode == 'y':
            return True
        else:
            return False


class OPNote_IVI(OPNote):
    def __init__(self, webclient, config):
        df_available = super().__init__(webclient, config)
        if df_available:
            if self.date is None:
                self.date = check_opdate()

            self.op_start = datetime.strptime(self.config['OP_START'], "%H%M")
            self.op_interval = timedelta(minutes=self.config['OP_INTERVAL'])
            self.start()

    def fill_data(self, hisno, index, data_web9op, data_gsheet, config):
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
            "man": "####", "ass1": "####",
            "mann": "####", "ass1n": "####",
            "bgndt": self.date, "enddt": self.date,
            "bgntm": (self.op_start + index * self.op_interval).strftime("%H%M"),
            "endtm": (self.op_start + index * self.op_interval + self.op_interval).strftime("%H%M"),
            "sel_opck": "",  # IVI 這欄位應該是空的
            "diagn": "##########",  # 特殊處理
            "diaga": "##########",  # 特殊處理
            "antyp": "LA",  # TODO 這要用擷取的?
            "opanam1": f"INTRAVITREAL INJECTION OF {data_gsheet['COL_DRUGTYPE'].upper()}",
            "opacod1": "OPH 1476",  # 使用通用碼
            "opaicd0": "3E0C3GC",
            "opaicdnm0": "Introduction of Other Therapeutic Substance into Eye, Percutaneous Approach",
            "op2data": "##########",  # 特殊處理
        }
        # 特殊處理
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
        # 如果google_sheet沒有OP1資訊自己組裝
        if existandnotnone(data_gsheet, 'COL_OP1'):
            post_data['diagn'] = data_gsheet['COL_OP1']
        else:
            post_data['diagn'] = f"{data_gsheet['COL_DIAGNOSIS']} {data_gsheet['COL_SIDE']}"
        #如果google_sheet沒有OP2資訊自己組裝
        if existandnotnone(data_gsheet, 'COL_OP2'):
            post_data['diaga'] = data_gsheet['COL_OP2']
        else:
            post_data['diaga'] = (f"Ditto s/p IVI-{data_gsheet['COL_DRUGTYPE']} {data_gsheet['COL_SIDE']}")
            if existandnotnone(data_gsheet, 'COL_OTHER_TREATMENT'):
                post_data['diaga'] = post_data['diaga'] + f" + {data_gsheet['COL_OTHER_TREATMENT']}"

        # 想利用data_gsheet這個dictionary直接丟入template substitue，多傳入參數不會錯，少傳會報錯，除非使用safe_substitute
        data_gsheet['TRANSFORMED_SIDE'] = NOTE_TRANSFORM_SIDE.get(data_gsheet['COL_SIDE'], "")
        data_gsheet['TRANSFORMED_DISTANCE'] = NOTE_TRANSFORM_IVIDISTANCE.get(data_gsheet['COL_PHAKIC'], "3.5")
        post_data['op2data'] = Template(config['TEMPLATE']).substitute(data_gsheet)

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


class OPNote_DELETE():  # TODO 尚未實作
    def __init__(self, hisno, date):
        super().__init__()
        r_app = super().app('DRWEBAPP')
        print("完成web9 cookies取得")

        # target_url = "https://web9.vghtpe.gov.tw/emr/OPAController"
        # payload = {
        #     "action": "DeleteOpnFormAction",
        #     "hisno": hisno,
        #     "dt": date,
        #     "tm": "1"
        # }
        # self.response = self.session.get(target_url, params=payload)
        # 這段可以解析times

        target_url = "https://web9.vghtpe.gov.tw/emr/OPAController?action=DeleteOpnAction"
        payload = {
            "histno": hisno,
            "bgndt": date,  # 1110121
            "times": 1  # 這不知道是不是會遞增??
        }
        self.response = self.session.post(target_url, data=payload)
        print(f"刪除記錄: {hisno}|{date}")


def existandnotnone(dictionary: dict, key):
    if dictionary.get(key) is not None:  # 不是None
        if type(dictionary[key]) == str and len(dictionary[key].strip()) > 0:  # 字串的話不是空白
            return True
        elif type(dictionary[key]) == int:
            return True
    return False


# 以下函數是取得gsheet資料並轉成dataframe
def get_dataframe_from_gsheet(service_file, spreadsheet, worksheet):  # return dataframe
    client = pygsheets.authorize(service_file=service_file)
    ssheet = client.open(spreadsheet)
    wsheet = ssheet.worksheet_by_title(worksheet)
    df = wsheet.get_as_df(include_tailing_empty=False)
    return df


# 以下函數是辨識指定格式日期列
def match_identifier_date(index, input, identifier):
    try:
        date = datetime.strptime(input, identifier)
        if identifier.find("Y") == -1:  # 如果辨識的pattern沒有年，就自行新增
            date = date.replace(year=datetime.today().year)
        return str(date.year - 1911) + date.strftime("%m%d")  # 改成中華民國紀年
    except:
        if input != "":
            print(f"{index + 2:2} 列:疑似日期格式配對異常:{input}|{identifier}")
        return False


def extract(df: pandas.DataFrame, column_hisno, column_date, identifier_date):  # return date, dataframe
    date_final = None  # 之後應該判斷若未指定應該要特別處理
    row_list_input = input("(-)表示連續範圍 (,)做分隔 (直接按enter)表示自動擷取最上方時間區塊模式\n請輸入符合格式gsheet列碼: ")
    if len(row_list_input.strip()) == 0:
        qualified_row = []
        meet_start = False
        for i, x in enumerate(df.loc[:, column_hisno]):  # 針對病歷號的欄位處理，邏輯是要抓出日期分隔線(灰色條)，有資料的，沒有資料的
            if x == '':  # pandas這個欄位是混雜數字空格，因此是object格式，若非數字會被轉成str輸出，選擇空白欄位要找str物件
                date = match_identifier_date(i, df.iloc[i, df.columns.get_loc(column_date)], identifier_date)
                if date:  # 想判斷出灰色分隔條的位置: 該row沒有病歷號但第一格卻有日期格式(ex:1/21)，若是單純空格就跳過
                    if meet_start:  # 已經通過第一灰條
                        # meet_end = True
                        break
                    else:
                        meet_start = True
                        date_final = date
            else:
                if meet_start:  # 要有通過時間條開始才開始記錄
                    qualified_row.append(i)
        # choice = input(f"手術紀錄日期為(自動辨識): {date_final}\n==請問正確嗎?(y/n)")
        # if choice.strip().lower() == 'n':
        #     while True:
        #         date_final = input("請更正手術日期(格式:1110121): ")
        #         if re.match(r"^1\d{6}$", date_final) is None:
        #             print("格式有誤，請重新輸入")
        #         else:
        #             break
        check_opdate(date_final)
        return date_final, df.iloc[qualified_row, :] # 自動擷取模式會回傳date_final的時間
    else:
        row_set = set()
        for i in row_list_input.split(','):
            if len(i.split('-')) > 1:  # 如果有範圍標示，把這段範圍加入set
                row_set.update(range(int(i.split('-')[0]), int(i.split('-')[1]) + 1))
            else:
                row_set.add(int(i))
        row_list = [row - 2 for row in
                    row_set]  # google spreadsheet的index是從1開始，所以df對應的相差一，此外還有一欄變成df的column headers，所以總共要減2
        row_list.sort()
        #### 針對IVI、其他手術，對於日期的判斷不同 => 改成在各自的class內設定手術紀錄日期

        return date_final, df.iloc[row_list, :] # 手動模式會回傳date_final=None

def get_name_from_code(id_code, session): 
    #TODO 應該把這個功能放在webclient架構下，因為需要連線的session

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

def check_opdate(default=None): # TODO 白內障的手術時間應該可以用排程系統抓取
    if default is None:
        default = str(datetime.today().year - 1911) + datetime.today().strftime("%m%d")  # 採用中華民國紀年，沒有指定就以當下時間

    date_final = input(f"目前預設手術紀錄日期為: {default}\n(正確請按enter，錯誤請輸入新日期[格式:1110121])\n==請問正確嗎? ")
    if len(date_final.strip()) == 0:
        return default
    else:
        while True:
            if re.match(r"^1\d{6}$", date_final) is None:
                print("格式有誤，請重新輸入")
                date_final = input("請更正手術日期[格式:1110121]: ")
            else:
                return date_final

def check_side_exist_in_string(input_string:str):
    if input_string.title().find('Od') > -1: #有右側的存在
        return 'OD'
    elif input_string.title().find('Os') > -1:
        return 'OS'
    elif input_string.title().find('Ou') > -1:
        return 'OU'
    else:
        return None

def acquire_id_psw():    
    while(1):
        login_id = input("Enter your ID: ")
        if len(login_id) != 0:
            break
    while(1):
        login_psw = input("Enter your PASSWORD: ")
        if len(login_psw) != 0:
            break
    return login_id, login_psw

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
log_filename = "notebot.log"
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

JSON_PATH = "config_bot_opnote.json"
TEST_MODE = False


if __name__ == '__main__':
    if TEST_MODE:
        print("##############測試模式##############\n")
    else:
        warnings.simplefilter("ignore")

    webclient = bot_login.Login()
    while True:
        checked = webclient.login_eip()
        if checked:
            checked = webclient.login_web9()
            r_app = webclient.app('DRWEBAPP')  # 完成DRWEBAPP登入
            break
    
    with open(JSON_PATH, "r") as f:
        config_dict = json.load(fp=f)
        while True:
            print("\n=========================")
            for i in config_dict.keys():
                if config_dict[i]['SPREADSHEET'] == 'BOT':
                    print(f"{i}: (BOT共用範本) {config_dict[i]['BOT']}|預設:{config_dict[i]['VS_CODE']}")
                else:
                    print(f"{i}: {config_dict[i]['BOT']}|{config_dict[i]['VS_CODE']}")
            print("=========================")
            selection = input("請選擇以上profile(0是退出): ")
            if selection != '0':
                if selection in config_dict.keys():
                    config = config_dict[selection]
                    if config['BOT'] == 'IVI':
                        OPNote_IVI(webclient, config)
                    elif config['BOT'] == 'CATA':
                        OPNote_CATA(webclient, config)
                else:
                    print("!!選擇錯誤!!\n")
            else:  # 等於0 => 退出
                break

'''
config_dict = {
    '1': {
        "BOT":"IVI",
        "SERVICE_FILE":"vghbot-5fe0aba1d3b9.json",
        "SPREADSHEET": "OP-4053周昱百",
        "WORKSHEET": "IVI",
        "VS_NAME": "周昱百",
        "VS_CODE": 4053,
        "R_NAME": "鄭明軒",
        "R_CODE": 4123,
        "DATE_PATTERN": "%Y/%m/%d",
        "COL_DATE": "時間",
        "COL_HISNO": "病歷號",
        "COL_NAME": "姓名",
        "COL_DIAGNOSIS": "診斷",
        "COL_SIDE": "側別",
        "COL_DRUGTYPE": "藥物種類",
        "COL_CHARGE": "收費",
        "COL_OTHER_TREATMENT": "STK/AC paracentesis",
        "COL_OP1": "Note_OP1",
        "COL_OP2": "Note_OP2",
        "COL_PHAKIC": "Phakic",
        "OP_INTERVAL": 5,
        "OP_START": "0900",
        "TEMPLATE":
            '1. Under LA, the $TRANSFORMED_SIDE eye was prepared and disinfected as usual.\n'
            '2. IVI was performed via pars plana $TRANSFORMED_DISTANCE(mm) from conjunctival limbus.\n'
            '3. Check bleeding with the cautery.\n'
            '4. The patient understood the whole procedures well.\n'
            'Complications: Nil'
    },
    '2': {
        "BOT":"CATA",
        "SERVICE_FILE":"vghbot-5fe0aba1d3b9.json",
        "SPREADSHEET": "OP-4053周昱百",
        "WORKSHEET": "OP",
        "VS_NAME": "周昱百",
        "VS_CODE": 4053,
        "R_NAME": "鄭明軒",
        "R_CODE": 4123,
        "DATE_PATTERN": "%Y/%m/%d",
        "COL_DATE": "時間",
        "COL_HISNO": "病歷號",
        "COL_NAME": "姓名",
        "COL_DIAGNOSIS": "診斷",
        "COL_SIDE": "側別",
        "COL_OP": "手術",
        "COL_LENSX": "LenSx",
        "COL_IOL": "IOL",
        "COL_FINAL": "Final",
        "COL_SN": "SN",
        "COL_COMPLICATIONS": "Complications",
        "COL_CDE": "CDE",
        "TEMPLATE":
            '1. After local anesthesia, $TRANSFORMED_SIDE eye was sterilized and draped.\n'
            '2. Moderated and severe lens opacity was identified under microscopy through illumination.\n' 
            '3. The upper and lower eyelids were opened with eyelid speculum.\n'
            '4. A temporal corneal incision was made with phaco-knife and then viscoelastic material was injected to the anterior chamber.\n'
            '5. Continous curvilinear anterior capsulorrhexis was performed with a capsular forceps. Hydrodissection of the lens capsule from cortex completely with resultant fluid waves seen.\n'
            '6. Side-port was made with 7515 bever knife over limbus. Ultrasonic phaco-tip was inserted into the anterior chamber and applicated for cataract extraction with assistance of a phaco-chopper.\n'
            '7. Residual cortex was removed from the capsular bag using IA.\n'
            '8. The viscoelastic material was injected into the capsular bag.\n'
            '9. A posterior chamber intraocular lens was inserted into the bag.\n'
            '10.Viscoelastic material was washed out with I/A irrigator. The wound was closed with stromal hydration.\n'
            '11.The patient stood the whole procedure well.\n\n'
            'Complication: $COL_COMPLICATIONS\n'
            '$DETAILS_OF_IOL',
        "TEMPLATE_LENSX":
            "1. After local anesthesia, Continuous curvilinear anterior capsulorrhexis of $TRANSFORMED_SIDE eye was performed with Lensx Femtosecond laser.\n"
            "2. Patient's eye was sterilized and draped as regular routine.\n"
            "3. Moderated and severe lens opacity was identified under microscopy through illumination.\n"
            "4. The upper and lower eyelids were opened with eyelid speculum.\n"
            "5. A temporal corneal incision was made with phaco-knife and then viscoelastic material was injected to the anterior chamber.\n"
            "6. Hydrodissection of the lens capsule from cortex completely with resultant fluid waves seen.\n"
            "7. Ultrasonic phaco-tip was inserted into the anterior chamber and applicated for cataract extraction with assistance of a phaco-chopper. Side-port was made with 7515 bever knife over limbus.\n"
            "8. Residual cortex was removed from the capsular bag using IA.\n"
            "9. The viscoelastic material was injected into the anterior chamber and sulcus.\n"
            "10. A posterior chamber intraocular lens was inserted into the bag.\n"
            "11. Viscoelastic material was washed out with simcoe irrigator. The wound was closed with stromal hydration.\n"
            "12.The patient stood the whole procedure well.\n\n"
            "Complication: $COL_COMPLICATIONS\n"
            "$DETAILS_OF_IOL",
        "TEMPLATE_ECCE":"TEMPLATE_ECCE_TEST"
    },
    '3': {
        "BOT":"CATA",
        "SERVICE_FILE":"vghbot-5fe0aba1d3b9.json",
        "SPREADSHEET": "OP-4081陳克華",
        "WORKSHEET": "OP",
        "VS_NAME": "陳克華",
        "VS_CODE": 4081,
        "R_NAME": "鄭明軒",
        "R_CODE": 4123,
        "DATE_PATTERN": "%Y/%m/%d",
        "COL_DATE": "時間",
        "COL_HISNO": "病歷號",
        "COL_NAME": "姓名",
        "COL_DIAGNOSIS": "診斷",
        "COL_SIDE": "側別",
        "COL_OP": "手術",
        "COL_LENSX": "LenSx",
        "COL_IOL": "IOL",
        "COL_FINAL": "Final",
        "COL_SN": "SN",
        "COL_COMPLICATIONS": "Complications",
        "TEMPLATE":
            '1. After retrobulbar anesthesia, the $TRANSFORMED_SIDE eye was sterilized and draped.\n'
            '2. Moderated and severe lens opacity was identified under microscopy through illumination.\n'
            '3. The eyelids were separated with a specular. A temporal corneal incision was made with phaco knife and then viscoelastic material was injected to the anterior chamber.\n'
            '4. Continous curvilinear anterior capsulorrhexis was performed with a bent 26 gauge needle and capsular forceps. And then hydrodissection of the lens capsule from cortex completely with resultant fluid waves seen.\n'
            '5. Ultrasonic phaco-tip was inserted into the anterior chamber. Side port was made the 7515 bever knife over limbus and then viscoelastic material was injected to the anterior chamber. Phaco-tip was applicated for cataract extraction with assistance of a phaco-chopper.\n'
            '6. Residual cortex was removed from the capsular bag using an IA.\n'
            '7. Viscoelastic material was injected into the anterior chamber and capsular bag.\n'
            '8. A posterior chamber intraocular lens was inserted into the capsular bag.\n'
            '9. Viscoelastic material was aspirated with Simcore irrigator.\n'
            '10.Miostate was injected into the anterior chamber to constrict the pupil and was then washed out. Stroma hydration was performed with BSS.\n'
            '11.An inferior fornix subconjunctival injection of decadron 0.4 ml and gentamyc in 0.4ml was given.\n'
            '12.The patient stood the whole procedure well.\n\n'
            'Complication: $COL_COMPLICATIONS\n'
            'DETAILS_OF_IOL',
        "TEMPLATE_LENSX":
            "1. After local anesthesia, Continuous curvilinear anterior capsulorrhexis of $TRANSFORMED_SIDE eye was performed with Lensx Femtosecond laser.\n"
            "2. Patient's eye was sterilized and draped as regular routine.\n"
            "3. Moderated and severe lens opacity was identified under microscopy through illumination.\n"
            "4. The upper and lower eyelids were opened with eyelid speculum.\n"
            "5. A temporal corneal incision was made with phaco-knife and then viscoelastic material was injected to the anterior chamber.\n"
            "6. Hydrodissection of the lens capsule from cortex completely with resultant fluid waves seen.\n"
            "7. Ultrasonic phaco-tip was inserted into the anterior chamber and applicated for cataract extraction with assistance of a phaco-chopper. Side-port was made with 7515 bever knife over limbus.\n"
            "8. Residual cortex was removed from the capsular bag using IA.\n"
            "9. The viscoelastic material was injected into the anterior chamber and sulcus.\n"
            "10. A posterior chamber intraocular lens was inserted into the bag.\n"
            "11. Viscoelastic material was washed out with simcoe irrigator. The wound was closed with stromal hydration.\n"
            "12.The patient stood the whole procedure well.\n\n"
            "Complication: $COL_COMPLICATIONS\n"
            "DETAILS_OF_IOL",
    },
}
import json
with open("config_notebot.json","w") as f:
    json.dump(config_dict, fp=f, ensure_ascii=False, indent=4)
    
'''
