import pygsheets
import pathlib
import json

SA_FILENAME_REGEX = 'vghbot*.json'
GSHEET_SPREADSHEET = 'config_vghbot'
# vghbot_note_op 會用
GSHEET_WORKSHEET_SURGERY = 'set_surgery'
GSHEET_WORKSHEET_IVI = 'set_ivi'
GSHEET_WORKSHEET_TEMPLATE_OPNOTE = 'template_opnote'
# vghbot_opd 會用
GSHEET_WORKSHEET_ACC = 'account'
GSHEET_WORKSHEET_DRUG = 'opd_drug'
GSHEET_WORKSHEET_OVD = 'opd_ovd'
GSHEET_WORKSHEET_IOL = 'opd_iol'
GSHEET_WORKSHEET_CONFIG = 'config'

SERVICE_ACCOUNT_JSON = ""

class GsheetClient:
    def __init__(self, client_secret='', service_account_file=None, service_account_env_var=None, service_account_json=SERVICE_ACCOUNT_JSON) -> None:
        '''
        認證尋找順序: service_account_json >> service_account_env_var >> service_account_file >> 
        在同目錄下透過SA_FILENAME_REGEX搜尋的service_account_file >> client_secret >> client_secret預設的搜尋路徑
        '''
        p = pathlib.Path() # 此檔案的目錄
        sa_files = list(p.glob(SA_FILENAME_REGEX)) # 尋找符合REGEX檔案
        
        try:
            if service_account_json is not None and service_account_json.strip()!='':
                self.client = pygsheets.authorize(service_account_json=service_account_json)
            elif service_account_env_var is not None:
                self.client = pygsheets.authorize(service_account_env_var=service_account_env_var)
            elif service_account_file is not None:
                self.client = pygsheets.authorize(service_account_file=service_account_file)
            elif service_account_file is None and len(sa_files)!=0:
                self.client = pygsheets.authorize(service_account_file=sa_files[0])
            elif client_secret != '':
                self.client = pygsheets.authorize(client_secret=client_secret)
            else:
                print("Not input any authorize parameters => seeking client_secret.json as default")
                self.client = pygsheets.authorize()
        except Exception as e:
            print(f"Authorize failed!:{e}")


    def get_df(self, spreadsheet, worksheet, format_string=True, column_uppercase=False):
        '''
        取得gsheet資料並轉成dataframe, 預設是全部轉成文字形態(format_string=True)處理
        Input: spreadsheet, worksheet, format_string=True, column_uppercase=False
        '''
        ssheet = self.client.open(spreadsheet)
        wsheet = ssheet.worksheet_by_title(worksheet)
        df = wsheet.get_as_df(has_header=True, include_tailing_empty=False, numerize=False) 
        # has_header=True: 第一列當作header; numerize=False: 不要轉化數字; include_tailing_empty=False: 不讀取每一row後面空白Column資料
        if format_string:
            df = df.astype('string') #將所有dataframe資料改成用string格式處理，新的格式比object更精準
        if column_uppercase:
            df.columns = df.columns.str.upper() #將所有的columns name改成大寫 => case insensitive
        return df


    def get_df_select(self, spreadsheet, worksheet, format_string=True, column_uppercase=False):
        '''
        透過get_df取得gsheet資料並轉成dataframe
        並讓使用者選擇指定的row number，(-)表示連續範圍 (,)做分隔
        '''
        # get_df取得gsheet資料
        df = self.get_df(spreadsheet=spreadsheet, worksheet=worksheet, format_string=format_string, column_uppercase=column_uppercase)
        
        # 使用者輸入選擇
        row_list_input = input("(-)表示連續範圍 (,)做分隔\n請選擇gsheet列碼: ")
        row_set = set()
        for i in row_list_input.split(','):
            if len(i.split('-')) > 1:  # 如果有範圍標示，把這段範圍加入set
                row_set.update(range(int(i.split('-')[0]), int(i.split('-')[1]) + 1))
            else:
                row_set.add(int(i))

        row_list = [row - 2 for row in row_set]  
        # google spreadsheet的index是從1開始，所以df對應的相差一，此外還有一欄變成df的column headers，所以總共要減2
        
        row_list.sort()
        return df.iloc[row_list, :]


    def get_col_dict(self, spreadsheet, worksheet):
        '''
        取得gsheet資料並轉成一個由欄位名為key、欄位資料為values list的格式,
        資料預設是全部轉成文字形態(format_string=True)處理，""case_sensitive""模式
        '''
        result_dict = {}
        df = self.get_df(spreadsheet = spreadsheet, worksheet = worksheet, column_uppercase=False)

        # 將 column內部有的空白cell清除
        column_names = df.columns
        for column in column_names:
            final_list = []
            tmp_list = df[column].to_list()
            for item in tmp_list:
                if item == '' or item.strip()=='':
                    continue
                final_list.append(item)
            result_dict[column] = final_list
        return result_dict


    def list_spreadsheet(self):
        result = self.client.spreadsheet_titles()
        # print(result)
        return result
        

    def list_worksheet(self, spreadsheet):
        ssheet = self.client.open(spreadsheet)
        result = []
        for wsheet in ssheet:
            result.append(wsheet.title)
        # print(result)
        return result
        