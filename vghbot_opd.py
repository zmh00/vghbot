import uiautomation as auto
import os
import time
import threading # for hotkey
from threading import Event # for hotkey
# import multiprocessing
import inspect
from ctypes import windll
import json
import sys
import subprocess
import pygsheets
import pandas
import datetime

import gsheet

# 製作EXE => 需要加上google json

# FIXME 目前沒用
# TODO 加入 window偵測方式應該能提升效率
class find_new_window():
    def __init__(self, pid, exclude_hwnd_set: set=None) -> None:
        self.hwnd_set = set()
        self.pid = pid
        if exclude_hwnd_set is None:
            self.exclude_hwnd_set = set()
    def reset(self):
        self.hwnd_set = set()
    def refind(self) -> list:
        new_hwnd_set = set(window_search_pid(self.pid, recursive=True, return_hwnd=True)) # 目前使用回傳handles來處理，未來可以考慮直接傳control元件，但關於set的運算會不會出問題?
        auto.Logger.WriteLine(f"NEW_HWND_SET: {new_hwnd_set}", auto.ConsoleColor.Yellow)
        diff = new_hwnd_set-self.hwnd_set-self.exclude_hwnd_set # 可以排除特定的視窗
        auto.Logger.WriteLine(f"DIFF_HWND_SET: {diff}", auto.ConsoleColor.Yellow)
        self.hwnd_set = new_hwnd_set
        return [auto.ControlFromHandle(d) for d in diff]

# TODO 把invoke丟到thread無法繞過彈窗錯誤


# ==== 基本操作架構

def process_exists(process_name):
    '''
    Check if a program (based on its name) is running
    Return yes/no exists window and its PID
    '''
    call = 'TASKLIST', '/FI', 'imagename eq %s' % process_name
    # use buildin check_output right away
    output = subprocess.check_output(call).decode('big5')  # 在中文的console中使用需要解析編碼為big5
    output = output.strip().split('\r\n')
    if len(output) == 1:  # 代表只有錯誤訊息
        return False, 0
    else:
        # check in last line for process name
        last_line_list = output[-1].lower().split()
    return last_line_list[0].startswith(process_name.lower()), last_line_list[1]


def process_responding(name):
    """Check if a program (based on its name) is responding"""
    cmd = 'tasklist /FI "IMAGENAME eq %s" /FI "STATUS eq running"' % name
    status = subprocess.Popen(cmd, stdout=subprocess.PIPE).stdout.read()
    status = str(status).lower() # case insensitive
    return name in status

# FIXME 目前沒用
def process_responding_PID(pid):
    """Check if a program (based on its PID) is responding"""
    cmd = 'tasklist /FI "PID eq %d" /FI "STATUS eq running"' % pid
    status = subprocess.Popen(cmd, stdout=subprocess.PIPE).stdout.read()
    status = str(status).lower()
    return str(pid) in status


def captureimage(control = None, postfix = ''):
    auto.Logger.WriteLine('CAPTUREIMAGE INITIATION', auto.ConsoleColor.Yellow)
    if control is None:
        c = auto.GetRootControl()
    else:
        c = control
    if postfix == '':
        path = f"{datetime.datetime.today().strftime('%Y%m%d_%H%M%S')}.png"
    else:
        path = f"{datetime.datetime.today().strftime('%Y%m%d_%H%M%S')}_{postfix}.png"
    c.CaptureToImage(path)

# FIXME 目前沒用
def window_search_pid(pid, search_from=None, maxdepth=1, recursive=False, return_hwnd=False):
    '''
    尋找有processID==pid視窗, 並從search_from往下找, 深度maxdepth
    recursive=True會持續往下找(DFS遞迴)
    return_hwnd=True會回傳NativeWindowHandle list
    '''
    target_list = []
    if search_from is None:
        search_from = auto.GetRootControl()
    for control, depth in auto.WalkControl(search_from, maxDepth=maxdepth):
        if (control.ProcessId == pid) and (control.ControlType==auto.ControlType.WindowControl):
            target_list.append(control)
            if recursive is True:
                target_list.extend(window_search_pid(pid, control, maxdepth=1, recursive=True))

    auto.Logger.WriteLine(f"From {search_from} MATCHED WINDOWS [PID={pid}]: {len(target_list)}", auto.ConsoleColor.Yellow)
    if return_hwnd:
        return [t.NativeWindowHandle for t in target_list]
    else:
        return target_list


def window_search(window, retry=5, topmost=False):  
    '''
    找尋傳入的window物件重覆retry次, 找到後會將其取得focus和可以選擇是否topmost, 若找不到會常識判斷其process有沒有responding
    retry<0: 無限等待 => 等待OPD系統開啟用
    '''
    # TODO 可以加上判斷物件是否IsEnabled => 這樣可以防止雖然找得到視窗或物件但其實無法對其操作
    _retry = retry
    try:
        while retry != 0:
            if window.Exists():
                auto.Logger.WriteLine(
                    f"{inspect.currentframe().f_code.co_name}|Window found: {window.GetSearchPropertiesStr()}", auto.ConsoleColor.Yellow)
                window.SetActive()  # 這有甚麼用??
                window.SetTopmost(True)
                if topmost is False:
                    window.SetTopmost(False)
                window.SetFocus()
                return window
            else:
                if process_responding(CONFIG['PROCESS_NAME']):
                    auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Window not found: {window.GetSearchPropertiesStr()}", auto.ConsoleColor.Red)
                    retry = retry-1
                else:
                    auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Process not responding", auto.ConsoleColor.Red)
                time.sleep(0.2)
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Window not found(after {_retry} times): {window.GetSearchPropertiesStr()}", auto.ConsoleColor.Red)
        captureimage(postfix=inspect.currentframe().f_code.co_name)
        return None
    except Exception as err:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Something wrong unexpected: {window.GetSearchPropertiesStr()}", auto.ConsoleColor.Red)
        print(err) # TODO remove in the future
        captureimage(postfix=inspect.currentframe().f_code.co_name)
        return window_search(window, retry=retry) # 目前使用遞迴處理 => 會無窮迴圈後續要考慮新方式 # TODO


def datagrid_list_pid(pid):  # 在PID框架下取得任意畫面下的所有datagrid，如果不指定PID就列出全部嗎?
    target_win = []
    for win in auto.GetRootControl().GetChildren(): # TODO 這段可以用來監測pop window
        if win.ProcessId == pid:
            target_win.append(win)
    auto.Logger.WriteLine(f"MATCHED WINDOWS(PID={pid}): {len(target_win)}", auto.ConsoleColor.Yellow)
    target_datagrid = []
    for win in target_win:
        for control, depth in auto.WalkControl(win, maxDepth=2):
            if control.ControlType == auto.ControlType.TableControl and control.Name == 'DataGridView':
                target_datagrid.append(control)
    if len(target_datagrid)==0:
        auto.Logger.WriteLine(f"NO DATAGRID RETRIEVED", auto.ConsoleColor.Red)
    return target_datagrid


def datagrid_values(datagrid, column_name=None, retry=5):
    '''
    Input: datagrid, column_name=None, retry=5 
    指定datagrid control並且取得內部所有values
    (資料列需要有'；'才會被收錄且會將(null)轉成'')
    # TODO: 未來考慮需不需要retry section?
    '''
    # 處理datagrid下完全沒有項目 => 回傳空list
    children = datagrid.GetChildren()
    if len(children) == 0:
        print(f"Datagrid({datagrid.AutomationId}): No values in datagrid")
        return []
    
    # retry section
    while (retry > 0):
        children = datagrid.GetChildren()
        if children[-1].Name == '資料列 -1':
            auto.Logger.WriteLine("Datagrid retrieved failed", auto.ConsoleColor.Red)
            datagrid.Refind()  # TODO 測試有效嗎? 可能要TopLevelControl那個level refind | 可以使用control.GetTopLevelControl()還是應該考慮傳進該control的top level
            retry = retry - 1
            continue
        else:
            break

    # get corresponding column_index
    column_index=None
    if column_name is not None:
        for i in children:
            if i.Name == '上方資料列':
                tmp = i.GetChildren()
                for index, j in enumerate(tmp):
                    if j.Name == column_name:
                        column_index = index
                        break
                break
    
    # parsing
    value_list = []
    for item in children:
        if TEST_MODE:
            print(f"Datagrid({datagrid.AutomationId}):{value}")
        value = item.GetLegacyIAccessiblePattern().Value
        if ';' in value:  # 有資料的列的判斷方式
            if column_index is not None:
                t = value.replace('(null)', '').split(';')[column_index]
                t = t.strip() # 把一些空格字元清掉
                value_list.append(t)
            else:
                t = value.replace('(null)', '').split(';') # 把(null)轉成''
                value_list.append([cell.strip() for cell in t])  # 把每個cell內空格字元清掉
    return value_list


def datagrid_search(search_text: list, datagrid, column_name=None, retry=5, only_one=True):
    '''
    Search datagrid based on search_text, each search_text can only be matched once(case insensitive), return the list of all the matched item
    search_text, 可以一次傳入要在此datagrid搜尋的資料陣列
    column_name=None, 指定column做搜尋
    retry=5, 預設重覆搜尋5次
    only_one=True, 找到符合一個target item就回傳 => 增加效率，但同時要搜尋多筆應設定only_one=False
    '''
    if type(search_text) is not list:
        search_text = [search_text]
    else:
        search_text = list(search_text) # 複製一個list目的是怕後面的pop影響原本傳入的參數
    
    target_list = []

    # FIXME 重新整理有效嗎? 如果沒有children似乎沒處理
    # 目前是針對datagrid去獲取children(row Control) => 如果children資料異常就會重新整理
    # 目前有時候會出現datagrid.Getchildren後會取到"資料列-1" => 並沒有這行，導致後續無法操作，視窗.refind()一次就有機會正常
    while (retry > 0):
        children = datagrid.GetChildren()
        if children[-1].Name == '資料列 -1': # 資料獲取有問題
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Datagrid retrieved failed", auto.ConsoleColor.Red)
            datagrid.Refind()  # TODO 測試有效嗎? 可能要TopLevelControl那個level refind
            retry = retry - 1
            continue
        else:
            break
    
    # children是datagrid下每個資料列，每個資料列下是每格column資料
    # 如果針對每個資料列下去搜尋特定column的control，再取值比對感覺效率差
    
    # 針對column name先取得對應的column index，後續再利用；做定位
    column_index=None
    if column_name is not None:
        for i in children:
            if i.Name == '上方資料列':
                tmp = i.GetChildren()
                for index, j in enumerate(tmp):
                    if j.Name == column_name:
                        column_index = index
                        break
                break
    
    already_list = []
    for item in children:
        value = item.GetLegacyIAccessiblePattern().Value
        if TEST_MODE:
            print(f"Datagrid({datagrid.AutomationId})|{search_text}:{value}")
        if '資料列' in item.Name: # 有資料的列才做判斷
            match = value.lower() # case insensitive
            if column_index is not None: # 有找到Column
                match = match.replace('(null)', '').split(';') # 將(null)轉成空字串''並且透過;分隔欄位資訊
                if column_index >= len(match):
                    continue
                else:
                    match = match[column_index]
            for text in search_text: # 用每一個serch_text去配對該row的資訊
                if text in already_list: # 讓搜尋字串只匹配一次
                    continue
                if text.lower() in match:
                    target_list.append(item)
                    already_list.append(text)  # 讓搜尋字串只匹配一次
                    if TEST_MODE:
                        print(f"Datagrid found: {text}")
                    break
            if (only_one == True) and len(target_list)>0:
                break
    return target_list
    

def click_blockinput(control, doubleclick=False, simulateMove=False):
    try:
        res = windll.user32.BlockInput(True)
        if TEST_MODE:
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|BLOCKINPUT START: {res}")
            print(f"DEBUG:{control.Name}:{control.GetClickablePoint()}") # FIXME
        if doubleclick:
            control.DoubleClick(waitTime=0.1, simulateMove=simulateMove)
        else:
            control.Click(waitTime=0.1, simulateMove=simulateMove)
    except Exception as e:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Blockinput&Click Failed: {e}", auto.ConsoleColor.Red) # TODO 如果因為物件不存在而沒點擊到，不會跳出exception，但會有error message=>抓response??
        return False
    res = windll.user32.BlockInput(False)
    if TEST_MODE:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|BLOCKINPUT END: {res}")
    return True


def click_retry(control, topwindow = None, retry=5, doubleclick=False):
    _retry = retry
    topwindow = control.GetTopLevelControl() # TODO　這行有意義嗎?
    # print(f"TOPWINDOW:{topwindow}")
    while (1):
        if retry <= 0: # 嘗試次數用完
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT CLICKABLE(after retry:{_retry}): {control.GetSearchPropertiesStr()}", auto.ConsoleColor.Red)
            break
            # return # 把return拿掉是為了至少按一次 => 因為有時候GetClickablePoint()是false但可以按
        if not control.Exists():
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT EXIST(CONTROL): {control.GetSearchPropertiesStr()}", auto.ConsoleColor.Red)
            if not topwindow.Exists():
                auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT EXIST(TOPWINDOW): {control.GetSearchPropertiesStr()}", auto.ConsoleColor.Red)
            time.sleep(1)
            retry = retry - 1
            continue
        elif control.BoundingRectangle.width() != 0 or control.BoundingRectangle.height() != 0:
            control.SetFocus()
            return click_blockinput(control, doubleclick)
        else:
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|SOMETHING WRONG: {control.GetSearchPropertiesStr()}", auto.ConsoleColor.Red)
            continue
    return False # 如果沒有成功點擊返回即return False


def click_datagrid(datagrid, target_list:list, doubleclick=False): # FIXME pop的位置怪怪的
    '''
    能在datagrid中點擊項目並使用scroll button
    成功完成回傳True, 失敗會回傳沒有點到的target_list
    '''
    if len(target_list) == 0: # target_list is empty
        return True
    
    scroll = datagrid.ScrollBarControl(searchDepth=1, Name="垂直捲軸")
    downpage = scroll.ButtonControl(searchDepth=1, Name="向下翻頁")
    auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|CLICKING DATAGRID: total {len(target_list)} items", auto.ConsoleColor.Yellow)
    remaining_target_list = target_list.copy()
    if downpage.Exists():
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|EXIST: scroll button", auto.ConsoleColor.Yellow)
        while True:
            for t in target_list:
                if t in remaining_target_list:
                    if t.BoundingRectangle.width() != 0 or t.BoundingRectangle.height() != 0:
                        t.SetFocus()
                        if click_blockinput(t, doubleclick=doubleclick):
                            remaining_target_list.remove(t)
            if len(remaining_target_list) == 0: # remaining_target_list is empty
                return True
            if downpage.Exists():
                downpage.GetInvokePattern().Invoke()
                auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|====DOWNPAGE====", auto.ConsoleColor.Yellow)
            else: # downpage按到最底了
                if len(remaining_target_list)!=0:
                    auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|ITEM NOT FOUND:{[j.Name for j in remaining_target_list]}", auto.ConsoleColor.Red)
                return remaining_target_list #沒點到的回傳list
    else:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT EXIST: scroll button", auto.ConsoleColor.Yellow)
        for t in target_list:
            t.SetFocus()
            if click_blockinput(t, doubleclick=doubleclick):
                remaining_target_list.remove(t)
        if len(remaining_target_list)!=0:
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|ITEM NOT FOUND:{[j.Name for j in remaining_target_list]}", auto.ConsoleColor.Red)
            return remaining_target_list #沒點到的回傳list
        return True


def get_patient_data():
    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    if window_soap.Exists():
        l = window_soap.Name.split()
        p_dict = {
            'hisno': l[0],
            'name': l[1],
            'id': l[6], 
            'charge': l[5],
            'birthday': l[4][1:-1],
            'age': l[3][:2]
        }
        return p_dict
    else:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|No window frmSoap", auto.ConsoleColor.Red)
        return False

# FIXME
def check_state():  # TODO 判斷目前是哪個視窗狀態
    pass

# FIXME
def wait_for_manual_control(text):
    auto.Logger.WriteLine(f"[視窗異常]請自行操作視窗到指定視窗:{text}", auto.ConsoleColor.Cyan)
    while(1):
        choice = input('確認已操作至指定視窗?(y/n): ')
        if choice.lower().strip() == 'y':
            return True
        else:
            continue    

# ==== 門診系統操作函數
# 每個操作可以分幾個階段: start_point, main_body, end_point
def login(path, account: str, password: str, section_id: str, room_id: str):
    os.startfile(path)
    auto.Logger.WriteLine("Finished: Start the OPD system", auto.ConsoleColor.Yellow)

    # 等待載入
    main = auto.WindowControl(SubName="台北榮民總醫院", searchDepth=1)
    main = window_search(main, -1)
    if main is not None:
        auto.Logger.WriteLine("Finished: Loading OPD system", auto.ConsoleColor.Yellow)
    else:
        auto.Logger.WriteLine("Failed: Loading OPD system", auto.ConsoleColor.Red)
        return False

    # 跳出略過按鈕
    msg = main.ButtonControl(SubName="略過", Depth=2)
    msg.GetInvokePattern().Invoke()

    # 填入開診資料
    acc = main.EditControl(AutomationId="txtSignOnID", Depth=1)
    acc.GetValuePattern().SetValue(account)
    psw = main.EditControl(AutomationId="txtSignOnPassword", Depth=1)
    psw.GetValuePattern().SetValue(password)
    section = main.EditControl(AutomationId="1001", Depth=2)
    section.GetValuePattern().SetValue(section_id)
    room = main.EditControl(AutomationId="txtRoom", Depth=1)
    room.GetValuePattern().SetValue(room_id)
    signin = main.ButtonControl(AutomationId="btnSignon", Depth=1)
    click_retry(signin)  # 為何改用click是因為只要使用invoke若遇到popping window會報錯且卡死API

    # 判斷登入是否成功
    check_login = main.WindowControl(SubName="錯誤訊息", searchDepth=1)
    if check_login.Exists(maxSearchSeconds=1, searchIntervalSeconds=0.2):
        auto.Logger.WriteLine(f"Login failed!", auto.ConsoleColor.Red)
        return False

    # 處理登入非該診醫師的警告
    warning_msg = auto.WindowControl(searchDepth=2, AutomationId="FlaxibleMessage")
    warning_msg = window_search(warning_msg,1)
    if warning_msg is not None:
        button = warning_msg.ButtonControl(Depth=2, AutomationId="btnOK")
        if button.Exists():
            # button.GetInvokePattern().Invoke() # TODO 需要測試會不會卡住API
            click_retry(button, warning_msg)
        else:
            auto.Logger.WriteLine("No OK_Button under FlaxibleMessage", auto.ConsoleColor.Red)
    else:
        auto.Logger.WriteLine("No FlaxibleMessage", auto.ConsoleColor.Red)

    # 醫師待辦事項通知 => 可window.close()
    warning_msg = auto.WindowControl(AutomationId="dlgMessageCenter", searchDepth=2)
    warning_msg = window_search(warning_msg,1)
    if warning_msg is not None:
        warning_msg.GetWindowPattern().Close()
    else:
        auto.Logger.WriteLine("No dlgMessageCenter", auto.ConsoleColor.Red)

    # 醫事卡非登入醫師本人通知
    warning_msg = auto.WindowControl(AutomationId="dlgWarMessage", searchDepth=2)
    warning_msg = window_search(warning_msg,1)
    if warning_msg is not None:
        warning_msg.GetWindowPattern().Close()
    else:
        auto.Logger.WriteLine("No dlgWarMessage", auto.ConsoleColor.Red)
    
    # 以下目標一次處理所有dialog(COVID警告通知+此診目前無掛號)，因為dialog不是獨立的，在walkatree內會有階層關係，但不是每一個都有enabled可以被操作，要按照順序處理
    # 此診目前無掛號 會是第一個跳出的dialog，但其他訊息後續出現後，此dialog會被設成not enabled => 導致不能被選中關閉
    time.sleep(0.5)
    auto.SendKeys("{SPACE}" * 3)
    auto.SendKeys("{SPACE}" * 3)

    # # COVID警告通知 => 可close
    # warning_msg = auto.WindowControl(Name="訊息", searchDepth=2)
    # warning_msg = search_window(warning_msg,1)
    # if warning_msg is not None:
    #     warning_msg.GetWindowPattern().Close()
    # else:
    #     auto.Logger.WriteLine("No 訊息", auto.ConsoleColor.Red)
    
    # 處理警告:(SOAP的填答通知、此診目前無掛號) 
    # TODO 取得該process ID下的top level window 去送space
    # TODO 用一個thread持續監測新的window且同樣process ID，去關掉這些window? => 可以先嘗試列印出來看可不可以多線程運行


def login_change(account: str, password: str, section_id: str, room_id: str):
    window_main = auto.WindowControl(searchDepth=1, SubName='台北榮民總醫院', AutomationId="frmPatList")
    window_main = window_search(window_main)
    if window_main is None:
        auto.Logger.WriteLine("No window frmPatList", auto.ConsoleColor.Red)
        return False

    c_menubar = window_main.MenuBarControl(searchDepth=1, AutomationId="MenuStrip1")
    c_f1 = c_menubar.MenuItemControl(searchDepth=1, Name='輔助功能')
    c_login_change = c_f1.MenuItemControl(searchDepth=1, Name='換科(診)登入')
    click_retry(c_f1)
    time.sleep(0.05)
    click_retry(c_login_change)
    # res = c_login_change.GetExpandCollapsePattern().Expand(waitTime=10) #榮總menubar不支援expand and collapse pattern
    # c_login_change.GetInvokePattern().Invoke() #會造成卡住API所以改用click

    window_relog = auto.WindowControl(searchDepth=2, AutomationId="dlgDCRRelog")
    window_relog = window_search(window_relog)
    if window_relog is not None:
        auto.Logger.WriteLine("Window Relog Exists", auto.ConsoleColor.Yellow)
        window_relog.EditControl(searchDepth=1, AutomationId="tbxUserID").GetValuePattern().SetValue(account)
        window_relog.EditControl(searchDepth=1, AutomationId="tbxUserPassword").GetValuePattern().SetValue(password)
        window_relog.ComboBoxControl(searchDepth=1, AutomationId="cbxSectCD").GetValuePattern().SetValue(section_id)
        window_relog.EditControl(searchDepth=1, AutomationId="tbxRoomNo").GetValuePattern().SetValue(room_id)
        btn = window_relog.ButtonControl(searchDepth=1, AutomationId="btnSignOn")
        click_retry(btn)

        # FIXME 這邊似乎會卡住而造成後續API出問題但動一下滑鼠似乎能解決????
        window_main.Click() # FIXME 測試看看有沒有用
        time.sleep(1)
        # 以下嘗試都失敗
        # btn.GetInvokePattern().Invoke()
        # process = multiprocessing.Process(target=process_invoke, args=(btn,))
        # process.start()
        # th = threading.Thread(target=thread_invoke, args=(btn,))
        # th.start()

        # 換到班表沒有匹配的診會出現警告彈窗 => 造成API卡住
        # TODO 在沒有實際存在此視窗下 就算前面用click 也出現API卡住，但等err過去可以繼續使用
        message = window_relog.WindowControl(searchDepth=1, AutomationId="FlaxibleMessage")
        message = window_search(message, 1)
        if message is not None:
            message.ButtonControl(Depth=2, AutomationId="btnOK").GetInvokePattern().Invoke()
        else:
            auto.Logger.WriteLine("No FlaxibleMessage under window_relog", auto.ConsoleColor.Red)
            message = auto.WindowControl(searchDepth=2, AutomationId="FlaxibleMessage")
            message = window_search(message, 1)
            if message is not None:
                message.ButtonControl(Depth=2, AutomationId="btnOK").GetInvokePattern().Invoke()
            else:
                auto.Logger.WriteLine("No FlaxibleMessage under root", auto.ConsoleColor.Red)
                return False

        # 換到沒有掛號的診會跳出無掛號的訊息視窗 => 可以用空白鍵解決
        auto.SendKeys("{SPACE}" * 3)
        auto.SendKeys("{SPACE}" * 3)
    else:
        auto.Logger.WriteLine("No Window Relog", auto.ConsoleColor.Red)


def appointment(hisno_list: list[str], skip_checkpoint = False):
    if type(hisno_list) is not list:
        hisno_list = [hisno_list]
    else:
        hisno_list = list(hisno_list)

    # start_point check
    if skip_checkpoint:
        window_main = auto.WindowControl(searchDepth=1, SubName='台北榮民總醫院', AutomationId="frmPatList")
        window_main = window_search(window_main)
        if window_main is None:
            return False
    else:
        while(1):
            window_main = auto.WindowControl(searchDepth=1, SubName='台北榮民總醫院', AutomationId="frmPatList")
            window_main = window_search(window_main)
            if window_main is None:
                auto.Logger.WriteLine("No window frmPatList", auto.ConsoleColor.Red)
                wait_for_manual_control(window_main.GetSearchPropertiesStr())
            else:
                break

    # get patient list => 減少重複掛號
    datagrid_patient = window_main.TableControl(searchDepth=1, SubName='DataGridView', AutomationId="dgvPatsList")
    patient_list = datagrid_values(datagrid=datagrid_patient, column_name='病歷號')

    for hisno in hisno_list:
        if hisno in patient_list: # 已經有病歷號了
            auto.Logger.WriteLine(f"Appointment exists: {hisno}", auto.ConsoleColor.Yellow)
            continue

        c_menubar = window_main.MenuBarControl(searchDepth=1, AutomationId="MenuStrip1")
        c_appointment = c_menubar.MenuItemControl(searchDepth=1, SubName='非常態掛號')
        click_retry(c_appointment) #為了防止popping window遇到invoke pattern會卡住

        # 輸入資料
        window_appointment = auto.WindowControl(searchDepth=2, AutomationId="dlgVIPRegInput")
        window_appointment = window_search(window_appointment)
        if window_appointment is None:
            auto.Logger.WriteLine("No window dlgVIPRegInput",auto.ConsoleColor.Red)
            return False
        # 找到editcontrol
        c_appoint_edit = window_appointment.EditControl(searchDepth=1, AutomationId="tbxIDNum")
        c_appoint_edit.GetValuePattern().SetValue(hisno)
        # 送出資料
        c_button_ok = window_appointment.ButtonControl(AutomationId="OK_Button")
        click_retry(c_button_ok) 
        #c_button_ok.GetInvokePattern().Invoke() # 如果後續跳出重覆掛號的dialog就會造成這步使用invoke會卡住

        # 判斷是否重覆掛號
        winmsg = window_appointment.WindowControl(searchDepth=1, Name="訊息")
        winmsg = window_search(winmsg, 1)
        if winmsg is not None:
            auto.Logger.WriteLine(f"Appointment[{hisno}] already exists",auto.ConsoleColor.Cyan)
            winmsg.GetWindowPattern().Close()


def retrieve(hisno):
    '''
    取暫存功能
    '''
    window_main = auto.WindowControl(searchDepth=1, SubName='台北榮民總醫院', AutomationId="frmPatList")
    window_main = window_search(window_main)
    if window_main is None:
        auto.Logger.WriteLine("No window frmPatList", auto.ConsoleColor.Red)
        return False

    # select病人
    c_datagrid_patient = window_main.TableControl(searchDepth=1, SubName='DataGridView', AutomationId="dgvPatsList")
    patient = datagrid_search([hisno], c_datagrid_patient)
    if len(patient)==0:
        auto.Logger.WriteLine(f"NOT EXIST PATIENT: {hisno}", auto.ConsoleColor.Red)
    else:
        click_datagrid(c_datagrid_patient, patient)
    # 按下取暫存按鍵
    c = window_main.ButtonControl(searchDepth=1, AutomationId="btnPatsTemp")
    click_retry(c)

    # 處理TOCC警告 => 這應該是隨機彈窗 會造成invoke使用後錯誤
    window_tocc = auto.WindowControl(searchDepth=2, AutomationId="dlgNewTOCC")
    if window_tocc.Exists(maxSearchSeconds=2, searchIntervalSeconds=0.2):
        window_tocc.CheckBoxControl(Depth=2, AutomationId="ckbAllNo").GetTogglePattern().Toggle()
        window_tocc.ButtonControl(Depth=2, AutomationId="btnOK").GetInvokePattern().Invoke()
    else:
        auto.Logger.WriteLine(f"NOT EXIST: window_tocc", auto.ConsoleColor.Yellow)


def ditto(hisno: str, skip_checkpoint = False): # TODO 改成hisno_list版本
    '''
    Ditto功能
    '''
    # start_point check
    if skip_checkpoint:
        window_main = auto.WindowControl(searchDepth=1, SubName='台北榮民總醫院', AutomationId="frmPatList")
        window_main = window_search(window_main)
        if window_main is None:
            return False
    else:
        while(1):
            window_main = auto.WindowControl(searchDepth=1, SubName='台北榮民總醫院', AutomationId="frmPatList")
            window_main = window_search(window_main)
            if window_main is None:
                auto.Logger.WriteLine("No window frmPatList", auto.ConsoleColor.Red)
                wait_for_manual_control(window_main.GetSearchPropertiesStr())
            else:
                break
    
    # select病人
    datagrid_patient = window_main.TableControl(searchDepth=1, SubName='DataGridView', AutomationId="dgvPatsList")
    patient = datagrid_search([hisno], datagrid_patient)
    if len(patient)==0:
        auto.Logger.WriteLine(f"NOT EXIST PATIENT: {hisno}", auto.ConsoleColor.Red)
    else:
        click_datagrid(datagrid_patient, patient, doubleclick=True)
    # 如果沒有點到該病人單純用select最後跳出的ditto資料會有錯誤
    # 對datagrid的病人資料使用doubleclick 也有ditto效果，另外也可以單點一下+按ditto按鈕
    # 按下ditto按鍵
    # c_ditto = window_main.ButtonControl(searchDepth=1, SubName='DITTO', AutomationId="btnPatsDITTO")
    # block_click(c_ditto)
        
    # 處理TOCC警告 => 這應該是隨機彈窗 會造成invoke使用後錯誤
    window_tocc = auto.WindowControl(searchDepth=2, AutomationId="dlgNewTOCC")
    if window_tocc.Exists(maxSearchSeconds=2, searchIntervalSeconds=0.2):
        window_tocc.CheckBoxControl(Depth=2, AutomationId="ckbAllNo").GetTogglePattern().Toggle()
        window_tocc.ButtonControl(Depth=2, AutomationId="btnOK").GetInvokePattern().Invoke()
    else:
        auto.Logger.WriteLine(f"NOT EXIST: window_tocc", auto.ConsoleColor.Yellow)

    # 健康行為登錄
    window_dlg = auto.WindowControl(searchDepth=2, AutomationId="dlgSMOBET")
    if window_dlg.Exists(maxSearchSeconds=2, searchIntervalSeconds=0.2):
        window_dlg.GetWindowPattern().Close()
    else:
        auto.Logger.WriteLine(f"NOT EXIST: dlgSMOBET", auto.ConsoleColor.Yellow)

    # 處理一堆警告視窗 => #TODO 我沒跳出這個 要保留彈性或是使用鍵盤處理?
    window_warn = auto.WindowControl(searchDepth=2, AutomationId="dlgWarMessage")
    if window_warn.Exists(maxSearchSeconds=2, searchIntervalSeconds=0.2):
        c_button_ok = window_warn.ButtonControl(searchDepth=1, AutomationId="OK_Button", SubName="繼續")
        c_button_ok.GetInvokePattern().Invoke()
    else:
        auto.Logger.WriteLine(f"NOT EXIST: window_warn", auto.ConsoleColor.Yellow)

    # TODO 以下拆開成另一個函數? 這樣可以幫助追蹤進度?
    # 進到ditto視窗
    window_ditto = auto.WindowControl(searchDepth=1, AutomationId="frmDitto")
    window_ditto = window_search(window_ditto)
    if window_ditto is None:
        auto.Logger.WriteLine("No window frmDitto", auto.ConsoleColor.Red)
        return False

    # 藥物過敏的視窗是在ditto視窗下
    window_allergy = window_ditto.WindowControl(searchDepth=1, AutomationId="dlgDrugAllergyDetailAndEdit")
    if window_allergy.Exists(maxSearchSeconds=1.0, searchIntervalSeconds=0.2):
        window_allergy.ButtonControl(Depth=3, SubName='無需更新', AutomationId="Button1").GetInvokePattern().Invoke()
    else:
        auto.Logger.WriteLine(f"NOT EXIST: window_allergy", auto.ConsoleColor.Yellow)

    # 進去選擇最近的一次眼科紀錄010, 110, 0PH, 1PH, 0C1,...?
    c_datagrid_ditto = window_ditto.TableControl(Depth=3, AutomationId="dgvPatDtoList")
    item = datagrid_search(CONFIG['SECTION_OPH'], c_datagrid_ditto, '科別')
    if len(item)==0:
        auto.Logger.WriteLine(f"NOT EXIST SECTIONS: {CONFIG['SECTION_OPH']}", auto.ConsoleColor.Red)
        return False
    click_datagrid(c_datagrid_ditto, item)

    # item.GetLegacyIAccessiblePattern().Select(2) # 這個select要有任一data row被點過才能使用，且只用select不會更新旁邊SOAP資料，要用Click!

    time.sleep(0.5)  # 怕執行太快
    # ditto 視窗右側
    c_text_s = window_ditto.EditControl(searchDepth=1, AutomationId="txtSOAP_S")
    if c_text_s.Exists(maxSearchSeconds=2.0, searchIntervalSeconds=0.2) and len(c_text_s.GetValuePattern().Value) > 0:
        # 選擇S、O copy selected
        c_check_s = window_ditto.CheckBoxControl(searchDepth=1, AutomationId="Check_S")
        c_check_s.GetTogglePattern().Toggle()
        c_check_o = window_ditto.CheckBoxControl(searchDepth=1, AutomationId="Check_O")
        c_check_o.GetTogglePattern().Toggle()
        c_check_a = window_ditto.CheckBoxControl(searchDepth=1, AutomationId="Check_A")
        c_check_a.GetTogglePattern().Toggle()
        c_check_p = window_ditto.CheckBoxControl(searchDepth=1, AutomationId="Check_P")
        c_check_p.GetTogglePattern().Toggle()
        window_ditto.ButtonControl(searchDepth=1, AutomationId="btnSelect").GetInvokePattern().Invoke()
    else:
        auto.Logger.WriteLine("txtSOAP_S is empty!", auto.ConsoleColor.Red)

    # TODO 處理慢簽視窗
    # 可以使用window.close()?


def select_package(index: int = -1, search_term: str = None):
    '''
    點擊組套功能(可以使用index[起始為0且需+3]或是用search term去搜尋組套視窗的項目)
    '''
    # Window SOAP 為起點
    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    window_soap = window_search(window_soap)
    if window_soap is None:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT EXIST WINDOW frmSoap", auto.ConsoleColor.Red)
        return False

    # Menubar
    c_menubar = window_soap.MenuBarControl(searchDepth=1, AutomationId="MenuStrip1")
    c_pkgroot = c_menubar.MenuItemControl(searchDepth=1, SubName='組套')
    c_pkgroot.GetInvokePattern().Invoke() # 這個可以使用invoke

    # 組套視窗
    window_pkgroot = auto.WindowControl(searchDepth=1, AutomationId="frmPkgRoot")
    window_pkgroot = window_search(window_pkgroot)
    c_datagrid_pkg = window_pkgroot.TableControl(searchDepth=1, AutomationId="dgvPkggroupPkg")
    
    # 使用的索引方式
    if index != -1: # 選擇指定index
        c_datalist_pkg = c_datagrid_pkg.GetChildren()
        # TODO 要不要加上如果找到資料 -1 要重新找?
        c_datalist_pkg[index].GetLegacyIAccessiblePattern().Select(2)
    elif search_term != None: # 選擇字串搜尋
        tmp_list = datagrid_search(search_text=search_term, datagrid=c_datagrid_pkg)
        if len(tmp_list)>0:
            tmp_list[0].GetLegacyIAccessiblePattern().Select(2)
        else:
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT EXIST {search_term} IN PACKAGE", auto.ConsoleColor.Red)
            return False
    else:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Wrong input", auto.ConsoleColor.Red)
        return False
    # 送出確認
    window_pkgroot.ButtonControl(searchDepth=1, AutomationId="btnPkgRootOK").GetInvokePattern().Invoke()


def select_iol_ovd(iol, ovd):
    '''
    select IOL and OVD
    '''
    # 尋找刀表IOL資訊的正式搜尋名稱
    iol_search_term, isNHI = gsheet_iol_search_term(iol)[iol]
    
    if isNHI:
        select_package(index=29)  # NHI IOL
    else:
        select_package(index=30)  # SP IOL
    
    # 組套第二視窗:frmPkgDetail window
    window_pkgdetail = auto.WindowControl(searchDepth=1, AutomationId="frmPkgDetail")
    window_pkgdetail = window_search(window_pkgdetail)
    c_datagrid_pkgorder = window_pkgdetail.TableControl(searchDepth=1, AutomationId="dgvPkgorder")
    
    # search_datagrid for target item
    target = []
    if c_datagrid_pkgorder.Exists():
        target = datagrid_search([iol_search_term, ovd], c_datagrid_pkgorder, only_one=False)
        if len(target) < 2:
            auto.Logger.WriteLine(f"LOSS OF RETURN: IOL:{iol}|OVD:{ovd}", auto.ConsoleColor.Red)
            auto.Logger.WriteLine(f"{[control.GetLegacyIAccessiblePattern().Value for control in target]}", auto.ConsoleColor.Red)
        # #### 分開搜尋的版本(效率較差但明確知道缺甚麼)
        # tmp = search_datagrid([iol], c_datagrid_pkgorder)
        # if len(tmp) == 0:
        #     auto.Logger.WriteLine(f"NO IOL: {iol}", auto.ConsoleColor.Red)
        # else:
        #     target.append(tmp[0])
        # tmp = search_datagrid([ovd], c_datagrid_pkgorder)
        # if len(tmp) == 0:
        #     auto.Logger.WriteLine(f"NO OVD: {ovd}", auto.ConsoleColor.Red)
        # else:
        #     target.append(tmp[0])
    else:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|No datagrid dgvPkgorder", auto.ConsoleColor.Red)
    
    # click_datagrid
    window_search(window_pkgdetail) #有需要嗎?
    c_datagrid_pkgorder.Refind() #有需要嗎?
    residual_list = click_datagrid(c_datagrid_pkgorder, target_list=target)

    # confirm
    window_pkgdetail.ButtonControl(searchDepth=1, AutomationId="btnPkgDetailOK").GetInvokePattern().Invoke()

    # 測試失敗: legacy.select
    # search_datagrid(c_datagrid_pkgorder, [iol])[0].GetLegacyIAccessiblePattern().Select(8)  # 無法被select不知道為何
    # search_datagrid(c_datagrid_pkgorder, [ovd])[0].GetLegacyIAccessiblePattern().Select(8)  # 無法被select不知道為何
    # c_datalist_pkgorder = c_datagrid_pkgorder.GetChildren()
    # c_datalist_pkgorder[3].GetLegacyIAccessiblePattern().Select(8)
    # c_datalist_pkgorder[8].GetLegacyIAccessiblePattern().Select(8)


def select_phaco_mode(mode=0):  
    '''
    選擇組套 => 依照有沒有LenSx區別
    mode=0|'nhi'
    mode=1|'lensx'
    '''    
    mode = str(mode).lower()
    if mode == "nhi" or  mode=='0':
        select_package(index = 31)  # 一般Phaco
    elif mode == "lensx" or mode=='1':
        select_package(index = 32)  # Lensx


def select_order_side_all(side: str = None): 
    # TODO 要能修改個別orders的側別和計價
    if side is None:
        while (1):
            side = input("Which side(1:R|2:L|3:B)? ").strip()
            if side == '1':
                side = 'R'
                break
            elif side == '2':
                side = 'L'
                break
            elif side == '3':
                side = 'B'
                break
    elif side.strip().upper() == 'OD':
        side = 'R'
    elif side.strip().upper() == 'OS':
        side = 'L'
    elif side.strip().upper() == 'OU':
        side = 'B'
    else:
        auto.Logger.WriteLine("UNKNOWN INPUT OF ORDER SIDE", auto.ConsoleColor.Red)
        return

    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    window_soap = window_search(window_soap)
    if window_soap is None:
        auto.Logger.WriteLine("No window frmSoap", auto.ConsoleColor.Red)
        return False
    
    # 修改order按鈕
    window_soap.ButtonControl(searchDepth=1, AutomationId="btnSoapAlterOrder").GetInvokePattern().Invoke()  
    # 進入修改window
    window_alterord = auto.WindowControl(searchDepth=1, AutomationId="dlgAlterOrd")
    window_alterord = window_search(window_alterord)
    
    # 在一個group下修改combobox
    group = window_alterord.GroupControl(searchDepth=1, AutomationId="GroupBox1")
    c_side = group.ComboBoxControl(searchDepth=1, AutomationId="cbxAlterOrdSpcnm").GetValuePattern().SetValue(side)

    # 按全選
    click_retry(window_alterord.ButtonControl(searchDepth=1, AutomationId="btnAOrdSelectAll"))

    # 點擊確認 => 不能用Invoke，且上面選擇項目後的click不能改變focus，否則選擇項目會被自動取消
    confirm = group.ButtonControl(searchDepth=1, AutomationId="btnAlterOrdOK")
    click_blockinput(confirm)
    # confirm.GetInvokePattern().Invoke() 
    # click_retry(confirm)

    # 點擊返回主畫面
    group.ButtonControl(searchDepth=1, AutomationId="btnAlterOrdReturn").GetInvokePattern().Invoke()


def select_eyedrop_side(drug_list, side):
    side = side.upper()
    if side not in ['OD','OS','OU']:
        auto.Logger.WriteLine("Wrong side input!", auto.ConsoleColor.Red)
        while (1):
            side = input("Which side(1:OD|2:OS|3:OU)? ").strip()
            if side == '1':
                side = 'OD'
                break
            elif side == '2':
                side = 'OS'
                break
            elif side == '3':
                side = 'OU'
                break
    for i in drug_list:
        if i['eyedrop'] == 1:
            i['route'] = side
    
    return drug_list


def select_prescription(drug_list): # TODO考慮拆成加藥+修改藥物頻次
    '''
    1. gsheet_drug: 先取得drug_list
    2. select_drug_side: 修改側別
    3. select_prescription: 加藥並修改
    '''
    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    window_soap = window_search(window_soap)
    if window_soap is None:
        auto.Logger.WriteLine("No window frmSoap", auto.ConsoleColor.Red)
        return False
    # 走藥物修改再加藥防止沒有診斷時不能加藥
    window_soap.ButtonControl(searchDepth=1, AutomationId="btnSoapAlterMed").GetInvokePattern().Invoke()  
    # 進入藥物修改window
    window_altermed = auto.WindowControl(searchDepth=1, AutomationId="dlgAlterMed")
    # 點擊加藥
    window_altermed.ButtonControl(searchDepth=1, AutomationId="btnDrugList").GetInvokePattern().Invoke()
    # 進入druglist window
    window_druglist = auto.WindowControl(searchDepth=1, AutomationId="frmDrugListExam")
    
    if len(drug_list) > 8:
        # raised exception or 分次處理
        auto.Logger.WriteLine("Too many drug at the same time!", auto.ConsoleColor.Red)
    
    # 輸入藥名: 搜尋框最多10個字元
    for i, drug in enumerate(drug_list):
        window_druglist.EditControl(AutomationId=f"TextBox{i}").GetValuePattern().SetValue(drug['name'][:10])
    # 搜尋按鈕
    window_druglist.ButtonControl(AutomationId="btnSearch").GetInvokePattern().Invoke()
    # 選擇datagrid內藥物項目
    c_datagrid_druglist = window_druglist.TableControl(Depth=3, AutomationId="dgvDrugList")
    target_list = datagrid_search([drug['name'] for drug in drug_list], c_datagrid_druglist, only_one=False)
    for i in target_list:
        click_retry(i)  # select可以選擇到欄位但要有點擊才能真的加藥

    # 點擊確認 => 選擇藥物的確認
    window_druglist.ButtonControl(Depth=3, AutomationId="btnAdd").GetInvokePattern().Invoke()

    # 修改藥物頻次
    window_altermed.Refind()  # TODO要處理error
    for drug in drug_list:
        c_modify = window_altermed.TabControl(searchDepth=1, AutomationId="TabControl1").PaneControl(searchDepth=1, AutomationId="TabPage1")
        # c_charge = c_modify.ListControl(searchDepth=1, AutomationId="ListBoxType").ListItemControl(SubName = "自購").GetSelectionItemPattern().Select() #FIXME 目前這樣使用會出錯，不知為何
        c_dose = c_modify.ComboBoxControl(searchDepth=1, AutomationId="ComboDose").GetValuePattern().SetValue(drug['dose'])
        c_freq = c_modify.ComboBoxControl(searchDepth=1, AutomationId="ComboFreq").GetValuePattern().SetValue(drug['frequency'])
        c_route = c_modify.ComboBoxControl(searchDepth=1, AutomationId="ComboRout").GetValuePattern().SetValue(drug['route'])
        c_duration = c_modify.ComboBoxControl(searchDepth=1, AutomationId="ComboDur").GetValuePattern().SetValue(drug['duration'])
        window_altermed.Refind()  # TODO要處理error
        c_datagrid_drugmodify = window_altermed.TableControl(searchDepth=1, Name="DataGridView")
        target_list = datagrid_search([drug['name']], c_datagrid_drugmodify)
        click_blockinput(target_list[0])
        click_blockinput(window_altermed.ButtonControl(searchDepth=1, AutomationId="btnModify"))
    # 只要對修改數據有任何input，選擇的datagrid就會跳成-1
    # 目前測試資料框使用setvalue或sendkey都可，但選擇藥物和送出都必須使用click()，流程必須是先設定好要更改的資料，再refind datagrid，然後點藥和送出都必須是click()
    # window_altermed.ButtonControl(searchDepth=1, AutomationId="btnModify").GetInvokePattern().Invoke()

    # 點擊返回主畫面
    window_altermed.ButtonControl(searchDepth=1, AutomationId="btnReturn").GetInvokePattern().Invoke()


def select_ivi(charge): # TODO
    charge = charge.upper()
    if charge == 'SP-A':
        data = get_patient_data()
        if '榮' in data['charge'] or '將' in data['charge']: # 榮民選擇
            select_package(index=33)
        else:
            select_package(index=34)
    elif charge == 'NHI' or charge == 'SP-1' or charge == 'SP-2' or charge == 'DRUG-FREE':
        select_package(index=35)
        # 組套第二視窗:frmPkgDetail window
        window_pkgdetail = auto.WindowControl(searchDepth=1, AutomationId="frmPkgDetail")
        window_pkgdetail = window_search(window_pkgdetail)
        c_datagrid_pkgorder = window_pkgdetail.TableControl(searchDepth=1, AutomationId="dgvPkgorder")

        target = datagrid_search(['Intravitreous'], c_datagrid_pkgorder)
        residual_list = click_datagrid(c_datagrid_pkgorder, target_list=target)
        # confirm
        window_pkgdetail.ButtonControl(searchDepth=1, AutomationId="btnPkgDetailOK").GetInvokePattern().Invoke()
    elif charge == 'ALL-FREE':
        pass
    else:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Wrong input", auto.ConsoleColor.Red)


def set_text(panel, text_input, location=0, replace=0):
    # panel = 's','o','p'
    # location=0 從頭寫入 | location=1 從尾寫入
    # replace=0 append | replace=1 取代原本的內容
    # 現在預設插入的訊息會換行
    # 門診系統解析換行是'\r\n'，如果只有\n會被忽視但仍可以被記錄 => 可以放入隱藏字元，不知道網頁版怎麼顯示?
    parameters = {
        's': ['PanelSubject', 'txtSoapSubject'],
        'o': ['PanelObject', 'txtSoapObject'],
        'p': ['PanelPlan', 'txtSoapPlan'],
    }
    panel = str(panel).lower()
    if panel not in parameters.keys():
        auto.Logger.WriteLine("Wrong panel in input_text",auto.ConsoleColor.Red)
        return False

    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    window_soap = window_search(window_soap)
    if window_soap is None:
        auto.Logger.WriteLine("No window frmSoap", auto.ConsoleColor.Red)
        return False

    edit_control = window_soap.PaneControl(searchDepth=1, AutomationId=parameters[panel][0]).EditControl(
        searchDepth=1, AutomationId=parameters[panel][1])
    if edit_control.Exists():
        text_original = edit_control.GetValuePattern().Value
        if replace == 1:
            text = text_input
        else:
            if location == 0:  # 從文本頭部增加訊息
                text = text_input+'\r\n'+text_original
            elif location == 1:  # 從文本尾部增加訊息
                text = text_original+'\r\n'+text_input
        try:
            edit_control.GetValuePattern().SetValue(text)  # SetValue完成後游標會停在最前面的位置
            # edit_control.SendKeys(text) # SendKeys完成後游標停在輸入完成的位置，輸入過程加上延遲有打字感，能直接使用換行(\n會自動變成\r\n)
            auto.Logger.WriteLine(
                f"{inspect.currentframe().f_code.co_name}|Input finished!", auto.ConsoleColor.Yellow)
        except:
            auto.Logger.WriteLine(
                f"{inspect.currentframe().f_code.co_name}|Input failed!", auto.ConsoleColor.Red)
        # TODO 需要考慮100行問題嗎?
    else:
        auto.Logger.WriteLine(
            f"{inspect.currentframe().f_code.co_name}|No edit control", auto.ConsoleColor.Red)
        return False

def set_S(text_input, location=0, replace=0):
    set_text('s', text_input, location, replace)

def set_O(text_input, location=0, replace=0):
    set_text('o', text_input, location, replace)

def set_P(text_input, location=0, replace=0):
    set_text('p', text_input, location, replace)

def get_text(panel):
    parameters = {
        's': ['PanelSubject', 'txtSoapSubject'],
        'o': ['PanelObject', 'txtSoapObject'],
        'p': ['PanelPlan', 'txtSoapPlan'],
    }
    panel = str(panel).lower()
    if panel not in parameters.keys():
        auto.Logger.WriteLine("Wrong panel in input_text",
                              auto.ConsoleColor.Red)
        return False

    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    window_soap = window_search(window_soap)
    if window_soap is None:
        auto.Logger.WriteLine("No window frmSoap", auto.ConsoleColor.Red)
        return False

    edit_contorl = window_soap.PaneControl(searchDepth=1, AutomationId=parameters[panel][0]).EditControl(
        searchDepth=1, AutomationId=parameters[panel][1])
    if edit_contorl.Exists():
        text_original = edit_contorl.GetValuePattern().Value
    return text_original

def get_S():
    get_text('s')

def get_O():
    get_text('o')

def get_P():
    get_text('p')


# def df_diagnosis_cata(df, config_schedule): # TODO 向量化版本 Not finished
#     date = datetime.datetime.today().strftime("%Y%m%d")
#     COLUMN = 'diagnosis_cata'
#     df[COLUMN] = ''
#     while(1):
#         check = input(f"Confirm or enter the new date(Default: {date})? ")
#         if check.strip() == '':
#             df[COLUMN] = df[COLUMN] + f"{date} s/p "
#             break
#         else:
#             if len(check)==8:
#                 df[COLUMN] = df[COLUMN] + f"{check} s/p "
#                 break
#             else:
#                 auto.Logger.WriteLine("WRONG FORMAT INPUT", auto.ConsoleColor.Red)

#     selector = df[config_schedule['COL_LENSX']].str.strip().lower() == 'lensx'
#     df.loc[selector, COLUMN] = df.loc[selector, COLUMN] + 'LenSx-'
    

def diagnosis_cata(df_selected_dict, config_schedule, date): 
    diagnosis = f"{date} s/p "
    
    #處理lensx+術式
    if df_selected_dict[config_schedule['COL_LENSX']].strip().lower() == 'lensx':
        diagnosis = diagnosis + 'LenSx-'
    if df_selected_dict[config_schedule['COL_OP']].lower().find('ecce') > -1:
        diagnosis = diagnosis + 'ECCE-IOL' + ' '
    else:
        diagnosis = diagnosis + 'Phaco-IOL' + ' '
    
    #處理分邊
    if df_selected_dict[config_schedule['COL_SIDE']].strip().lower() == '':
        if df_selected_dict[config_schedule['COL_DIAGNOSIS']].lower().find('od') > -1:
            diagnosis = diagnosis + 'OD'
        elif df_selected_dict[config_schedule['COL_DIAGNOSIS']].lower().find('os') > -1:
            diagnosis = diagnosis + 'OS'
        elif df_selected_dict[config_schedule['COL_DIAGNOSIS']].lower().find('ou') > -1:
            diagnosis = diagnosis + 'OU'
        else:
            auto.Logger.WriteLine("NO SIDE INFORMATION FROM SCHEDULE", auto.ConsoleColor.Red)
            return False
    elif df_selected_dict[config_schedule['COL_SIDE']].strip().lower() == 'od':
        diagnosis = diagnosis + 'OD'
    elif df_selected_dict[config_schedule['COL_SIDE']].strip().lower() == 'os':
        diagnosis = diagnosis + 'OS'
    elif df_selected_dict[config_schedule['COL_SIDE']].strip().lower() == 'ou':
        diagnosis = diagnosis + 'OU'
    else:
        auto.Logger.WriteLine("NO SIDE INFORMATION FROM SCHEDULE", auto.ConsoleColor.Red)
        return False
    
    #處理IOL+Final度數
    diagnosis = diagnosis + f"({df_selected_dict[config_schedule['COL_IOL']]} F+{df_selected_dict[config_schedule['COL_FINAL']]}D"

    #處理target
    if config_schedule['COL_TARGET'].strip() == '':
        diagnosis = diagnosis + ')' # 收尾用
    else:
        target = str(df_selected_dict[config_schedule['COL_TARGET']])
        if target.strip() == '':
            diagnosis = diagnosis + ')' # 收尾用
        else:
            diagnosis = diagnosis + f" T:{target})"

    return diagnosis


def diagnosis_ivi(df_selected_dict, config_schedule, date):
    diagnosis = 's/p IVI'
    if df_selected_dict[config_schedule['COL_CHARGE']].lower().find('(') > -1 :
        diagnosis = diagnosis + df_selected_dict[config_schedule['COL_CHARGE']] + ' '
    else:   
        diagnosis = diagnosis + df_selected_dict[config_schedule['COL_DRUGTYPE']].upper()[0]
        transform = {
            'drug-free': 'c',
            'all-free': 'f'
        }
        if df_selected_dict[config_schedule['COL_CHARGE']].lower() in transform.keys():
            diagnosis = diagnosis+ f"({transform.get(df_selected_dict[config_schedule['COL_CHARGE']].lower())}) "
        else:
            diagnosis = diagnosis+ f"({df_selected_dict[config_schedule['COL_CHARGE']]}) "
    
    # side
    if df_selected_dict[config_schedule['COL_SIDE']].strip().lower() == '':
        if df_selected_dict[config_schedule['COL_DIAGNOSIS']].lower().find('od') > -1:
            diagnosis = diagnosis + 'OD'
        elif df_selected_dict[config_schedule['COL_DIAGNOSIS']].lower().find('os') > -1:
            diagnosis = diagnosis + 'OS'
        elif df_selected_dict[config_schedule['COL_DIAGNOSIS']].lower().find('ou') > -1:
            diagnosis = diagnosis + 'OU'
        else:
            auto.Logger.WriteLine("NO SIDE INFORMATION FROM SCHEDULE", auto.ConsoleColor.Red)
            return False
    elif df_selected_dict[config_schedule['COL_SIDE']].strip().lower() == 'od':
        diagnosis = diagnosis + 'OD'
    elif df_selected_dict[config_schedule['COL_SIDE']].strip().lower() == 'os':
        diagnosis = diagnosis + 'OS'
    elif df_selected_dict[config_schedule['COL_SIDE']].strip().lower() == 'ou':
        diagnosis = diagnosis + 'OU'
    else:
        auto.Logger.WriteLine("NO SIDE INFORMATION FROM SCHEDULE", auto.ConsoleColor.Red)
        return False

    diagnosis = diagnosis+f" {date}"

    return diagnosis


def save(backtolist = True):
    '''
    存檔跳出功能
    '''
    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    window_soap = window_search(window_soap)
    if window_soap is None:
        auto.Logger.WriteLine("No window frmSoap", auto.ConsoleColor.Red)
        return False
    if backtolist:
        window_soap.SendKeys('{Ctrl}s', waitTime=0.05)
        # TODO 要確認有跳出?
    else:
        pane = window_soap.PaneControl(searchDepth=1, AutomationId="panel_bottom")
        button = pane.ButtonControl(searchDepth=1, AutomationId="btnSoapTempSave")
        # button.GetInvokePattern().Invoke()
        click_retry(button)
        message = window_soap.WindowControl(searchDepth=1, SubName='提示訊息')
        message = window_search(message)
        message.GetWindowPattern().Close()


def procedure_button(mode='ivi'): # FIXME沒辦法使用scroll and click功能
    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    window_soap = window_search(window_soap)
    if window_soap is None:
        auto.Logger.WriteLine("No window frmSoap", auto.ConsoleColor.Red)
        return False
    
    # search for edit button
    target_grid = None
    target_cell = None
    grids = datagrid_list_pid(window_soap.ProcessId)
    for grid in grids:
        if target_cell is not None:
            break
        row = datagrid_search('Edit', grid)
        for cell, depth in auto.WalkControl(row[0], maxDepth=1):
            if "處置" in cell.Name:
                target_cell = cell
                target_grid = grid
                break
    
    # 點擊Edit button
    auto.Logger.WriteLine(f"MATCHED BUTTON: {target_cell}", auto.ConsoleColor.Yellow)
    if click_datagrid(target_grid, [target_cell]) is True: # 換頁後不知道為何資料列變成-1
        # 跳出選擇PCS的視窗  
        win = auto.WindowControl(searchDepth=2,  AutomationId="dlgICDPCS")
        win = window_search(win)
        # Click datagrid
        if mode == 'ivi':
            search_term = '3E0C3GC'
        elif mode == 'phaco':
            pass # TODO 如果有側別要處理?
        datagrid = win.TableControl(searchDepth=1,  AutomationId="dgvICDPCS")
        t = datagrid_search([search_term], datagrid)     
        click_datagrid(datagrid, t)
    else:
        auto.Logger.WriteLine(f"FAILED: Clicking button({target_cell})", auto.ConsoleColor.Red)
    

def procedure_button_old(mode='ivi'):
    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    window_soap = window_search(window_soap)
    if window_soap is None:
        auto.Logger.WriteLine("No window frmSoap", auto.ConsoleColor.Red)
        return False
    target_grid = None
    target_list = []
    grids = datagrid_list_pid(window_soap.ProcessId)
    for grid in grids:
        if target_grid is not None:
            break
        for row, depth in auto.WalkControl(grid, maxDepth=1):
            if row.Name == '上方資料列':
                for cell, depth in auto.WalkControl(row, maxDepth=1):
                    if cell.Name == "處置":
                        target_grid = grid
                        break
                if target_grid is None:
                    break
            else:
                for cell, depth in auto.WalkControl(row, maxDepth=1):
                    if "處置" in cell.Name and cell.GetValuePattern().Value == 'Edit':
                        target_list.append(cell)
                        target_grid = grid
                        break
    
    auto.Logger.WriteLine(f"MATCHED BUTTON: {len(target_list)}", auto.ConsoleColor.Yellow)
    click_datagrid(target_grid, target_list)                   

    # 跳出選擇PCS的視窗 # TODO 如果有側別要處理? 
    win = auto.WindowControl(searchDepth=2,  AutomationId="dlgICDPCS")
    win = window_search(win)

    datagrid = win.TableControl(searchDepth=1,  AutomationId="dgvICDPCS")
    if mode == 'ivi':
        t = datagrid_search('3E0C3GC', datagrid)
        click_datagrid(datagrid, t)
    elif mode == 'phaco':
        pass # TODO 要處理側別
    

def confirm(mode=0):
    '''
    # TODO 目前這功能是給沒插卡的IVI出單用
    mode=0:直接不印病歷貼單送出
    mode=1:檢視帳單
    mode=2:檢視帳單後送出
    '''
    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    window_soap = window_search(window_soap)
    if window_soap is None:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|No window frmSoap", auto.ConsoleColor.Red)
        return False
    pane = window_soap.PaneControl(searchDepth=1, AutomationId="panel_bottom")
    button = pane.ButtonControl(searchDepth=1, AutomationId="btnSoapConfirm")
    # button.GetInvokePattern().Invoke() # 需要改成click
    click_retry(button)

    # 處理ICD換左右邊診斷視窗
    window = auto.WindowControl(searchDepth=2, AutomationId="dlgICDReply")
    window = window_search(window,3)
    button = window.ButtonControl(searchDepth=1, AutomationId="btnCancel")
    click_retry(button)
    
    # 解決讀卡機timeout的錯誤訊息
    window = auto.WindowControl(searchDepth=2, Name="錯誤訊息")
    window = window_search(window)
    if window is None:
        return
    button = window.ButtonControl(searchDepth=1, Name="確定")
    click_retry(button)

    # 處理繳費視窗
    window = auto.WindowControl(searchDepth=2, AutomationId="dlgNhiPpay")
    window = window_search(window)
    if mode == 0:
        # 送出(不印病歷)
        pane = window.PaneControl(searchDepth=1, AutomationId="btnBillViewOK")
        button = pane.ButtonControl(searchDepth=1, AutomationId="Button1")
        click_retry(button)
    elif mode == 1:
        # 檢視帳單
        button = window.ButtonControl(searchDepth=1, AutomationId="btnNhiPpayOK")
        click_retry(button)
    elif mode == 2:
        # 檢視帳單
        button = window.ButtonControl(searchDepth=1, AutomationId="btnNhiPpayOK")
        click_retry(button)
        # 檢視帳單後送出
        window = auto.WindowControl(searchDepth=2, AutomationId="frmBillView")
        window = window_search(window)
        pane = window.PaneControl(searchDepth=1, AutomationId="btnBillViewOK")
        button = pane.ButtonControl(searchDepth=1, AutomationId="Button1")
        click_retry(button)


# ==== Googlespreadsheet 資料擷取與轉換

# def df_from_gsheet(spreadsheet, worksheet, format_string=True, case_insensitive=True):
#     '''
#     取得gsheet資料並轉成dataframe, 預設是全部轉成文字形態(format_string=True)處理+小寫處理(case_insensitive=True)
#     Input: spreadsheet, worksheet, format_string=True, case_insensitive=True
#     '''
#     if CONFIG.get('SERVICE_JSON', None) != None:
#         client = pygsheets.authorize(service_account_json=CONFIG['SERVICE_JSON'])
#     else:
#         client = pygsheets.authorize(service_account_file=CONFIG['SERVICE_FILE'])

#     ssheet = client.open(spreadsheet)
#     wsheet = ssheet.worksheet_by_title(worksheet)
#     df = wsheet.get_as_df(has_header=True, include_tailing_empty=False, numerize=False) 
#     # has_header=True: 第一列當作header; numerize=False: 不要轉化數字; include_tailing_empty=False: 不讀取每一row後面空白Column資料
#     if format_string:
#         df = df.astype('string') #將所有dataframe資料改成用string格式處理，新的格式比object更精準
#     if case_insensitive:
#         df.columns = df.columns.str.lower() #將所有的columns name改成小寫 => case insensitive
#     return df


# def df_select_row(df: pandas.DataFrame):
#     '''
#     讓使用者選擇指定的row number，(-)表示連續範圍 (,)做分隔
#     '''
#     row_list_input = input("(-)表示連續範圍 (,)做分隔\n請輸入符合格式gsheet列碼: ")
#     row_set = set()
#     for i in row_list_input.split(','):
#         if len(i.split('-')) > 1:  # 如果有範圍標示，把這段範圍加入set
#             row_set.update(
#                 range(int(i.split('-')[0]), int(i.split('-')[1]) + 1))
#         else:
#             row_set.add(int(i))
#     # google spreadsheet的index是從1開始，所以df對應的相差一，此外還有一欄變成df的column headers，所以總共要減2
#     row_list = [row - 2 for row in row_set]
#     row_list.sort()

#     return df.iloc[row_list, :]  # 手動模式會回傳date_final=None
gc = gsheet.GsheetClient()


def gsheet_acc(id_list: list[str]):
    '''
    Input: list of short code of account. Ex:[4033,4123]
    Output: return dictionary of {'account short code':['account','password']} pairs
    '''
    return_dict= {}
    if type(id_list) is not list:
        id_list = [id_list]
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_ACC)
    for i in id_list:
        i=str(i).lower()
        selector = df['account'].str.contains(i, case = False)
        selected_df = df.loc[selector,['account','password']]
        if len(selected_df) == 0:
            auto.Logger.WriteLine(f"NOT EXIST ACCOUNT: {i}")
            continue
        # TODO 沒有查詢到帳密紀錄的帳號要從return dict內移除嗎?
        return_dict[i] = selected_df.iloc[0,:].to_list() #df變成series再輸出成list
    return return_dict


def gsheet_order_ovd(id_list: list[str]):
    '''
    Input: list of id. Ex:[4033,4123]
    Output: return dictionary of {'id':'ovd'} pairs
    '''
    return_dict= {}
    if type(id_list) is not list:
        id_list = [id_list]
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_OVD)
    
    default = df.loc[df['id']=='0','order'].values[0]
    
    for i in id_list:
        i=str(i).lower() # case insensitive and compare in string format
        selector = (df['id'].str.lower()==i) # case insensitive and compare in string format
        selected_df = df.loc[selector,'order']
        if len(selected_df) == 0:
            return_dict[i] = default #如果找不到資料使用預設的參數
        else:
            return_dict[i] = selected_df.values[0]
    return return_dict


def gsheet_iol_search_term(iol_list: list[str]):
    '''
    Input: list of iol that recorded on the surgery schedule
    Output: {'input iol':[formal iol, isNHI]...}, return the formal iol term, which is used in the OPD system and isNHI
    '''
    nhi_flag = '[NHI]'
    iol_search_dict = {}
    if type(iol_list) is not list:
        iol_list = [iol_list]
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_IOL, case_insensitive=False)

    for iol in iol_list:
        found = 0
        if iol in iol_search_dict.keys():
            continue
        for search_term in df.columns:
            if found == 1: # TODO 目前匹配到一個就停 => 但有些字串可能集合會重疊例如MX60和MX60T可能會錯誤，目前解法是將toric類型的放在前面讓他先去匹配長的
                break
            if search_term.strip() == '':
                continue
            for abbrev in df[search_term].to_list():
                abbrev = abbrev.strip().lower() # case insensitive
                if abbrev == '':
                    continue
                if abbrev in iol.lower(): # case insensitive
                    if nhi_flag in search_term: # 確認這個column是不是NHI IOL
                        isNHI = True
                        search_term = search_term.replace('[NHI]','')
                    else:
                        isNHI = False
                    iol_search_dict[iol] = [search_term, isNHI]
                    found = 1
                    break
        if found == 0:
            iol_search_dict[iol] = None, None
    
    return iol_search_dict


def gsheet_drug_to_druglist(df: pandas.DataFrame):
    '''
    transform the gsheet drugs data to specific form of drug list
    Example: drug_list = [{'name': 'Cravit oph sol', 'charge': '', 'dose': '', 'frequency': 'QID', 'route': '', 'duration': '7', 'eyedrop': 1}, {'name': 'Scanol tab', 'charge': '', 'dose': '', 'frequency': 'QIDPRN', 'route': '', 'duration': '1', 'eyedrop': 0}]
    '''
    drug_list = []

    for drug in df.columns:
        if drug =='id':
            continue

        # 處理是否為口服藥物
        if '[oral]' in drug:
            eye_tag = 0
            drug_name = drug.replace('[oral]','')
        else:
            eye_tag = 1
            drug_name = drug
        
        #轉換內部資料
        if df[drug].values[0] =='': # 空格表示沒有使用該藥物
            continue
        elif df[drug].values[0] =='0': # 使用門診系統的原始設定
            tmp = {
                'name': drug_name,
                'charge': '',
                'dose': '',
                'frequency': '',
                'route': '',
                'duration': '',
                'eyedrop': eye_tag
            }
            drug_list.append(tmp)
        else:
            frequency, duration = df[drug].values[0].split('*')
            tmp = {
                'name': drug_name,
                'charge': '',
                'dose': '',
                'frequency': frequency,
                'route': '',
                'duration': duration,
                'eyedrop': eye_tag
            }
            drug_list.append(tmp)
    
    return drug_list


def gsheet_drug(id_list: list[str]):
    '''
    Input: list of id. Ex:[4033,4123]
    Output: return form of drug list, which can be input to the select_precription
    '''
    return_dict= {}
    if type(id_list) is not list:
        id_list = [id_list]
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_DRUG)    
    
    default = df.loc[df['id']=='0',:]
    
    for i in id_list:
        i=str(i).lower() # case insensitive and compare in string format
        selector = (df['id'].str.lower()==i) # case insensitive and compare in string format
        selected_df = df.loc[selector,:]
        if len(selected_df) == 0:
            return_dict[i] = gsheet_drug_to_druglist(default) #如果找不到資料使用預設的參數
        else:
            return_dict[i] = gsheet_drug_to_druglist(selected_df)
    return return_dict


def gsheet_config_cata(id_list: list[str]):
    '''
    取得config_cata的資料, 回傳{id: {column_name:value}, {...}}
    '''
    return_dict= {}
    if type(id_list) is not list:
        id_list = [id_list]
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_SURGERY)
    
    for i in id_list:
        i=str(i)
        selector = (df['VS_CODE']==i) # TODO 考慮使用VS_CODE => 如果多筆VS_CODE怎辦?
        selected_df = df.loc[selector,:]
        if len(selected_df) == 0:
            auto.Logger.WriteLine(f"NOT EXIST: {i} in {inspect.currentframe().f_code.co_name}", auto.ConsoleColor.Red) # 找不到資料
        elif len(selected_df) > 1:
            auto.Logger.WriteLine(f"MORE than 2 same ID", auto.ConsoleColor.Red) # 找不到資料
            # TODO 未來處理可能加上讓使用者選擇功能，或是讓ID和VS_CODE不同
        else:
            return_dict[i] = selected_df.to_dict('records')[0]
            
            # 空值需要加入config_schedule? 
            # 需要跳過ID這一欄位?

    return return_dict


def gsheet_config_ivi(id_list: list[str], default='0'):
    '''
    取得config_ivi的資料, 回傳{id: {column_name:value}, {...}}
    '''
    return_dict= {}
    if type(id_list) is not list:
        id_list = [id_list]
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_IVI)

    for i in id_list:
        i=str(i)
        selector = (df['ID']==i) # 未來可以考慮還要使用ID還是乾脆VS_CODE?
        selected_df = df.loc[selector,:]
        if len(selected_df) == 0:
            auto.Logger.WriteLine(f"NOT EXIST: {i} in {inspect.currentframe().f_code.co_name}", auto.ConsoleColor.Red) # 找不到資料
            auto.Logger.WriteLine(f"USING DEFAULT for {i}", auto.ConsoleColor.Yellow)
            selected_df = df.loc[(df['ID']==default),:]
            return_dict[i] = selected_df.to_dict('records')[0]
        elif len(selected_df) > 1:
            auto.Logger.WriteLine(f"MORE than 2 same ID", auto.ConsoleColor.Red) # 找不到資料
            # TODO 未來處理可能加上讓使用者選擇功能，或是讓ID和VS_CODE不同
        else:
            return_dict[i] = selected_df.to_dict('records')[0]
            
            # 空值需要加入config_schedule? 
            # 需要跳過ID這一欄位?

    return return_dict


def gsheet_schedule_cata(config_schedule):
    '''
    (CATA)依照config_schedule資訊取得對應的刀表內容且輸出讓使用者確認
    '''
    auto.Logger.WriteLine(f"== ID:{config_schedule['ID']}|SPREADSHEET:{config_schedule['SPREADSHEET']}|WORKSHEET:{config_schedule['WORKSHEET']} ==", auto.ConsoleColor.Yellow)
    while(1):
        df = gc.get_df_select(gsheet.GSHEET_SPREADSHEET, config_schedule['WORKSHEET'], case_insensitive=False, format_string=False) # FIXME 真的要讓case_insensitive=False, format_string=False?
        # print dataframe將原始的時間去除
        print(df.reset_index()[[config_schedule['COL_DATE'], config_schedule['COL_HISNO'], config_schedule['COL_NAME'], config_schedule['COL_LENSX'], config_schedule['COL_IOL']]])
        check = input('Confirm the above-mentioned information(yes:Enter|no:n)? ')
        if check.strip() == '':
            return df


def gsheet_schedule_ivi(config_schedule):
    '''
    (IVI)依照config_schedule資訊取得對應的刀表內容且輸出讓使用者確認
    '''
    auto.Logger.WriteLine(f"== ID:{config_schedule['ID']}|SPREADSHEET:{config_schedule['SPREADSHEET']}|WORKSHEET:{config_schedule['WORKSHEET']} ==", auto.ConsoleColor.Yellow)
    while(1):
        df = gc.get_df_select(gsheet.GSHEET_SPREADSHEET, config_schedule['WORKSHEET'], case_insensitive=False, format_string=False) # FIXME 真的要讓case_insensitive=False, format_string=False?
        print(df.reset_index()[[config_schedule['COL_HISNO'], config_schedule['COL_NAME'], config_schedule['COL_DIAGNOSIS'], config_schedule['COL_DRUGTYPE'], config_schedule['COL_CHARGE']]])
        check = input('Confirm the above-mentioned information(yes:Enter|no:n)? ')
        if check.strip() == '':
            return df


def get_id_psw():    
    while(1):
        login_id = input("Enter your ID: ")
        if len(login_id) != 0:
            break
    while(1):
        login_psw = input("Enter your PASSWORD: ")
        if len(login_psw) != 0:
            break
    return login_id, login_psw

def get_date(mode:str='0'):
    '''
    取得時間.mode=0(西元紀年)|mode=1(民國紀年)
    '''
    mode = str(mode)
    if mode=='0':
        date = datetime.datetime.today().strftime("%Y%m%d") # 西元紀年
        while(1):
            check = input(f"Confirm or Enter the new date (NOW: {date})? ")
            if check.strip() == '':
                auto.Logger.WriteLine(f"DATE: {date}", auto.ConsoleColor.Yellow)
                return date
            else:
                if len(check)==8:
                    auto.Logger.WriteLine(f"DATE: {check}", auto.ConsoleColor.Yellow)
                    return check
                else:
                    auto.Logger.WriteLine("WRONG FORMAT INPUT", auto.ConsoleColor.Red)
    elif mode=='1':
        date = str(datetime.datetime.today().year-1911) + datetime.datetime.today().strftime("%m%d") # 民國紀年
        while(1):
            check = input(f"Confirm or Enter the new date (NOW: {date})? ")
            if check.strip() == '':
                auto.Logger.WriteLine(f"DATE: {date}", auto.ConsoleColor.Yellow)
                return date
            else:
                if len(check)==7:
                    auto.Logger.WriteLine(f"DATE: {check}", auto.ConsoleColor.Yellow)
                    return check
                else:
                    auto.Logger.WriteLine("WRONG FORMAT INPUT", auto.ConsoleColor.Red)
    else:
        return False

def get_date_today(mode:str='0'):
    '''
    取得今日時間.mode=0(西元紀年)|mode=1(民國紀年)
    '''
    mode = str(mode)
    if mode=='0':
        date = datetime.datetime.today().strftime("%Y%m%d") # 西元紀年
        auto.Logger.WriteLine(f"DATE: {date}", auto.ConsoleColor.Yellow)
        return date
    elif mode=='1':
        date = str(datetime.datetime.today().year-1911) + datetime.datetime.today().strftime("%m%d") # 民國紀年
        auto.Logger.WriteLine(f"DATE: {date}", auto.ConsoleColor.Yellow)
        return date




# HOTKEY
def hk_prescription_cata(stopEvent: Event):
    with auto.UIAutomationInitializerInThread():
        drug_list = gsheet_drug('0')['0']
        select_prescription(drug_list)

def hk_patientdata(stopEvent: Event):
    with auto.UIAutomationInitializerInThread():
        data = get_patient_data()
        print(data)


def main():
    # 選擇CATA|IVI mode
    mode = input("Choose the OPD program mode (1:CATA | 2:IVI | 0:hotkey): ")
    while(1):
        if mode not in ['1','2','0']:
            auto.Logger.WriteLine(f"WRONG MODE INPUT", auto.ConsoleColor.Red)
            mode = input("Choose the OPD program mode (1:CATA | 2:IVI | 0:hotkey): ")
        else:
            break
    
    ##########################################################
    if mode == '0': # 進入hotkey模式就會一直在回圈內，可能需要用另一個thread來跑?
        thread = threading.currentThread()
        auto.Logger.WriteLine(f"{thread.name}, {thread.ident}, MAIN", auto.ConsoleColor.Yellow)
        auto.RunByHotKey({
            (auto.ModifierKey.Control, auto.Keys.VK_1): hk_prescription_cata,
            (auto.ModifierKey.Control, auto.Keys.VK_2): hk_patientdata,
        }, waitHotKeyReleased=False)
    ##########################################################

    # 輸入要操作OPD系統的帳密
    acc_code = input("Please enter the short code of account (Ex:4123): ")
    dict_acc = gsheet_acc(acc_code)
    if len(dict_acc[acc_code])==0:
        # CONFIG上沒有此帳號密碼登記
        auto.Logger.WriteLine(f"USER({acc_code}) NOT EXIST IN CONFIG", auto.ConsoleColor.Red)
        login_id, login_psw = get_id_psw()
        acc_code = login_id[3:7]
    else:
        login_id, login_psw = dict_acc[acc_code]

    running, pid = process_exists(CONFIG['PROCESS_NAME'])

    if mode == '1': # CATA
        # 使用者輸入: 獲取刀表+日期模式
        config_schedule = gsheet_config_cata(acc_code)[acc_code]
        date = get_date_today(config_schedule['DATE_MODE'])
        df = gsheet_schedule_cata(config_schedule)

        # 開啟門診程式
        if running:
            auto.Logger.WriteLine("OPD program is running", auto.ConsoleColor.Yellow)
            login_change(login_id, login_psw, CONFIG['SECTION_CATA'], CONFIG['ROOM_CATA'])
        else:
            login(CONFIG['OPD_PATH'], login_id, login_psw, CONFIG['SECTION_CATA'], CONFIG['ROOM_CATA'])
        
        # 將所有病歷號加入非常態掛號
        hisno_list = df[config_schedule['COL_HISNO']].to_list()
        appointment(hisno_list, skip_checkpoint=True)
        
        # 取得已有暫存list
        exclude_patient_list = []
        window_main = auto.WindowControl(searchDepth=1, SubName='台北榮民總醫院', AutomationId="frmPatList")
        window_main = window_search(window_main)
        datagrid_patient = window_main.TableControl(searchDepth=1, SubName='DataGridView', AutomationId="dgvPatsList")
        patient_list_values = datagrid_values(datagrid=datagrid_patient)
        for row in patient_list_values:
            if row[3] in hisno_list and row[9]=='是': # row[3]表示病歷號；row[9]表示暫存欄位 => 未來應改成column_name搜尋方式
                exclude_patient_list.append(row[3])
        
        # 逐一病人處理
        df.set_index(keys=config_schedule['COL_HISNO'], inplace=True)
        for hisno in hisno_list:
            # 跳過已有暫存者
            if hisno in exclude_patient_list:
                auto.Logger.WriteLine(f"Already saved: {hisno}", auto.ConsoleColor.Yellow)
                continue
            
            # ditto
            res = ditto(hisno, skip_checkpoint=True)
            if res == False:
                continue

            # 選擇phaco模式
            if df.loc[hisno, config_schedule['COL_LENSX']].strip() == '': # 沒有選擇lensx
                select_phaco_mode(0)
            elif df.loc[hisno, config_schedule['COL_LENSX']].strip().lower() == 'lensx':
                select_phaco_mode(1)
            else:
                auto.Logger.WriteLine(f"Lensx資訊辨識問題({df.loc[hisno,config_schedule['COL_LENSX']].strip()}) 先以NHI組套處理")
                select_phaco_mode(0)
            
            side = df.loc[hisno, config_schedule['COL_SIDE']].strip() # TODO 以後side要以side欄位資訊為主還是診斷內的資訊為主?
            # 取得刀表iol資訊
            iol = df.loc[hisno, config_schedule['COL_IOL']].strip()
            # 取得該燈號常用OVD
            ovd = gsheet_order_ovd(acc_code)[acc_code]
            select_iol_ovd(iol=iol, ovd=ovd)
            # 修改order的side
            select_order_side_all(side)
            
            # 處理藥物
            drug_list = gsheet_drug(acc_code)[acc_code]
            drug_list = select_eyedrop_side(drug_list, side=side)
            select_prescription(drug_list)

            # 在Subject框內輸入手術資訊 => 要先組合手術資訊
            diagnosis = diagnosis_cata(df.loc[[hisno], :].to_dict('records')[0], config_schedule, date) # 目前使用將資料以dict方式傳入
            set_S(diagnosis)

            # 暫存退出
            save()
    
    # FIXME 有些函數尚未更新
    elif mode == '2': # IVI 
        # 使用者輸入: 獲取刀表+日期模式
        config_schedule = gsheet_config_ivi(acc_code)[acc_code]
        date = get_date_today(config_schedule['DATE_MODE'])
        df = gsheet_schedule_ivi(config_schedule)

        # 開啟門診程式
        if running:
            auto.Logger.WriteLine("OPD program is running", auto.ConsoleColor.Yellow)
            login_change(login_id, login_psw, CONFIG['SECTION_PROCEDURE'], CONFIG['ROOM_PROCEDURE']) 
        else:
            login(CONFIG['OPD_PATH'],login_id, login_psw, CONFIG['SECTION_PROCEDURE'], CONFIG['ROOM_PROCEDURE'])

        # 將所有病歷號加入非常態掛號
        hisno_list = df[config_schedule['COL_HISNO']].to_list()
        appointment(hisno_list)

        # 逐一病人處理
        df.set_index(keys=config_schedule['COL_HISNO'], inplace=True)
        for hisno in hisno_list:
            # ditto
            ditto(hisno)
            
            side = df.loc[hisno, config_schedule['COL_SIDE']].strip()
            charge = df.loc[hisno, config_schedule['COL_CHARGE']].strip()
            drug_ivi = df.loc[hisno, config_schedule['COL_DRUGTYPE']].strip()
            # TODO
            # TODO 要依照charge處理order
            # TODO 要依照charge決定drug_ivi要不要開上去
            # TODO 依照charge決定出單方式? => 兩次出單

            
            # 處理其它藥物
            other_drug_list = gsheet_drug('ivi')['ivi']
            other_drug_list = select_eyedrop_side(other_drug_list, side=side)
            select_prescription(other_drug_list)

            # 在Subject框內輸入手術資訊 => 要先組合手術資訊
            diagnosis = diagnosis_ivi(df.loc[[hisno], :].to_dict('records')[0], config_schedule, date)
            set_S(diagnosis)

            # 暫存退出
            save()

def load_config():
    # load from local json if path existed
    if os.path.exists(CONFIG_JSON_PATH):
        auto.Logger.WriteLine(f"CONFIG_JSON_PATH EXISTS AND APPLY CONFIG FROM FILE", auto.ConsoleColor.Yellow)
        with open(CONFIG_JSON_PATH, 'r', encoding='utf-8') as f:
            tmp = json.load(fp=f)
            for i in tmp: # update CONFIG
                CONFIG[i] = tmp[i]
    
    # Check if OPD_PATH exist
    if os.path.exists(CONFIG['OPD_PATH']) == False:
        auto.Logger.WriteLine(f"OPD_PATH NOT EXISTS: {CONFIG['OPD_PATH']}", auto.ConsoleColor.Red)
        CONFIG['OPD_PATH'] = input("Please enter path of OPD system: ")

    # Check if SERVICE_FILE exist
    if (CONFIG.get('SERVICE_JSON', None) == None) and (os.path.exists(CONFIG['SERVICE_FILE']) == False):
        auto.Logger.WriteLine(f"SERVICE_FILE NOT EXISTS: {CONFIG['SERVICE_FILE']}", auto.ConsoleColor.Red)
        CONFIG['SERVICE_FILE'] = input("Please enter path of SERVICE_FILE: ")

    def clean(datalist:list): # 清掉從google sheet上取得資料會有空白row的狀況
        final_list = []
        for i in datalist:
            if i.strip() == '':
                continue
            else:
                final_list.append(i)
        if len(final_list) == 1:
            return final_list[0]
        return final_list

    # load from web(gsheet)
    df = df_from_gsheet(CONFIG['SPREADSHEET_CONFIG'], 'config_others', case_insensitive=False)
    CONFIG['PROCESS_NAME'] = clean(df['PROCESS_NAME'].to_list())
    CONFIG['SECTION_CATA'] = clean(df['SECTION_CATA'].to_list())
    CONFIG['ROOM_CATA'] = clean(df['ROOM_CATA'].to_list())
    CONFIG['SECTION_PROCEDURE'] = clean(df['SECTION_PROCEDURE'].to_list())
    CONFIG['ROOM_PROCEDURE'] = clean(df['ROOM_PROCEDURE'].to_list())
    CONFIG['SECTION_OPH'] = clean(df['SECTION_OPH'].to_list())
    return CONFIG


'''
TEST_MODE = False
GSHEET_SPREADSHEET = 'config_vghbot'
GSHEET_WORKSHEET_SURGERY = 'set_surgery'
GSHEET_WORKSHEET_IVI = 'set_ivi'
GSHEET_WORKSHEET_TEMPLATE_OPNOTE = 'template_opnote'
'''


TEST_MODE = False
CONFIG_JSON_PATH = "config_bot_opd.json"
SERVICE_JSON = None
CONFIG = {
    "OPD_PATH": "C:\\Users\\Public\\Desktop\\門診系統.appref-ms",
    #如果有json字串會優先於file，權限問題通常用file
    "SERVICE_JSON": SERVICE_JSON, 
    "SERVICE_FILE": "vghbot-5fe0aba1d3b9.json",
    "SPREADSHEET_CONFIG": "config_vgh_automation"
}
auto.uiautomation.SetGlobalSearchTimeout(10)  # 應該使用較長的timeout來防止電腦反應太慢，預設就是10秒
auto.uiautomation.DEBUG_SEARCH_TIME = TEST_MODE 
CONFIG = load_config()

if __name__ == '__main__':
    if auto.IsUserAnAdmin():
        main()
    else:
        print('RunScriptAsAdmin', sys.executable, sys.argv)
        auto.RunScriptAsAdmin(sys.argv)


# OLD MAIN
# running, pid = process_exists(PROCESS_NAME)
#     if running:
#         auto.Logger.WriteLine("OPD program is running", auto.ConsoleColor.Yellow)
#         with open(JSON_PATH, 'r', encoding='utf-8') as f:
#             config = json.load(fp=f)
#             login_change(config['OPD_ACCOUNT'], config['OPD_PASSWORD'], config['MODE']['OP_CATA']['OPD_SECTION'], config['MODE']['OP_CATA']['OPD_ROOM'])
#     else:
#         with open(JSON_PATH, 'r', encoding='utf-8') as f:
#             config = json.load(fp=f)
#             login(config['OPD_PATH'], config['OPD_ACCOUNT'], config['OPD_PASSWORD'], config['MODE']['OP_CATA']['OPD_SECTION'], config['MODE']['OP_CATA']['OPD_ROOM'])

# HOT KEY

# import sys
# import uiautomation as auto


# def demo1(stopEvent: Event):
#     thread = threading.currentThread()
#     print(thread.name, thread.ident, "demo1")
#     auto.SendKeys('12312313')


# def demo2(stopEvent: Event):
#     thread = threading.currentThread()
#     print(thread.name, thread.ident, "demo2")


# def demo3(stopEvent: Event):
#     thread = threading.currentThread()
#     print(thread.name, thread.ident, "demo3")


# def main():
#     thread = threading.currentThread()
#     print(thread.name, thread.ident, "main")
#     auto.RunByHotKey({
#         (0, auto.Keys.VK_F2): demo1,
#         (auto.ModifierKey.Control, auto.Keys.VK_1): demo2,
#         (auto.ModifierKey.Control | auto.ModifierKey.Shift, auto.Keys.VK_2): demo3,
#     }, waitHotKeyReleased=False)

# FOR NOTEPAD
# window = auto.WindowControl(searchDepth=1, SubName='Untitled')
# c_edit = window.EditControl(AutomationId = "15")
# c_menubar = window.MenuBarControl(AutomationId = "MenuBar", Name="Application", searchDepth=1 )
# c_menuitem = c_menubar.MenuItemControl(Name='Format', searchDepth=1)
# res = c_menuitem.GetExpandCollapsePattern().Expand(waitTime=10)
# c_format = window.MenuControl(SubName="Format")
# c_font = c_format.MenuItemControl(SubName="Font", searchDepth=1)
# c_font.GetInvokePattern().Invoke()

