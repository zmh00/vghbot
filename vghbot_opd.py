import uiautomation as auto
import os
import time
# import threading # for hotkey
# from threading import Event # for hotkey
# import multiprocessing
import inspect
from ctypes import windll
import sys
import subprocess

import pandas
import datetime

from pathlib import Path
import gsheet
import updater_cmd


# pyinstaller
# pyinstaller --paths ".\.venv-opd\Lib\site-packages\uiautomation\bin" -F vghbot_opd.py 

# ==== 基本操作架構

def process_exists(process_name):
    '''
    Check if a program (based on its name) is running
    Return yes/no exists window and its PID
    '''
    call = 'TASKLIST', '/FI', 'imagename eq %s' % process_name
    # use buildin check_output right away
    try:
        output = subprocess.check_output(call, universal_newlines=True) # 在中文的console中使用需要解析編碼為big5???
        output = output.strip().split('\n')
        if len(output) == 1:  # 代表只有錯誤訊息
            return False, 0
        else:
            # check in last line for process name
            last_line_list = output[-1].lower().split()
        return last_line_list[0].startswith(process_name.lower()), int(last_line_list[1])
    except subprocess.CalledProcessError:
        return False, 0


def process_responding(name):
    """Check if a program (based on its name) is responding"""
    cmd = 'tasklist /FI "IMAGENAME eq %s" /FI "STATUS eq running"' % name
    status = subprocess.Popen(cmd, stdout=subprocess.PIPE).stdout.read()
    status = str(status).lower() # case insensitive
    return name in status


def process_responding_PID(pid):
    """Check if a program (based on its PID) is responding"""
    cmd = 'tasklist /FI "PID eq %d" /FI "STATUS eq running"' % pid
    status = subprocess.Popen(cmd, stdout=subprocess.PIPE).stdout.read()
    status = str(status).lower()
    return str(pid) in status


def captureimage(control = None, postfix = ''):
    pass
    # auto.Logger.WriteLine('CAPTUREIMAGE INITIATION', auto.ConsoleColor.Yellow)
    # if control is None:
    #     c = auto.GetRootControl()
    # else:
    #     c = control
    # if postfix == '':
    #     path = f"{datetime.datetime.today().strftime('%Y%m%d_%H%M%S')}.png"
    # else:
    #     path = f"{datetime.datetime.today().strftime('%Y%m%d_%H%M%S')}_{postfix}.png"
    # c.CaptureToImage(path)


def window_dfs(processId, search_from = None, depth=0, maxDepth=2, only_one = False):
    '''
    監控指定PID程序下新視窗的產生，DFS方式搜尋，只回傳enabled的視窗
    如果only_one == True，會直接回傳一個control;如果only_one == False，會回傳一個control list
    '''
    if type(processId) == str:
        processId = int(processId)
    
    target_list = []
    if depth > maxDepth:
        if only_one is True:
            return None
        else:
            return target_list
    
    if search_from is None:
        search_from = auto.GetRootControl()
    
    t_start = time.perf_counter()
    for control, _depth in auto.WalkControl(search_from, maxDepth=1):
        if (control.ProcessId == processId) and (control.ControlType==auto.ControlType.WindowControl):
            if control.IsEnabled ==  True: # 利用別的判斷方式? control.GetWindowPattern().WindowInteractionState
                if only_one is True:
                    if TEST_MODE:
                        t_end = time.perf_counter()
                        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|監控第{depth}層以下共花費:{t_end-t_start}")
                    return control
                else:
                    target_list.append(control)
            else:
                depth = depth + 1
                l = window_dfs(processId=processId, search_from=control, depth=depth, maxDepth=maxDepth, only_one=only_one)
                if only_one is True and l is None:
                    continue
                elif only_one is True and l is not None:
                    if TEST_MODE:
                        t_end = time.perf_counter()
                        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|監控第{depth}層以下共花費:{t_end-t_start}")
                    return l
                else:
                    target_list.extend(l)
    if TEST_MODE:
        t_end = time.perf_counter()
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|監控第{depth}層以下共花費:{t_end-t_start}")
    if only_one is True:
        return None
    else:
        return target_list

def window_check_exist_enabled(control: auto.WindowControl):
    auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|== TopWindow (Name:{control.Name}|AutomationId:{control.AutomationId}) ==")
    if control.Exists():
        if control.IsEnabled == True:
            return True
        else:
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT ENABLED: TopWindow (Name:{control.Name}|AutomationId:{control.AutomationId})")
    else:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT EXIST: TopWindow (Name:{control.Name}|AutomationId:{control.AutomationId})")
    return False


def window_policy(control: auto.WindowControl):
    '''
    設定每個視窗處理方式的函數，將條件較唯一的放在較上面(control.AutomationId or control.Name)
    '''
    try:
        if control.AutomationId == "frmDCRSignOn": # 登入介面
            if window_check_exist_enabled(control):
                acc = control.EditControl(AutomationId="txtSignOnID", Depth=1)
                acc.GetValuePattern().SetValue(CONFIG['ACCOUNT'])
                psw = control.EditControl(AutomationId="txtSignOnPassword", Depth=1)
                psw.GetValuePattern().SetValue(CONFIG['PASSWORD'])
                section = control.EditControl(AutomationId="1001", Depth=2)
                section.GetValuePattern().SetValue(CONFIG['SECTION_ID'])
                room = control.EditControl(AutomationId="txtRoom", Depth=1)
                room.GetValuePattern().SetValue(CONFIG['ROOM_ID'])
                signin = control.ButtonControl(AutomationId="btnSignon", Depth=1)
                click_retry(signin)
            else:
                return False
        elif control.AutomationId == "dlgMessageCenter": # 登入後，醫師待辦事項通知
            if window_check_exist_enabled(control):
                control.GetWindowPattern().Close()
            else:
                return False
        elif control.AutomationId == "##########": # 複製用樣板
            if window_check_exist_enabled(control):
                control.GetWindowPattern().Close()
            else:
                return False
        elif control.AutomationId == "dlgNewTOCC": # DITTO，詢問TOCC
            if window_check_exist_enabled(control):
                control.CheckBoxControl(Depth=2, AutomationId="ckbAllNo").GetTogglePattern().Toggle()
                control.ButtonControl(Depth=2, AutomationId="btnOK").GetInvokePattern().Invoke()
            else:
                return False
        elif control.AutomationId == "dlgSMOBET": # DITTO，健康行為登錄
            if window_check_exist_enabled(control):
                control.GetWindowPattern().Close()
            else:
                return False
        elif control.AutomationId == "dlgWarMessage": # DITTO，警告提示訊息; 登入後，醫事卡非登入醫師本人通知
            if window_check_exist_enabled(control):
                control.GetWindowPattern().Close() # TODO 適用兩個狀況嗎?
                # c_button_ok = control.ButtonControl(searchDepth=1, AutomationId="OK_Button", SubName="繼續")
                # c_button_ok.GetInvokePattern().Invoke()
            else:
                return False
        elif control.AutomationId == "dlgDrugAllergyDetailAndEdit": # DITTO後過敏提示視窗
            if window_check_exist_enabled(control):
                control.ButtonControl(Depth=3, SubName='無需更新', AutomationId="Button1").GetInvokePattern().Invoke()
            else:
                return False
        elif control.AutomationId == "FlaxibleMessage": # 改變OPD時，跳出警告訊息;
            if window_check_exist_enabled(control):
                control.ButtonControl(Depth=2, AutomationId="btnOK").GetInvokePattern().Invoke()
            else:
                return False
        elif control.Name == "訊息":
            if window_check_exist_enabled(control):
                if control.TextControl(searchDepth=1, SubName="卡機重新連線").Exists(): # 登入前卡機重新連線警告
                    button = control.ButtonControl(searchDepth=1, SubName="略過")
                    button.GetInvokePattern().Invoke()
                else:
                    text = control.TextControl(searchDepth=1)
                    if text.Exists():
                        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|未知訊息視窗:{text.Name}", auto.ConsoleColor.Yellow)
                    else:
                        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|未知訊息視窗", auto.ConsoleColor.Yellow)
                        captureimage(postfix=inspect.currentframe().f_code.co_name)
                    control.GetWindowPattern().Close()
            else:
                return False
        else: # 未登錄視窗
            auto.Logger.WriteLine(f"TopWindow (Name:{control.Name}|AutomationId:{control.AutomationId}) => No available policy")
            captureimage(postfix=inspect.currentframe().f_code.co_name)
            control.GetWindowPattern().Close()
        return True
    except Exception as e:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Something wrong happened:{e}", auto.ConsoleColor.Red)


def window_pending(processId, pending_control, retry = 5, excluded_control=None):
    '''
    等待一個指定的視窗，嘗試retry次數內，迴圈去抓目前最底部的視窗，每個不是指定的視窗都交給視窗處理原則，直到指定視窗出現就返回
    '''
    while retry>=0:
        time.sleep(0.2)
        try:
            top_window = window_dfs(processId=processId, only_one=True)
            if top_window is None:
                auto.Logger.WriteLine(f"No qualified TopWindow")
            if (top_window.AutomationId == pending_control.searchProperties.get('AutomationId')) or (top_window.Name == pending_control.searchProperties.get('Name')) or (pending_control.searchProperties.get('SubName', '!@#$') in top_window.Name): 
                if window_check_exist_enabled(top_window):
                    auto.Logger.WriteLine(f"PENDING EXIST: (Name:{pending_control.Name}|AutomationId:{pending_control.AutomationId})", auto.ConsoleColor.Yellow)
                    return True
                else:
                    continue
            else:
                if excluded_control is not None:
                    if top_window.AutomationId == excluded_control.searchProperties.get('AutomationId'): # 將原本起始的視窗排除避免無窮迴圈
                        continue
                res = window_policy(top_window)
                if res == False:
                    continue
            retry = retry - 1
        except Exception as e:
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Error Message: {e}", auto.ConsoleColor.Red)

    wait_for_manual_control(f"Window Name:{pending_control.searchProperties.get('Name')}|Window AutomationId:{pending_control.searchProperties.get('AutomationId')}")


def wait_for_manual_control(info):
    auto.Logger.WriteLine(f"[視窗異常]請自行操作視窗到指定視窗:\n{info}", auto.ConsoleColor.Cyan)
    while(1):
        choice = input('確認已操作至指定視窗?(y/n): ')
        if choice.lower().strip() == 'y':
            return True
        else:
            continue


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
                if process_responding(CONFIG['PROCESS_NAME'][0]):
                    auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Window not found: {window.GetSearchPropertiesStr()}", auto.ConsoleColor.Red)
                    retry = retry-1
                else:
                    auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Process not responding", auto.ConsoleColor.Red)
                time.sleep(0.2)
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Window not found(after {_retry} times): {window.GetSearchPropertiesStr()}", auto.ConsoleColor.Red)
        captureimage(postfix=inspect.currentframe().f_code.co_name)
        return None
    except Exception as e:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Something wrong unexpected: {window.GetSearchPropertiesStr()}", auto.ConsoleColor.Red)
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Error Message: {e}", auto.ConsoleColor.Red)
        captureimage(postfix=inspect.currentframe().f_code.co_name)
        retry = retry - 1 
        return window_search(window, retry=retry) # 目前使用遞迴處理 => 會無窮迴圈後續要考慮新方式


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
        value = item.GetLegacyIAccessiblePattern().Value
        if TEST_MODE:
            print(f"Datagrid({datagrid.AutomationId}):{value}")
        if ';' in value:  # 有資料的列的判斷方式
            if column_index is not None:
                t = value.replace('(null)', '').split(';')[column_index]
                t = t.strip() # 把一些空格字元清掉
                value_list.append(t)
            else:
                t = value.replace('(null)', '').split(';') # 把(null)轉成''
                value_list.append([cell.strip() for cell in t])  # 把每個cell內空格字元清掉
    return value_list


def datagrid_search(search_text: list, datagrid, column_name=None, retry=5, only_one=True, skip = 0):
    '''
    Search datagrid based on search_text, each search_text can only be matched once(case insensitive), return the list of all the matched item
    search_text, 可以一次傳入要在此datagrid搜尋的資料陣列
    column_name=None, 指定column做搜尋
    retry=5, 預設重覆搜尋5次
    only_one=True, 找到符合一個target item就回傳 => 增加效率，但同時要搜尋多筆應設定only_one=False
    skip = N，表示跳過匹配N次
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
    
    checked_list = [False] * len(search_text)
    for row in children:
        if (only_one == True) and len(target_list)>0: # 有找到一個就跳出
            break
        if ('資料列' in row.Name) and (row.Name != '上方資料列'): # 有資料的列才做判斷
            value = row.GetLegacyIAccessiblePattern().Value
            match = value.lower() # case insensitive
            if column_index is not None: # 有找到Column
                match = match.replace('(null)', '').split(';') # 將(null)轉成空字串''並且透過;分隔欄位資訊
                if column_index >= len(match):
                    continue
                else:
                    match = match[column_index]
            if TEST_MODE:
                auto.Logger.WriteLine(f"Datagrid:{datagrid.AutomationId})|Search target:{search_text}|Row value:{value}|Match value:{match}")
            for i, text in enumerate(search_text): # 用每一個serch_text去配對該row的資訊
                if checked_list[i]: # 搜尋過字串不再使用，但若有同樣的search_text會各別找一次
                    continue
                if text.lower() in match:
                    if skip == 0:
                        target_list.append(row)
                        checked_list[i] = True  # 讓搜尋字串只用一次
                        if TEST_MODE:
                            auto.Logger.WriteLine(f"Datagrid found:{text}", auto.ConsoleColor.Yellow)
                        break
                    else:
                        if TEST_MODE:
                            auto.Logger.WriteLine(f"Skip:{skip}|Datagrid found:{text}")
                        skip = skip - 1
                        break
    return target_list
    

def click_blockinput(control, doubleclick=False, simulateMove=False, waitTime = 0.2):
    try:
        res = windll.user32.BlockInput(True)
        if TEST_MODE:
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Control:{control.Name}|{control.GetClickablePoint()}", auto.ConsoleColor.Yellow)
        if doubleclick:
            control.DoubleClick(waitTime=waitTime, simulateMove=simulateMove)
        else:
            control.Click(waitTime=waitTime, simulateMove=simulateMove)
    except Exception as e:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Blockinput&Click Failed: {e}", auto.ConsoleColor.Red) # TODO 如果因為物件不存在而沒點擊到，不會跳出exception，但會有error message=>抓response??
        res = windll.user32.BlockInput(False)
        return False
    res = windll.user32.BlockInput(False)
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
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT EXIST: scroll button")
        for t in target_list:
            t.SetFocus()
            if click_blockinput(t, doubleclick=doubleclick):
                time.sleep(0.5)
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
 

# ==== 門診系統操作函數
# 每個操作可以分幾個階段: start_point, main_body, end_point
def login(account: str, password: str, section_id: str, room_id: str):
    CONFIG['ACCOUNT'] = account
    CONFIG['PASSWORD'] = password
    CONFIG['SECTION_ID'] = section_id
    CONFIG['ROOM_ID'] = room_id
    running, CONFIG['PROCESS_ID'] = process_exists(CONFIG['PROCESS_NAME'][0])
    if running is False:
        os.startfile(CONFIG['OPD_PATH'])
        auto.Logger.WriteLine("OPD system started", auto.ConsoleColor.Yellow)
        while (running is False):
            running, CONFIG['PROCESS_ID'] = process_exists(CONFIG['PROCESS_NAME'][0])
            time.sleep(1)
    else:
        auto.Logger.WriteLine("OPD program is running", auto.ConsoleColor.Yellow)
    
    while True:
        window_top = window_dfs(CONFIG['PROCESS_ID'], only_one=True)
        if window_top is not None:
            break
    if window_top.AutomationId == "frmPatList":
        login_change_opd(account, password, section_id, room_id)
    else:
        # 等待進入病人清單主視窗
        window_main = auto.WindowControl(AutomationId="frmPatList", searchDepth=1)
        window_pending(CONFIG['PROCESS_ID'], pending_control=window_main, retry=10)
        

def login_all(account: str, password: str, section_id: str, room_id: str): # TODO 判斷登入是否成功的部份還沒移植出來，其他已更新
    # 等待載入
    window_login = auto.WindowControl(AutomationId="frmDCRSignOn", searchDepth=1)
    window_pending(CONFIG['PROCESS_ID'], pending_control=window_login, retry=20)

    # TODO 測試完可淘汰
    # login_window = window_search(login_window, -1)
    # if login_window is not None:
    #     auto.Logger.WriteLine("Finished: Loading OPD system", auto.ConsoleColor.Yellow)
    # else:
    #     auto.Logger.WriteLine("Failed: Loading OPD system", auto.ConsoleColor.Red)
    #     return False

    # # 跳出略過按鈕
    # msg = login_window.ButtonControl(SubName="略過", Depth=2)
    # msg.GetInvokePattern().Invoke()
    # TODO 測試完可淘汰

    # 填入開診資料
    acc = window_login.EditControl(AutomationId="txtSignOnID", Depth=1)
    acc.GetValuePattern().SetValue(account)
    psw = window_login.EditControl(AutomationId="txtSignOnPassword", Depth=1)
    psw.GetValuePattern().SetValue(password)
    section = window_login.EditControl(AutomationId="1001", Depth=2)
    section.GetValuePattern().SetValue(section_id)
    room = window_login.EditControl(AutomationId="txtRoom", Depth=1)
    room.GetValuePattern().SetValue(room_id)
    signin = window_login.ButtonControl(AutomationId="btnSignon", Depth=1)
    click_retry(signin)  # 為何改用click是因為只要使用invoke若遇到popping window會報錯且卡死API

    # 判斷登入是否成功
    check_login = window_login.WindowControl(SubName="錯誤訊息", searchDepth=1)
    if check_login.Exists(maxSearchSeconds=1, searchIntervalSeconds=0.2):
        auto.Logger.WriteLine(f"Login failed!", auto.ConsoleColor.Red)
        return False

    window_main = auto.WindowControl(AutomationId="frmPatList", searchDepth=1)
    window_pending(CONFIG['PROCESS_ID'], pending_control=window_main, retry=8)

    # TODO 測試後可淘汰
    # # 處理登入非該診醫師的警告
    # warning_msg = auto.WindowControl(searchDepth=2, AutomationId="FlaxibleMessage")
    # warning_msg = window_search(warning_msg,1)
    # if warning_msg is not None:
    #     button = warning_msg.ButtonControl(Depth=2, AutomationId="btnOK")
    #     if button.Exists():
    #         # button.GetInvokePattern().Invoke() # TODO 需要測試會不會卡住API
    #         click_retry(button, warning_msg)
    #     else:
    #         auto.Logger.WriteLine("No OK_Button under FlaxibleMessage", auto.ConsoleColor.Red)
    # else:
    #     auto.Logger.WriteLine("No FlaxibleMessage", auto.ConsoleColor.Red)

    # # 醫師待辦事項通知 => 可window.close()
    # warning_msg = auto.WindowControl(AutomationId="dlgMessageCenter", searchDepth=2)
    # warning_msg = window_search(warning_msg,1)
    # if warning_msg is not None:
    #     warning_msg.GetWindowPattern().Close()
    # else:
    #     auto.Logger.WriteLine("No dlgMessageCenter", auto.ConsoleColor.Red)

    # # 醫事卡非登入醫師本人通知
    # warning_msg = auto.WindowControl(AutomationId="dlgWarMessage", searchDepth=2)
    # warning_msg = window_search(warning_msg,1)
    # if warning_msg is not None:
    #     warning_msg.GetWindowPattern().Close()
    # else:
    #     auto.Logger.WriteLine("No dlgWarMessage", auto.ConsoleColor.Red)
    # TODO 測試後可淘汰

    # 以下目標一次處理所有dialog(COVID警告通知+此診目前無掛號)，因為dialog不是獨立的，在walkatree內會有階層關係，但不是每一個都有enabled可以被操作，要按照順序處理
    # 此診目前無掛號 會是第一個跳出的dialog，但其他訊息後續出現後，此dialog會被設成not enabled => 導致不能被選中關閉
    # time.sleep(0.5)
    # auto.SendKeys("{SPACE}" * 3)
    # auto.SendKeys("{SPACE}" * 3)

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


def login_change_opd(account: str, password: str, section_id: str, room_id: str):
    window_main = auto.WindowControl(searchDepth=1, AutomationId="frmPatList")
    window_main = window_search(window_main)
    if window_main is None:
        auto.Logger.WriteLine("No window frmPatList", auto.ConsoleColor.Red)
        return False
    
    section_room = window_main.TextControl(searchDepth=1, AutomationId="Label2").Name
    doctor_name = window_main.TextControl(searchDepth=1, AutomationId="lblPatsDocname").Name[3:6]
    if TEST_MODE:
        auto.Logger.WriteLine(f"Section: {section_room}|Doctor Name:{doctor_name}", auto.ConsoleColor.Yellow)

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
    else:
        auto.Logger.WriteLine("No Window Relog", auto.ConsoleColor.Red)

    # 換到班表沒有匹配的診會出現警告彈窗 => 造成API卡住
    # TODO 在沒有實際存在此視窗下 就算前面用click 也出現API卡住，但等err過去可以繼續使用
    window_main = auto.WindowControl(searchDepth=1, AutomationId="frmPatList")
    window_pending(CONFIG['PROCESS_ID'], window_main)

    # # 換到沒有掛號的診會跳出無掛號的訊息視窗 => 可以用空白鍵解決 # TODO還需要這個嗎?
    # auto.SendKeys("{SPACE}" * 3)
    # auto.SendKeys("{SPACE}" * 3)
    

def main_appointment(hisno_list: list[str]):
    if type(hisno_list) is not list:
        hisno_list = [hisno_list]
    else:
        hisno_list = list(hisno_list)

    window_main = auto.WindowControl(searchDepth=1, AutomationId="frmPatList")
    window_pending(CONFIG['PROCESS_ID'], pending_control=window_main)

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


def main_retrieve(hisno):
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


def main_ditto(hisno: str): # TODO 改成hisno_list版本
    '''
    Ditto功能
    '''
    # start_point check
    window_main = auto.WindowControl(searchDepth=1, SubName='台北榮民總醫院', AutomationId="frmPatList")
    window_main = window_search(window_main)
    if window_main is None:
        auto.Logger.WriteLine("No window frmPatList", auto.ConsoleColor.Red)
        return False
    
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
    
    # # FIXME
    # # 處理TOCC警告 => 這應該是隨機彈窗 會造成invoke使用後錯誤
    # window_tocc = auto.WindowControl(searchDepth=2, AutomationId="dlgNewTOCC")
    # if window_tocc.Exists(maxSearchSeconds=2, searchIntervalSeconds=0.2):
    #     window_tocc.CheckBoxControl(Depth=2, AutomationId="ckbAllNo").GetTogglePattern().Toggle()
    #     window_tocc.ButtonControl(Depth=2, AutomationId="btnOK").GetInvokePattern().Invoke()
    # else:
    #     auto.Logger.WriteLine(f"NOT EXIST: window_tocc", auto.ConsoleColor.Yellow)

    # # 健康行為登錄
    # window_dlg = auto.WindowControl(searchDepth=2, AutomationId="dlgSMOBET")
    # if window_dlg.Exists(maxSearchSeconds=2, searchIntervalSeconds=0.2):
    #     window_dlg.GetWindowPattern().Close()
    # else:
    #     auto.Logger.WriteLine(f"NOT EXIST: dlgSMOBET", auto.ConsoleColor.Yellow)

    # # 處理一堆警告視窗 => #TODO 我沒跳出這個 要保留彈性或是使用鍵盤處理?
    # window_warn = auto.WindowControl(searchDepth=2, AutomationId="dlgWarMessage")
    # if window_warn.Exists(maxSearchSeconds=2, searchIntervalSeconds=0.2):
    #     c_button_ok = window_warn.ButtonControl(searchDepth=1, AutomationId="OK_Button", SubName="繼續")
    #     c_button_ok.GetInvokePattern().Invoke()
    # else:
    #     auto.Logger.WriteLine(f"NOT EXIST: window_warn", auto.ConsoleColor.Yellow)
    # # FIXME
    
    
    # TODO 以下拆開成另一個函數? 這樣可以幫助追蹤進度?
    # 進到ditto視窗
    window_ditto = auto.WindowControl(searchDepth=1, AutomationId="frmDitto")
    window_pending(CONFIG['PROCESS_ID'], pending_control=window_ditto, excluded_control=window_main) # FIXME
    
    # TODO 可刪除
    # window_ditto = window_search(window_ditto)
    # if window_ditto is None:
    #     auto.Logger.WriteLine("No window frmDitto", auto.ConsoleColor.Red)
    #     return False

    # TODO 可刪除
    # # 藥物過敏的視窗是在ditto視窗下
    # window_allergy = window_ditto.WindowControl(searchDepth=1, AutomationId="dlgDrugAllergyDetailAndEdit")
    # if window_allergy.Exists(maxSearchSeconds=1.0, searchIntervalSeconds=0.2):
    #     window_allergy.ButtonControl(Depth=3, SubName='無需更新', AutomationId="Button1").GetInvokePattern().Invoke()
    # else:
    #     auto.Logger.WriteLine(f"NOT EXIST: window_allergy", auto.ConsoleColor.Yellow)

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


def main_excluded_hisno_list(hisno_list):
    exclude_hisno_list = []
    window_main = auto.WindowControl(searchDepth=1, SubName='台北榮民總醫院', AutomationId="frmPatList")
    window_main = window_search(window_main)
    if window_main is None:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT EXIST WINDOW ENTRY", auto.ConsoleColor.Red)
        return False
    
    datagrid_patient = window_main.TableControl(searchDepth=1, SubName='DataGridView', AutomationId="dgvPatsList")
    patient_list_values = datagrid_values(datagrid=datagrid_patient)
    for row in patient_list_values:
        if row[3] in hisno_list and row[9]=='是': # row[3]表示病歷號；row[9]表示暫存欄位 => 未來應改成column_name搜尋方式
            exclude_hisno_list.append(row[3])
    
    return exclude_hisno_list


def package_open(index: int = -1, search_term: str = None):
    '''
    點擊組套功能(可以使用index[起始為0且需+3]或是用search term去搜尋組套視窗的項目)
    '''
    try:
        # Window SOAP 為起點
        window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
        window_soap = window_search(window_soap)
        if window_soap is None:
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT EXIST WINDOW ENTRY", auto.ConsoleColor.Red)
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
        return True
    except Exception as e:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|{e}", auto.ConsoleColor.Red)
        return False


def package_iol_ovd(iol, ovd):
    '''
    select IOL and OVD
    '''
    # 組套第二視窗:frmPkgDetail window
    window_pkgdetail = auto.WindowControl(searchDepth=1, AutomationId="frmPkgDetail")
    window_pkgdetail = window_search(window_pkgdetail)
    if window_pkgdetail is None:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|NOT EXIST WINDOW ENTRY", auto.ConsoleColor.Red)
        return False
    
    c_datagrid_pkgorder = window_pkgdetail.TableControl(searchDepth=1, AutomationId="dgvPkgorder")
    
    # search_datagrid for target item
    target_list = []
    if c_datagrid_pkgorder.Exists():
        target_list = datagrid_search([iol, ovd], c_datagrid_pkgorder, only_one=False)
        if len(target_list) < 2:
            auto.Logger.WriteLine(f"LOSS OF RETURN: IOL:{iol}|OVD:{ovd}", auto.ConsoleColor.Red)
            auto.Logger.WriteLine(f"{[control.GetLegacyIAccessiblePattern().Value for control in target_list]}", auto.ConsoleColor.Red)
    else:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|No datagrid dgvPkgorder", auto.ConsoleColor.Red)
        return False
    
    # click_datagrid
    residual_list = click_datagrid(c_datagrid_pkgorder, target_list=target_list)
    # 測試失敗紀錄: legacy.select
    # search_datagrid(c_datagrid_pkgorder, [iol])[0].GetLegacyIAccessiblePattern().Select(8)  # 無法被select不知道為何
    # search_datagrid(c_datagrid_pkgorder, [ovd])[0].GetLegacyIAccessiblePattern().Select(8)  # 無法被select不知道為何
    # c_datalist_pkgorder = c_datagrid_pkgorder.GetChildren()
    # c_datalist_pkgorder[3].GetLegacyIAccessiblePattern().Select(8)
    # c_datalist_pkgorder[8].GetLegacyIAccessiblePattern().Select(8)

    # confirm
    window_pkgdetail.ButtonControl(searchDepth=1, AutomationId="btnPkgDetailOK").GetInvokePattern().Invoke()


def order_modify_side(side: str = None): 
    # TODO 要能修改個別orders的側別和計價
    if side.strip().upper() == 'OD':
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
    return True


def drug(drug_list):
    window_soap = auto.WindowControl(searchDepth=1, AutomationId="frmSoap")
    window_soap = window_search(window_soap)
    if window_soap is None:
        auto.Logger.WriteLine("No window frmSoap", auto.ConsoleColor.Red)
        return False
    
    # 走藥物修改再加藥防止沒有診斷時不能加藥
    window_soap.ButtonControl(searchDepth=1, AutomationId="btnSoapAlterMed").GetInvokePattern().Invoke()
    
    drug_delete(drug_list)
    for i in range(int(len(drug_list)/4)+1):
        if i == int(len(drug_list)/4):
            split_drug_list = drug_list[i*4:]
        split_drug_list = drug_list[i*4:i*4+4]
        drug_add(split_drug_list)
    drug_modify(drug_list)

    # 點擊返回主畫面
    window_altermed = auto.WindowControl(searchDepth=1, AutomationId="dlgAlterMed")
    window_altermed.ButtonControl(searchDepth=1, AutomationId="btnReturn").GetInvokePattern().Invoke()

    return True
    

def drug_add(drug_list):  
    try:
        added_list = []
        # 藥物修改window
        window_altermed = auto.WindowControl(searchDepth=1, AutomationId="dlgAlterMed")
        datagrid = window_altermed.TableControl(searchDepth=1, Name="DataGridView")
        for drug in drug_list:
            res = datagrid_search(drug['name'], datagrid, '藥名', skip=drug['same_index'], only_one=True)
            if len(res) == 0:
                added_list.append(drug['name'])
        
        if TEST_MODE:
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Added List:{added_list}", auto.ConsoleColor.Yellow)

        # 點擊加藥
        window_altermed.ButtonControl(searchDepth=1, AutomationId="btnDrugList").GetInvokePattern().Invoke()
        # 進入druglist window
        window_druglist = auto.WindowControl(searchDepth=1, AutomationId="frmDrugListExam")
        
        # 輸入藥名: 搜尋框最多10個字元
        for i, drug in enumerate(added_list):
            window_druglist.EditControl(AutomationId=f"TextBox{i}").GetValuePattern().SetValue(drug[:10])
        # 搜尋按鈕
        window_druglist.ButtonControl(AutomationId="btnSearch").GetInvokePattern().Invoke()
        # 選擇datagrid內藥物項目
        c_datagrid_druglist = window_druglist.TableControl(Depth=3, AutomationId="dgvDrugList")
        target_list = datagrid_search(added_list, c_datagrid_druglist, only_one=False)
        click_datagrid(c_datagrid_druglist, target_list)
        # for i in target_list:
        #     click_retry(i)  # select可以選擇到欄位但要有點擊才能真的加藥 # TODO 如果待選選項超過一頁範圍應設計類似datagrid_click捲動機制

        # 點擊確認 => 選擇藥物的確認
        window_druglist.ButtonControl(Depth=3, AutomationId="btnAdd").GetInvokePattern().Invoke()
    except Exception as e:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|{e}", auto.ConsoleColor.Red)
        return False

def drug_delete(drug_list):
    try:
        deleted_list = []
        # 藥物修改window
        window_altermed = auto.WindowControl(searchDepth=1, AutomationId="dlgAlterMed")
        datagrid = window_altermed.TableControl(searchDepth=1, Name="DataGridView")
        if datagrid.Exists():
            # 藥物編輯視窗內的datagrid需要先點擊一下後再重新抓一次datagrid才不會出現"資料列 -1"錯誤
            children = datagrid.GetChildren()
            click_blockinput(children[1], waitTime=0.2)
        
        retry = 5
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
        for i in children:
            if i.Name == '上方資料列':
                tmp = i.GetChildren()
                for index, j in enumerate(tmp):
                    if j.Name == '藥名':
                        column_index = index
                        break
                break
        
        for row in children:
            if ('資料列' in row.Name) and (row.Name != '上方資料列') : # 有資料的列才做判斷
                match = row.GetLegacyIAccessiblePattern().Value.lower().replace('(null)', '').split(';')[column_index]
                found = False
                for drug in drug_list:
                    if drug['name'].lower() in match: # TODO 匹配過的是否應該後續匹配剔除 => 這樣可以確保藥名相同的數量是和drug_list一致的
                        found = True
                        break
                if found == False:
                    deleted_list.append(row)
        
        if TEST_MODE:
            auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|Deleted list:{[drug.Name for drug in deleted_list]}", auto.ConsoleColor.Yellow)

        # 點擊選擇項目
        for deleted_item in deleted_list:
            click_blockinput(deleted_item)
            if TEST_MODE:
                print(deleted_item.Name)
        # click_datagrid(datagrid, deleted_list) # FIXME 無法選擇需要刪除的項目 => 會出現資料列 -1
        # 點擊刪除
        click_blockinput(window_altermed.ButtonControl(searchDepth=1, AutomationId="btnDelete"))
        # window_altermed.ButtonControl(searchDepth=1, AutomationId="btnDelete").GetInvokePattern().Invoke() => 應該會失效
        return True
    except Exception as e:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|{e}", auto.ConsoleColor.Red)
        return False
    

def drug_modify(drug_list):
    try:
        window_altermed = auto.WindowControl(searchDepth=1, AutomationId="dlgAlterMed")
        
        # 修改藥物頻次
        for drug in drug_list:
            if drug['default']: # 跳過要刪除或是使用預設設定的藥物
                continue

            c_modify = window_altermed.TabControl(searchDepth=1, AutomationId="TabControl1").PaneControl(searchDepth=1, AutomationId="TabPage1")
            # c_charge = c_modify.ListControl(searchDepth=1, AutomationId="ListBoxType").ListItemControl(SubName = "自購").GetSelectionItemPattern().Select() #FIXME 目前這樣使用會出錯，不知為何
            if drug['dose'] != '':
                c_dose = c_modify.ComboBoxControl(searchDepth=1, AutomationId="ComboDose").GetValuePattern().SetValue(drug['dose'])
            if drug['frequency'] != '':
                c_freq = c_modify.ComboBoxControl(searchDepth=1, AutomationId="ComboFreq").GetValuePattern().SetValue(drug['frequency'])
            if drug['route'] != '':
                c_route = c_modify.ComboBoxControl(searchDepth=1, AutomationId="ComboRout").GetValuePattern().SetValue(drug['route'])
            if drug['duration'] != '':
                c_duration = c_modify.ComboBoxControl(searchDepth=1, AutomationId="ComboDur").GetValuePattern().SetValue(drug['duration'])
            if (drug['dose'] != '') or (drug['frequency'] != '') or (drug['route'] != '') or (drug['duration'] != ''):
                c_datagrid = window_altermed.TableControl(searchDepth=1, Name="DataGridView")
                target_list = datagrid_search(drug['name'], c_datagrid, skip=drug['same_index'], only_one=True)
                click_blockinput(target_list[0])
                click_blockinput(window_altermed.ButtonControl(searchDepth=1, AutomationId="btnModify"))
        # 只要對修改數據有任何input，選擇的datagrid就會跳成-1
        # 目前測試資料框使用setvalue或sendkey都可，但選擇藥物和送出都必須使用click()，流程必須是先設定好要更改的資料，再refind datagrid，然後點藥和送出都必須是click()
        # window_altermed.ButtonControl(searchDepth=1, AutomationId="btnModify").GetInvokePattern().Invoke()
    except Exception as e:
        auto.Logger.WriteLine(f"{inspect.currentframe().f_code.co_name}|{e}", auto.ConsoleColor.Red)
        return False


def select_ivi(charge): # TODO
    charge = charge.upper()
    if charge == 'SP-A':
        data = get_patient_data()
        if '榮' in data['charge'] or '將' in data['charge']: # 榮民選擇
            package_open(index=33)
        else:
            package_open(index=34)
    elif charge == 'NHI' or charge == 'SP-1' or charge == 'SP-2' or charge == 'DRUG-FREE':
        package_open(index=35)
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


def diagnosis_cata(df_schedule, config_schedule, hisno, side, date): 
    df_selected_dict = df_schedule.loc[[hisno], :].to_dict('records')[0]
    
    surgery = ''
    # side = ""
    iol = df_selected_dict[config_schedule['COL_IOL']].strip()
    final = df_selected_dict[config_schedule['COL_FINAL']].strip()
    
    #處理lensx+術式
    if df_selected_dict[config_schedule['COL_LENSX']].strip().lower() == 'lensx':
        surgery = surgery + 'LenSx-'
    if df_selected_dict[config_schedule['COL_OP']].lower().find('ecce') > -1:
        surgery = surgery + 'ECCE-IOL'
    else:
        surgery = surgery + 'Phaco-IOL'    

    # 處理final
    if float(final) >=0:
        final = '+'+str(final)+'D'
    else:
        final = str(final)+'D'

    # # 處理target => 如果沒有target欄位
    # target = df_selected_dict[config_schedule['COL_TARGET']].strip()
    # if config_schedule['COL_TARGET'].strip() == '':
    #     target = ''
    # else:
    #     target = str(df_selected_dict[config_schedule['COL_TARGET']])
    #     if target.strip() == '':
    #         diagnosis = diagnosis + ')' # 收尾用
    #     else:
    #         diagnosis = diagnosis + f" T:{target})"

    diagnosis = f"s/p {surgery} {side}({iol} {final}) on {date}"

    return diagnosis


def diagnosis_ivi(df_selected_dict, config_schedule, date): # FIXME 尚未更新
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


def soap_save(backtolist = True):
    '''
    存檔+是否跳出功能
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
    

def soap_confirm(mode=0):
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
    window.GetWindowPattern().Close()
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

def gsheet_acc(dr_code: str):
    '''
    Input: short code of account. Ex:4123
    Output: return dictionary of {'ACCOUNT':...,'PASSWORD':...,'NAME':...} 
    '''
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_ACC)
    dr_code = str(dr_code).upper()
    selector = df['ACCOUNT'].str.contains(dr_code, case = False)
    selected_df = df.loc[selector,:]
    if len(selected_df) == 0: # 資料庫中沒有此帳密
        auto.Logger.WriteLine(f"USER({dr_code}) NOT EXIST IN CONFIG", auto.ConsoleColor.Red)
        return None, None
        
    elif len(selected_df) > 1:
        auto.Logger.WriteLine(f"MORE THAN ONE RECORD: {dr_code}, WILL USE THE FIRST ONE")

    result = selected_df.iloc[0,:].to_dict() #df變成series再輸出成dict
    return result['ACCOUNT'], result['PASSWORD']
    


def gsheet_ovd(dr_code: str) -> str:
    '''
    Input: list of id. Ex:4123
    Output: return ovd choice
    '''
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_OVD)
    
    default = df.loc[df['INDEX'] == CONFIG['DEFAULT'], 'ORDER'].values[0]
    
    dr_code=str(dr_code).upper() # case insensitive and compare in string format
    selector = (df['INDEX'].str.lower()==dr_code) # case insensitive and compare in string format
    selected_df = df.loc[selector,'ORDER']
    if len(selected_df) == 0:
        return default # 如果找不到資料使用預設的參數
    else:
        return selected_df.values[0]


def gsheet_iol(iol_input: str):
    '''
    Input: iol that recorded on the surgery schedule
    Output: (formal iol, isNHI), return the formal iol term, which is used in the OPD system and isNHI
    '''
    iol_map = gc.get_col_dict(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_IOL)
    final_keyword = None
    final_length = -1
    for key in iol_map.keys():
        if key.lower() in iol_input.lower(): # 用欄位key直接搜一次
            if len(key) > final_length:
                final_keyword = key
                final_length = len(key)
        for search_word in iol_map[key]: # 用每個欄位下的關鍵字搜尋
            if search_word.lower() in iol_input.lower():
                if len(search_word) > final_length:
                    final_keyword = key
                    final_length = len(search_word)
    return final_keyword


def gsheet_drug_to_druglist(df: pandas.DataFrame, side: str):
    '''
    transform the gsheet drugs data to specific form of drug list
    Example: drug_list = [{'name': 'Cravit oph sol', 'charge': '', 'dose': '', 'frequency': 'QID', 'route': '', 'duration': '7', 'eyedrop': 1}, {'name': 'Scanol tab', 'charge': '', 'dose': '', 'frequency': 'QIDPRN', 'route': '', 'duration': '1', 'eyedrop': 0}]
    '''
    tag_oral = CONFIG['TAG_ORAL'][0]
    drug_list = []

    for column_name in df.columns:    
        # 跳過INDEX行
        if column_name =='INDEX':
            continue
        # 跳過沒有資料的欄位
        if df[column_name].values[0] =='': # 空格表示沒有使用該藥物
            continue
        
        # 資料結構
        result = {
            'name': '',
            'charge': '',
            'dose': '',
            'frequency': '',
            'route': '',
            'duration': '',
            'eyedrop': True,
            'default': False,
            'same_index': 0 # 若有藥名相同 => 此參數會=個數-1
        }
    
        # 處理是否為口服藥物
        if tag_oral in column_name:
            result['eyedrop'] = False
            result['name'] = column_name.replace(tag_oral,'')
        else:
            result['eyedrop'] = True
            result['name'] = column_name
            result['route'] = side
        
        # 轉換內部資料
        text_in_cell = df[column_name].values[0]
        if text_in_cell == CONFIG['DEFAULT']: # 使用藥物的預設設定
            result['default'] = True
            drug_list.append(result)
        else:
            orders = text_in_cell.split('+') # 同種藥物可以用'+'開立多筆
            for i, order in enumerate(orders):
                result_copy = result.copy()
                result_copy['same_index'] = i

                if '$' in order: # 處理自費項目
                    result_copy['charge'] = 'Y' # FIXME 需要驗證這樣能否使用
                    order = order.replace('$','')
                if '#' in order: # 處理顆數
                    tmp = order.split('#') 
                    result_copy['dose'] = tmp[0]
                    order = tmp[-1]
                if '*' in order: # 處理天數
                    tmp = order.split('*')
                    if len(tmp) != 1: 
                        result_copy['frequency'] = tmp[0]
                        result_copy['duration'] = tmp[-1]
                    else:
                        result_copy['duration'] = tmp[0]
                else:
                    if len(order) !=0:
                        result_copy['frequency'] = order

                drug_list.append(result_copy)
    
    return drug_list


def gsheet_drug(index: str, side: str):
    '''
    Input: index ex: 4033
    Output: return drug list, which can be input to the select_precription
    '''
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_DRUG)    
    
    default = df.loc[df['INDEX'] == CONFIG['DEFAULT'],:]
    
    index=str(index).lower() # case insensitive and compare in string format
    selector = (df['INDEX'].str.lower()==index) # case insensitive and compare in string format
    selected_df = df.loc[selector,:]
    if len(selected_df) == 0:
        return gsheet_drug_to_druglist(default, side) #如果找不到資料使用預設的參數
    else:
        return gsheet_drug_to_druglist(selected_df, side)


def gsheet_config_surgery(dr_code: str) -> dict:
    '''
    取得set_surgery的資料, 回傳{column_name:value,...}
    value皆為字串形式
    '''
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_SURGERY)
    
    dr_code=str(dr_code).upper()
    selected_df = df.loc[(df['VS_CODE']==dr_code),:]

    if len(selected_df) == 0: # 找不到資料
        auto.Logger.WriteLine(f"NOT EXIST: {dr_code} in {gsheet.GSHEET_SPREADSHEET}||{gsheet.GSHEET_WORKSHEET_SURGERY}", auto.ConsoleColor.Red) 
    elif len(selected_df) == 1: # 一個選項
        return selected_df.iloc[0,:].to_dict()
    elif len(selected_df) > 1: # 多個選項
        auto.Logger.WriteLine(f"==MORE than 2 same VS_CODE==", auto.ConsoleColor.Red) # 找不到資料
        print("選項:\t組套名稱")
        for i,index in enumerate(selected_df.loc[:,'INDEX']):
            print(f"{i}:\t{index}")
        choice = input("請輸入選項編號: ")
        return selected_df.iloc[int(choice),:].to_dict()
    else:
        auto.Logger.WriteLine(f"Something Wrong in {inspect.currentframe().f_code.co_name}", auto.ConsoleColor.Red)
    
    return None


def gsheet_config_ivi(index: str):
    '''
    取得set_ivi的資料, 回傳{column_name:value,...}
    '''
    df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_IVI)
    selected_df = df.loc[(df['INDEX']==str(index)),:]

    if len(selected_df) == 0: # 找不到資料
        auto.Logger.WriteLine(f"NOT EXIST: {index} in {gsheet.GSHEET_SPREADSHEET}||{gsheet.GSHEET_WORKSHEET_SURGERY}", auto.ConsoleColor.Red) 
    elif len(selected_df) == 1: # 一個選項
        return selected_df.iloc[0,:].to_dict()
    elif len(selected_df) > 1: # 多個選項
        auto.Logger.WriteLine(f"==MORE than 2 same INDEX==", auto.ConsoleColor.Red) # 找不到資料
        print("選項:\t組套名稱")
        for i,index in enumerate(selected_df.loc[:,'INDEX']):
            print(f"{i}:\t{index}")
        choice = input("請輸入選項編號: ")
        return selected_df.iloc[int(choice),:].to_dict()
    else:
        auto.Logger.WriteLine(f"Something Wrong in {inspect.currentframe().f_code.co_name}", auto.ConsoleColor.Red)
    
    return None


def gsheet_schedule_surgery(config_schedule): # TODO 未考慮完成
    '''
    (其他手術)依照config_schedule資訊取得對應的刀表內容且輸出讓使用者確認
    '''
    auto.Logger.WriteLine(f"== INDEX:{config_schedule['INDEX']}|VS_CODE:{config_schedule['VS_CODE']}|SPREADSHEET:{config_schedule['SPREADSHEET']}|WORKSHEET:{config_schedule['WORKSHEET']} ==", auto.ConsoleColor.Yellow)
    while(1):
        df = gc.get_df_select(config_schedule['SPREADSHEET'], config_schedule['WORKSHEET'])
        print(df.reset_index()[[config_schedule['COL_HISNO'], config_schedule['COL_NAME'], config_schedule['COL_DIAGNOSIS'], config_schedule['COL_OP']]])
        check = input('Confirm the above-mentioned information(yes:Enter|no:n)? ')
        if check.strip() == '':
            return df


def gsheet_schedule_cata(config_schedule):
    '''
    (CATA)依照config_schedule資訊取得對應的刀表內容且輸出讓使用者確認
    '''
    auto.Logger.WriteLine(f"== INDEX:{config_schedule['INDEX']}|VS_CODE:{config_schedule['VS_CODE']}|SPREADSHEET:{config_schedule['SPREADSHEET']}|WORKSHEET:{config_schedule['WORKSHEET']} ==", auto.ConsoleColor.Yellow)
    while(1):
        df = gc.get_df_select(config_schedule['SPREADSHEET'], config_schedule['WORKSHEET'])
        print(df.reset_index()[[config_schedule['COL_HISNO'], config_schedule['COL_NAME'], config_schedule['COL_LENSX'], config_schedule['COL_IOL']]])
        check = input('Confirm the above-mentioned information(yes:Enter|no:n)? ')
        if check.strip() == '':
            return df


def gsheet_schedule_ivi(config_schedule):
    '''
    (IVI)依照config_schedule資訊取得對應的刀表內容且輸出讓使用者確認
    '''
    auto.Logger.WriteLine(f"== INDEX:{config_schedule['INDEX']}|VS_CODE:{config_schedule['VS_CODE']}|SPREADSHEET:{config_schedule['SPREADSHEET']}|WORKSHEET:{config_schedule['WORKSHEET']} ==", auto.ConsoleColor.Yellow)
    while(1):
        df = gc.get_df_select(config_schedule['SPREADSHEET'], config_schedule['WORKSHEET'])
        print(df.reset_index()[[config_schedule['COL_HISNO'], config_schedule['COL_NAME'], config_schedule['COL_DIAGNOSIS'], config_schedule['COL_DRUGTYPE'], config_schedule['COL_CHARGE']]])
        check = input('Confirm the above-mentioned information(yes:Enter|no:n)? ')
        if check.strip() == '':
            return df


def gsheet_schedule_side(df_schedule, config_schedule, hisno):
    '''
    取得側別資訊: 刀表側別欄位 > 刀表手術欄位 > 刀表診斷欄位 > 手術排程側別? > 手動輸入?
    '''
    # 抓刀表側別欄位
    if (config_schedule.get('COL_SIDE') is not None) and config_schedule['COL_SIDE'].strip() != '':
        text = df_schedule.loc[hisno, config_schedule['COL_SIDE']].strip()
        if check_op_side(text) is not None:
            return check_op_side(text)
    # 抓刀表手術欄位
    if (config_schedule.get('COL_OP') is not None) and config_schedule['COL_OP'].strip() != '':
        text = df_schedule.loc[hisno, config_schedule['COL_OP']].strip()
        if check_op_side(text) is not None:
            return check_op_side(text)
    # 抓刀表診斷欄位
    if (config_schedule.get('COL_DIAGNOSIS') is not None) and config_schedule['COL_DIAGNOSIS'].strip() != '':
        text = df_schedule.loc[hisno, config_schedule['COL_DIAGNOSIS']].strip()
        if check_op_side(text) is not None:
            return check_op_side(text)
    # 抓手術排程側別 # TODO
    # 手動輸入?
    auto.Logger.WriteLine(f"Side automated recognition failed!", auto.ConsoleColor.Yellow)
    side = input("Which is the side of the surgery (1:OD|2:OS|3:OU)? ").strip()
    if side == '1':
        return 'OD'
    elif side == '2':
        return 'OS'
    elif side == '3':
        return 'OU'


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


def check_op_side(input_string):
    '''
    取得手術側別，回傳'OD'|'OS'|'OU'|None
    '''
    if input_string is None:
        return None
    
    input_string = str(input_string).strip()
    if input_string.upper().find('OD') > -1:
        return 'OD'
    elif input_string.upper().find('OS') > -1:
        return 'OS'
    elif input_string.upper().find('OU') > -1:
        return 'OU'
    else:
        return None


def check_op_type():
    pass
    

def search_opd_program(path_list, filename_list):
    '''
    搜尋路徑順序: 當前目錄=>桌面=>path_list位置找
    '''
    pathlib_list = [Path(), (Path.home()/'Desktop')] # 第一個為當前目錄，第二個為桌面
    for path in path_list:
        pathlib_list.append(Path(path))
    for p in pathlib_list:
        for filename in filename_list:
            result_list = list(p.glob(f'{filename}'))
            if len(result_list) > 0:
                return result_list[0]


def main():
    while True:
        # 搜尋OPD程式位置
        CONFIG['OPD_PATH'] = search_opd_program(CONFIG['OPD_PATH_LIST'], CONFIG['OPD_FILENAME_LIST'])

        # 選擇CATA|IVI mode
        mode = input("Choose the OPD program mode (1:CATA | 2:IVI | 0:hotkey): ")
        if mode.strip() not in ['1','2','0']:
            auto.Logger.WriteLine(f"WRONG MODE INPUT", auto.ConsoleColor.Red)
        elif mode.strip() == '1':
            gc = gsheet.GsheetClient()
            df = gc.get_df(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_SURGERY) # 讀取config
            while True:
                selected_col = ['INDEX','VS_CODE','SPREADSHEET','WORKSHEET']
                selected_df = df.loc[:, selected_col]
                selected_df.index +=1 # 讓index從1開始方便選擇
                selected_df.rename(columns={'INDEX':'組套名'}, inplace=True) # rename column
                # 印出現有組套讓使用者選擇
                print("\n=========================")
                print(selected_df) 
                print("=========================")
                selection = input("請選擇以上profile(0是退回): ").strip()
                if selection == '0': # 等於0 => 退到上一層
                    break
                else:
                    if int(selection) not in selected_df.index:
                        auto.Logger.WriteLine(f"WRONG PROFILE INPUT", auto.ConsoleColor.Red)
                    else:
                        config_schedule = df.loc[int(selection)-1,:].to_dict() # 讀取刀表設定檔
                        if config_schedule['VS_CODE'] == CONFIG['DEFAULT']:
                            dr_code = input("Using default config...please enter the short code of VS (Ex:4123): ")
                        else:
                            dr_code = config_schedule['VS_CODE']

                        # 載入要操作OPD系統的帳密
                        login_id, login_psw = gsheet_acc(dr_code)
                        if login_id is None or login_psw is None:
                            login_id, login_psw = get_id_psw()
                            dr_code = login_id[3:7]
                            
                        # 獲取刀表內容+日期模式
                        date = get_date_today(config_schedule['OPD_DATE_MODE'])
                        df = gsheet_schedule_cata(config_schedule)

                        # 開啟門診程式
                        login(login_id, login_psw, CONFIG['SECTION_CATA'][0], CONFIG['ROOM_CATA'][0])

                        # 將所有病歷號加入非常態掛號
                        hisno_list = df[config_schedule['COL_HISNO']].to_list()
                        main_appointment(hisno_list)
                        
                        # 取得已有暫存list => 之後處理部分會跳過
                        exclude_hisno_list = main_excluded_hisno_list(hisno_list)
                        
                        # 逐一病人處理
                        df.set_index(keys=config_schedule['COL_HISNO'], inplace=True)
                        for hisno in hisno_list:
                            # 跳過已有暫存者
                            if hisno in exclude_hisno_list:
                                auto.Logger.WriteLine(f"Already saved: {hisno}", auto.ConsoleColor.Yellow)
                                continue
                            
                            # ditto
                            res = main_ditto(hisno)
                            if res == False:
                                continue

                            # 選擇phaco模式
                            if df.loc[hisno, config_schedule['COL_LENSX']].strip() == '': # 沒有選擇lensx
                                package_open(index = 31)
                            elif df.loc[hisno, config_schedule['COL_LENSX']].strip().lower() == 'lensx':
                                package_open(index = 32)
                            else:
                                auto.Logger.WriteLine(f"Lensx資訊辨識問題({df.loc[hisno,config_schedule['COL_LENSX']].strip()}) 先以NHI組套處理", auto.ConsoleColor.Red)
                                package_open(index = 31)
                            
                            # 取得側別資訊
                            side = gsheet_schedule_side(df_schedule=df, config_schedule=config_schedule, hisno=hisno)

                            # 在Subject框內輸入手術資訊 => 組合手術資訊
                            diagnosis = diagnosis_cata(df_schedule=df, config_schedule=config_schedule, hisno=hisno, side=side, date=date)
                            set_S(diagnosis)

                            # 取得刀表iol資訊
                            iol = df.loc[hisno, config_schedule['COL_IOL']].strip()
                            iol_search_term = gsheet_iol(iol) # 尋找刀表IOL資訊的正式搜尋名稱
                            if iol_search_term is None:
                                auto.Logger.WriteLine(f"無此IOL({iol})登錄", auto.ConsoleColor.Red)
                                continue

                            # 取得該DOC常用OVD
                            ovd = gsheet_ovd(dr_code)
                            
                            # 打開package
                            if iol_search_term in CONFIG['NHI_IOL']:
                                package_open(index=29)  # NHI IOL
                            else:
                                package_open(index=30)  # SP IOL
                            # IOL和OVD package設定
                            package_iol_ovd(iol=iol_search_term, ovd=ovd)
                            # 修改order的side
                            order_modify_side(side)
                            
                            # 處理藥物
                            drug_list = gsheet_drug(dr_code, side)
                            drug(drug_list)

                            # 暫存退出
                            soap_save()
        
        elif mode.strip() == '2': # IVI  # FIXME 有些函數尚未更新
            # 輸入要操作OPD系統的帳密
            dr_code = input("Please enter the short code of account (Ex:4123): ")
            dr_code, login_id, login_psw = gsheet_acc(dr_code)

            # 判斷程式運行與否
            running, pid = process_exists(CONFIG['PROCESS_NAME'][0])

            # 使用者輸入: 獲取刀表+日期模式
            config_schedule = gsheet_config_ivi(0) # 使用共用組套
            date = get_date_today(config_schedule['OPD_DATE_MODE'])
            df = gsheet_schedule_ivi(config_schedule)

            # 開啟門診程式
            if running:
                auto.Logger.WriteLine("OPD program is running", auto.ConsoleColor.Yellow)
                login_change_opd(login_id, login_psw, CONFIG['SECTION_PROCEDURE'][0], CONFIG['ROOM_PROCEDURE'][0]) 
            else:
                login_all(CONFIG['OPD_PATH'],login_id, login_psw, CONFIG['SECTION_PROCEDURE'][0], CONFIG['ROOM_PROCEDURE'][0])

            # 將所有病歷號加入非常態掛號
            hisno_list = df[config_schedule['COL_HISNO']].to_list()
            main_appointment(hisno_list)

            # 逐一病人處理
            df.set_index(keys=config_schedule['COL_HISNO'], inplace=True)
            for hisno in hisno_list:
                # ditto
                main_ditto(hisno)
                
                side = df.loc[hisno, config_schedule['COL_SIDE']].strip()
                charge = df.loc[hisno, config_schedule['COL_CHARGE']].strip()
                drug_ivi = df.loc[hisno, config_schedule['COL_DRUGTYPE']].strip()
                # TODO
                # TODO 要依照charge處理order
                # TODO 要依照charge決定drug_ivi要不要開上去
                # TODO 依照charge決定出單方式? => 兩次出單

                
                # 處理其它藥物
                other_drug_list = gsheet_drug('ivi')
                drug(other_drug_list)

                # 在Subject框內輸入手術資訊 => 要先組合手術資訊
                diagnosis = diagnosis_ivi(df.loc[[hisno], :].to_dict('records')[0], config_schedule, date)
                set_S(diagnosis)

                # 暫存退出
                soap_save()


TEST_MODE = False
CONFIG = {}

UPDATER_OWNER = 'zmh00'
UPDATER_REPO = 'vghbot'
UPDATER_FILENAME = 'opd'
UPDATER_VERSION_TAG = 'v2.4'


gc = gsheet.GsheetClient()
CONFIG.update(gc.get_col_dict(gsheet.GSHEET_SPREADSHEET, gsheet.GSHEET_WORKSHEET_CONFIG))
CONFIG['DEFAULT'] = CONFIG['DEFAULT'][0]

auto.uiautomation.SetGlobalSearchTimeout(10)  # 應該使用較長的timeout來防止電腦反應太慢，預設就是10秒
auto.uiautomation.DEBUG_SEARCH_TIME = TEST_MODE 


if __name__ == '__main__':
    if TEST_MODE == False:
        u = updater_cmd.Updater_github(UPDATER_OWNER, UPDATER_REPO, UPDATER_FILENAME, UPDATER_VERSION_TAG)
        if u.start() == False:
            sys.exit()
        
        if auto.IsUserAnAdmin():
            main()
        else:
            print('RunScriptAsAdmin', sys.executable, sys.argv)
            auto.RunScriptAsAdmin(sys.argv)
    else: 
        main()


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

