# import usually use library
import datetime
import logging
import pandas
import pyautogui
import numpy as np
import threading
# Work with clipboard
import pyperclip
import time
import win32clipboard
from PIL import Image
from io import BytesIO
from logging import info, critical
import pywinauto
from pywinauto import Application, Desktop
import sys
from data_collectors import run_query, run_multi_query
from LM_Control import run_lm_control
import webbrowser as w
import os
from json import dumps
from httplib2 import Http

logging.basicConfig(format='%(asctime)s %(message)s', level=logging.CRITICAL)

program_path = r"C:\Program Files\WindowsApps\Microsoft.MicrosoftPowerBIDesktop_2.105.1143.0_x64__8wekyb3d8bbwe\bin\PBIDesktop.exe"
pbi_path = r"G:\My Drive\Ninjavan Project\Dataset\Dataset Daily Task Fleet.pbix"
pbi_name = "Dataset Daily Task Fleet - Power BI Desktop"
# 2. List report
col_list = ['report', 'page_name', 'caption', 'report_name', 'x1', 'y1', 'x2', 'y2', 'hook_message', 'link_saved_data', 'export_folder']
df_report = pandas.read_csv(r'G:\My Drive\Ninjavan Project\Send_report_Bot_id_52717207\pbi_report_data.csv', low_memory=False, index_col='report', usecols=col_list)
export_path = r"G:/My Drive/Ninjavan Data/Fleet/Daily Task"
keys = pandas.read_csv(r'G:/My Drive/Ninjavan Project/Send_report_Bot_id_52717207/keys.csv', index_col='key_name')
# 3. Channel
zalo_GTSC = "HN DP GTSC"
zalo_FM = "OPEX - DP : First Mile"
zalo_LM = "OPEX - DP : Last Mile"
zalo_Cloud = "Cloud của tôi"
chat_test = "Just a little test"
chat_vol = "Dự Báo Vol Pickup - Deli : Fleet - WH"
chat_FM = "Firstmile KPI"
chat_LM = "Lastmile KPI"
browser_list = [r'Google Chrome', r'Microsoft\u200b Edge']
# 4. Position
zalo_send_file_button_position = (533, 959)
box_gg_chat = (1037, 969)
box_zalo_web = (987, 1014)
# 5. PBI web link
DAILY_TASK_LINK = keys.loc['DAILY_TASK_LINK']['key']
DATASET_FLEET_LINK = keys.loc['DATASET_FLEET_LINK']['key']



url_chat_test = keys.loc['gchat_test']['key']
url_chat_vol = keys.loc['gchat_vol']['key']
url_chat_FM = keys.loc['gchat_FM']['key']
url_chat_LM = keys.loc['gchat_LM']['key']
channel_send = {chat_test:url_chat_test, chat_vol:url_chat_vol, chat_FM:url_chat_FM, chat_LM:url_chat_LM}

pendingpu_message = r'Truy cập link để lấy thông tin chi tiết những đơn Pending Pickup'
link_pendingpu_detail = keys.loc['link_pendingpu_detail']['key']

list_redash = [[19767, 'LMFN api'], [13798, 'Pending PU'], [14881, 'Inventory'], [13671, 'FMFN API'], [17067, 'Overview_PP'],
                [20865, 'Push_NonShopee_KPI_Tag_5_API'], [20873, 'Push_FSR_LM_KPI'], [20874, 'Push_Shopee_KPI_Tag_20'], [20901, 'Push_NonShopee_PU_DO'], [13671, 'Inbound HN_API'], 
                [19777, 'Recovery_input_API'], [14200, 'GTSC'], [21059, 'Push_FSR_FM_KPI']]



def send_webhook(channel, report_2_send):
    """Hangouts Chat incoming webhook quickstart."""
    message_send = df_report.loc[report_2_send]['hook_message']
    link_send = df_report.loc[report_2_send]['link_saved_data']
    url = channel_send[channel]
    bot_message = {'text': f'{message_send}: {link_send}'}
    message_headers = {'Content-Type': 'application/json; charset=UTF-8'}
    http_obj = Http()
    response = http_obj.request(
        uri=url,
        method='POST',
        headers=message_headers,
        body=dumps(bot_message),
    )
    print(response)
    time.sleep(5)

# A part - general features
def get_running_windows():
    windows = Desktop(backend="uia").windows()
    return [w.window_text() for w in windows]

def open_browser():
    global browser_name
    # Browser always opens
    for name in browser_list:
        try:
            try:
                browser = Application(backend="uia").connect(title_re=f".*{name}.*", timeout=5, found_index=0).window(title_re=f".*{name}.*", found_index=0)
                browser.restore().maximize()
                browser_name = name
                break
            except:
                os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
                # os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
                time.sleep(10)
                w.open('https://chat.zalo.me/')
                time.sleep(1)
                w.open('https://mail.google.com/chat/u/0/#chat/space/AAAAeqSmDqE')
                browser = Application(backend="uia").connect(title_re=f".*{name}.*", timeout=5, found_index=0).window(title_re=f".*{name}.*", found_index=0)
                browser.restore().maximize()
                browser_name = name
                break
        except pywinauto.findwindows.ElementNotFoundError:
            lst = get_running_windows()
            for _ in lst:
                if name in _:
                    browser = Application(backend="uia").connect(title=f"{_}", timeout=5, found_index=0).window(title=f"{_}", found_index=0)
                    browser.restore().maximize()
                    browser_name = name
                    break
    return browser


def open_pbi():
    # Open pbi file or connect exist one
    if pbi_name not in get_running_windows():
        Application(backend="uia").start(f'{program_path} "{pbi_path}"', timeout=20)
        time.sleep(20)
    pbi = Application(backend="uia").connect(title=pbi_name, timeout=5).window(title=pbi_name)
    return pbi


def set_up(web_only = 0):
    global pbi
    global browser
    if web_only == 0:
        pbi = open_pbi()
    browser = open_browser()


# Function 1 to screenshot and send to the clipboard
def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()


# Function 2 to screenshot and send to the clipboard
def screenshot_to_clipboard(position):
    image = pyautogui.screenshot(region=position)
    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    send_to_clipboard(win32clipboard.CF_DIB, data)
    time.sleep(1)


# Get position of an Element
def get_position(w):
    str_ = str(w.rectangle())
    str_ = str_.replace('(', '').replace(')', '')
    list_ = str_.split(',')
    for _, _v in enumerate(list_):
        list_[_] = int(list_[_].strip()[1:])
    x = int((list_[0] + list_[2]) / 2)
    y = int((list_[1] + list_[3]) / 2)
    return [x, y]


# B part - duplicate code, specific functions
def open_tab(tab='chat', channel=None):

    if tab == 'chat':
        key = 'Chat'
        link = 'https://mail.google.com/chat'
    elif tab == 'zalo':
        key = 'Zalo'
        link = 'https://chat.zalo.me/'
    else:
        return None

    browser.maximize()
    time.sleep(1)
    pyautogui.hotkey('ctrl', '1')  # Tránh 1 lỗi
    time.sleep(1)

    try:  # If there is one window's name contains "chat"
        browser.child_window(title_re=f".*{key}.*", control_type="TabItem", found_index=0).click_input()  # Click tab chat
    except:  # Open a new chat
        browser.child_window(title="New Tab", control_type="Button").click_input()  # Open new tab
        time.sleep(0.5)
        pyperclip.copy(link)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        pyautogui.hotkey('enter')
        time.sleep(10)  # Wait tab

    # Open channel in other environments
    if channel is not None and tab == 'chat' and browser_name == 'Edge':
        browser.child_window(title="Just a little test Space Pinned conversation Press tab for more options.", control_type="Hyperlink").click_input()
    elif channel is not None and tab == 'chat' and browser_name == 'Google Chrome':
        browser.child_window(title=channel, control_type='Text', found_index=0).click_input()
    elif channel is not None and tab == 'zalo':
        browser[channel].click_input()


def zalo_click_send_file_button():
    # Zalo has already been opened
    time.sleep(1)
    for _ in range(10):
        pyautogui.click(zalo_send_file_button_position)
        time.sleep(0.5)  # Don't too fast
        try:
            browser.child_window(title='Chọn File', control_type="Text", found_index=0).click_input()
            break
        except pywinauto.findwindows.ElementNotFoundError:
            try:
                print('First option fail')
                browser.child_window(title='Chọn tập tin', control_type="Text", found_index=0).click_input()
                break
            except pywinauto.findwindows.ElementNotFoundError:
                print('Second option fail')
                continue


# Refresh data in pbi
def pbi_refresh_data():
    pbi.maximize()
    pbi.child_window(title="Home", control_type="TabItem", found_index=0).click_input()  # Click Home Tab
    pbi.child_window(title="Refresh", control_type="Button", found_index=0).click_input()  # Python finds out 3 refresh buttons but just the first is true
    run_time_start = datetime.datetime.now()
    # for i in range(120):  # try in about 18 minutes
    while True:
        try:
            pbi.child_window(title="Close", control_type="Button", found_index=1).click_input()  # Refresh success -> there are 2 close buttons
            time.sleep(3)  # Founded -> wait 3s -> click
            info("Refresh success")
            return True
        except:
            cur_time = datetime.datetime.now()
            time_diff = (cur_time - run_time_start).total_seconds()
            if time_diff <= 20*60:
                continue
            else:
                info("Refresh fail")
                pbi.child_window(title="Close", control_type="Button", found_index=0).click_input()  # Close refresh
                return False


def pbi_web_refresh(link, web_only = 0):
    try:
        if web_only == 1:
            set_up(1)
        else:
            set_up()
        browser.restore().maximize()
        time.sleep(1)
        browser.NewTabButton.click_input()
        time.sleep(1)
        pyperclip.copy(f'{link}')
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        pyautogui.hotkey('enter')
        if web_only == 1:
            time.sleep(5)
        time.sleep(5)
        browser.child_window(title='Refresh').click_input()
        browser.child_window(title='Refresh now').click_input()
        time.sleep(2)
    except:
        pass


def export_data_pbi(key):

    if df_report.loc[key]['report_name'] is np.nan:
        return None
    export_folder = df_report.loc[key]['export_folder']
    page_name = df_report.loc[key]['page_name']
    report_name = df_report.loc[key]['report_name']

    pbi.maximize()
    pbi.child_window(title=page_name, control_type='TabItem').click_input()
    # Click report to show "..." button
    x = pbi.child_window(title_re=report_name, control_type="Group", found_index=0)
    x.click_input()
    time.sleep(5)
    x.click_input()
    x.child_window(title="More options", found_index=0).click_input()
    pbi.child_window(title="Export data", control_type="MenuItem").click_input()
    time.sleep(5)  # wait the dialog to show
    file_out = f'{export_path}/{export_folder}/{report_name}.csv'
    file_out = file_out.replace(r"/", chr(92))
    pyperclip.copy(file_out)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pbi.child_window(title="Save As", control_type="Window").child_window(title="Save", control_type="Button",
                                                                          found_index=0).click_input()
    try:
        pbi.child_window(title="Confirm Save As", control_type="Window").child_window(title="Yes",
                                                                                      auto_id="CommandButton_6",
                                                                                      control_type="Button").click_input()
    except:
        pass


def convert_to_excel(filename):
    pandas.read_csv(f'{export_path}/{filename}.csv', low_memory=False).to_excel(f'{export_path}/{filename}.xlsx', index=False, engine='openpyxl')


# Clear message box
def clear_message_box(tab='chat', channel=None):

    if tab == 'chat':
        key = f'Message {channel}'
        position = box_gg_chat
    elif tab == 'zalo':
        key = 'Nhập @'
        position = box_zalo_web
    else:
        return None
    
    try: # Khi message box không có chữ
        browser.child_window(title_re=f'^{key}.*', found_index=0).click_input()
    except: # Khi message box có chữ
        pyautogui.moveTo(position)
        time.sleep(1)
        pyautogui.click()
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(1)
        pyautogui.hotkey('delete')
        time.sleep(1)


# Tag all
def tag_all():
    time.sleep(1)
    pyautogui.write('@')
    time.sleep(2)
    pyautogui.write('all')
    time.sleep(2)
    pyautogui.hotkey('Tab')
    time.sleep(1)


# Send Report
def send_report(key):
    # cursor has already been in the message bot
    today_str = datetime.datetime.strftime(datetime.datetime.now(), "%d/%m/%Y")
    if df_report.loc[key]['caption'] is not np.nan:
        time.sleep(1)
        pyperclip.copy(df_report.loc[key]['caption'] + today_str)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)

    time.sleep(1)    
    pbi.maximize()
    time.sleep(1)
    pbi.child_window(title=df_report.loc[key]['page_name'], control_type='TabItem').click_input()
    time.sleep(5) # Wait pbi to load report
    x1 = df_report.loc[key]['x1']
    y1 = df_report.loc[key]['y1']
    x2 = df_report.loc[key]['x2']
    y2 = df_report.loc[key]['y2']
    screenshot_to_clipboard((x1, y1, x2-x1, y2-y1))
    time.sleep(1)

    # Paste to message box
    browser.maximize()
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(15)  # Wait image to load
    pyautogui.hotkey('enter')
    time.sleep(3)
    # cursor end in the message box


# Send data file
def send_data_to_zalo(key):
    browser.maximize()
    zalo_click_send_file_button()
    time.sleep(10)  # Wait the dialog
    if isinstance(key, str):        
        text = f'{export_path}\ {df_report.loc[key]["report_name"]}.csv'
    elif isinstance(key, list):
        text = f'{export_path}\ '
        for _ in key:
            text += f'"{df_report.loc[_]["report_name"]}.csv" '
    pyperclip.copy(text)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(5)  # Don't too fast
    pyautogui.hotkey('enter')


def send_reports_to_ggchat(channel, report, refresh=True, is_tag_all=True):
    set_up()
    time.sleep(2)
    ref = pbi_refresh_data() if refresh else True
    if not ref:
        critical("Refresh fail, Dont send")
    else:
        open_tab(tab='chat', channel=channel)
        time.sleep(3)
        # clear_message_box(tab='chat')
        if is_tag_all == True:
            tag_all()
        time.sleep(1)
        send_report(report)
        if is_tag_all == True:
            try:  # Click Send now button when tag all in ggchat
                browser.child_window(title="Send now", control_type="Button").click_input()
            except:
                pass

        # # Export data
        # export_data_pbi('lmdetail')
        # time.sleep(1)

        # # Send data file to zalo
        # send_data_to_zalo('lmdetail')
        # time.sleep(1)



# D part - wrap with command_list
def send_vol(refresh=True, send_lm_speed=True, send_vol_lm=False, channel=chat_vol, collect_data=True, command=None):
    print(f'Start send vol, command = {command}')
    list_send = ['vol', 'vol2', 'vol3', 'lmspeed']
    today = datetime.date.today()
    now = datetime.datetime.now()
    if now.hour < 15:
        list_redash.append([20866, 'Push_Shopee_KPI_Tag_12_API'])
    
    if command is not None:
        # optional command
        if 'q' in command:
            collect_data = False
        if '20' in command:
            list_send.append('vollm')

        # end command
        if 't' in command:
            channel=chat_test
            list_send = ['vol', 'vol2', 'vol3', 'lmspeed', 'vollm']
            for reports in list_send:
                if reports == 'vol':
                    is_tag_all = True
                else:
                    is_tag_all = False
                send_reports_to_ggchat(channel=channel, report=reports, refresh=False, is_tag_all=is_tag_all)
                if reports == 'vol3':
                    send_webhook(channel, 'vol')
            print('Test done')
            return None
        # elif 'cn' in command:
        #     send_lm_speed = False
        #     # list_redash = [[13798, 'Pending PU'], [13671, 'FMFN API'], [17067, 'Overview_PP']]
        #     if today.weekday() != 6:
        #         return None
        # else:
        #     if today.weekday() == 6:
        #         return None

    if collect_data:
        run_multi_query(list_redash)
    print('prepare to ref and send')
    pbi_web_refresh(DAILY_TASK_LINK)
    for reports in list_send:
        if reports == 'vol':
            refresh = True
            is_tag_all = True
        else:
            refresh = False
            is_tag_all = False
        send_reports_to_ggchat(channel=channel, report=reports, refresh=refresh, is_tag_all=is_tag_all)
        if reports == 'vol3':
            send_webhook(channel, 'vol')


def send_lm20(channel=chat_LM, refresh=True, collect_data=True, command=None):
    report_to_send = 'lm20'
    print(f'Start send {report_to_send}, command = {command}')
    today = datetime.date.today()

    if command is not None:
        # optional command
        if 'q' in command:
            collect_data = False

        # end command
        if 't' in command:
            channel=chat_test
            send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=False, is_tag_all=True)
            export_data_pbi(report_to_send)
            send_webhook(channel, report_to_send)
            return None
        # else:
        #     if today.weekday() == 6:
        #         return None

    if collect_data:
        # lm = run_lm_slot_2(detail=True)
        run_multi_query(list_redash)
        pbi_web_refresh(DAILY_TASK_LINK)
        send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=refresh, is_tag_all=True)
        export_data_pbi(report_to_send)
        send_webhook(channel, report_to_send)
    else:
        pbi_web_refresh(DAILY_TASK_LINK)
        send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=refresh, is_tag_all=True)
        export_data_pbi(report_to_send)
        send_webhook(channel, report_to_send)


def send_lm12(channel=chat_LM, refresh=True, collect_data=True, command=None):
    report_to_send = 'lm12'
    print(f'Start send {report_to_send}, command = {command}')
    today = datetime.date.today()

    if command is not None:
        # optional command
        if 'q' in command:
            collect_data = False

        # end command
        if 't' in command:
            channel=chat_test
            send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=False, is_tag_all=True)
            export_data_pbi(report_to_send)
            send_webhook(channel, report_to_send)
            return None
        # else:
        #     if today.weekday() == 6:
        #         return None

    if collect_data:
        list_redash.append([20866, 'Push_Shopee_KPI_Tag_12_API'])
        run_multi_query(list_redash)
    pbi_web_refresh(DAILY_TASK_LINK)
    send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=refresh, is_tag_all=True)
    export_data_pbi(report_to_send)
    send_webhook(channel, report_to_send)


def send_lmc(channel=chat_LM, refresh=True, collect_data=True, command=None):
    report_to_send = 'lmc'
    print(f'Start send {report_to_send}, command = {command}')
    today = datetime.date.today()
    if command is not None:
        # optional command
        if 'q' in command:
            collect_data = False

        # end command
        if 't' in command:
            channel=chat_test
            send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=False, is_tag_all=True)
            return None

    if collect_data:
        lm = run_lm_control()
        if lm:
            pbi_web_refresh(DAILY_TASK_LINK)
            send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=refresh, is_tag_all=True)
    else:
        pbi_web_refresh(DAILY_TASK_LINK)
        send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=refresh, is_tag_all=True)


def send_fm(channel=chat_FM, refresh=True, collect_data=True, command=None):
    report_to_send = 'fm'
    print(f'Start send {report_to_send}, command = {command}')
    today = datetime.date.today()
    if command is not None:
        # optional command
        if 'q' in command:
            collect_data = False

        # end command
        if 't' in command:
            channel=chat_test
            report_to_send = 'fmD0'
            send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=False, is_tag_all=True)
            export_data_pbi(report_to_send)
            send_webhook(channel, report_to_send)
            report_to_send = 'fmfsr'
            send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=False, is_tag_all=False)
            export_data_pbi(report_to_send)
            send_webhook(channel, report_to_send)
            return None

    if collect_data:
        run_multi_query(list_redash)
    pbi_web_refresh(DAILY_TASK_LINK)
    report_to_send = 'fmD0'
    send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=refresh, is_tag_all=True)
    export_data_pbi(report_to_send)
    send_webhook(channel, report_to_send)
    report_to_send = 'fmfsr'
    send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=False, is_tag_all=False)
    export_data_pbi(report_to_send)
    send_webhook(channel, report_to_send)



def send_lmfsr(channel=chat_LM, refresh=True, collect_data=True, command=None):
    report_to_send = 'lmfsr'
    print(f'Start send {report_to_send}, command = {command}')
    today = datetime.date.today()
    if command is not None:
        # optional command
        if 'q' in command:
            collect_data = False

        # end command
        if 't' in command:
            channel=chat_test
            send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=False, is_tag_all=True)
            export_data_pbi(report_to_send)
            send_webhook(channel, report_to_send)
            return None


    if collect_data:
        run_multi_query(list_redash)
    pbi_web_refresh(DAILY_TASK_LINK)
    send_reports_to_ggchat(channel=channel, report=report_to_send, refresh=refresh, is_tag_all=True)
    export_data_pbi(report_to_send)
    send_webhook(channel, report_to_send)


def call_func(name_list, *args, **kwargs):
    if len(name_list) == 1:
        globals()[f'send_{name_list[0]}'](*args, **kwargs)
    else:
        for _ in name_list:
            globals()[f'send_{_}'](*args, **kwargs)
