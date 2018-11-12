import wx
import os
import copy
import threading
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.ui import Select

USER_PROFILE_SECTION = "*S"
ITEM_SECTION = "*I"
SITE_URL = ""
PROFILE_DIR = ""
PROFILE_NUM = -1
SITE_URL_KEY = "URL"
PROFILE_DIR_KEY = "PROFILE_DIR"
PROFILE_NUM_KEY = "PROFILE_NUM"
CHROME_DRIVER_PATH = "./chromedriver.exe"
CONFIG_FILE = "./config.txt"

SIZE_S = "S"
SIZE_M = "M"
SIZE_L = "L"
SIZE_XL = "XL"

HTML_ITEM_NAME_PATH = "//*[@id='img_codes']"
HTML_ITEM_SIZE_PATH = "//*[@id='size']"
HTML_LAST_NAME_PATH = "//*[@id='credit_card_last_name']"
HTML_FIRST_NAME_PATH = "//*[@id='credit_card_first_name']"
HTML_MAIL_PATH = "//*[@id='order_email']"
HTML_PHONE_PATH = "//*[@id='order_tel']"
HTML_STATE_PATH = "//*[@id='state']"
HTML_CITY_PATH = "//*[@id='order_billing_city']"
HTML_ADDRESS_PATH = "//*[@id='order_billing_address']"
HTML_ZIP_CODE_PATH = "//*[@id='order_billing_zip']"
HTML_PAYMENT_TYPE_PATH = "//*[@id='payment']"
HTML_CARD_TYPE_PATH = "//*[@id='card-type']"
HTML_CARD_NUMBER_PATH = "//*[@id='cnb']"
HTML_CARD_LIMIT_MONTH_PATH = "//*[@id='credit_card_month']"
HTML_CARD_LIMIT_YEAR_PATH = "//*[@id='credit_card_year']"
HTML_CARD_CVV_PATH = "//*[@id='vval']"
HTML_DELAY_PATH = "//*[@id='checkout_delay']"
HTML_SAVE_BUTTON_PATH = "//*[@id='save']"

ERROR = True
SUCCESS = False


class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title="Drop Target", size=(520, 150))
        self.panel = wx.Panel(self)
        sizer = wx.GridSizer(rows=1, cols=1, gap=(0,0))
        self.text = wx.StaticText(self.panel, -1, "<既存ツールフォルダ>\\data\\tmp.txt をドラッグ＆ドロップしてください")
        sizer.Add(self.text, flag=wx.CENTER | wx.ALIGN_CENTER)
        self.panel.SetSizer(sizer)

        dt = MyFileDropTarget(self)
        self.SetDropTarget(dt)
        #open_paste_window("C:\\Users\\ishimoto\\Desktop\\Key-kun\\クラウドワークス\\マルチサイトペーストツール\\CsvToText_1x01x06_zip\\data\\tmp.txt")
        #self.Close()


class MyFileDropTarget(wx.FileDropTarget):
    def __init__(self, frame):
        wx.FileDropTarget.__init__(self)
        self.frame = frame

    def OnDropFiles(self, x, y, filenames):

        if len(filenames) != 1:
            self.frame.text.SetLabel("ファイルは１つだけ指定してください")
            return False

        if not os.path.isfile(filenames[0]):
            self.frame.text.SetLabel("ファイルをドロップしください")
            return False

        self.frame.Close()
        open_paste_window(filenames[0])
        return True


class ListPanel(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent, wx.ID_ANY)
        header = ("#", "姓", "名", "商品名", "プロファイル")
        self.listctrl = wx.ListCtrl(self, size=(700, 400), style=wx.LC_REPORT | wx.LC_HRULES)
        for col,val in enumerate(header):
            self.listctrl.InsertColumn(col, val)
        self.listctrl.SetColumnWidth(0, 50)
        self.listctrl.SetColumnWidth(3, 300)
        self.listctrl.SetColumnWidth(4, 190)

    def list_order(self, order_dict):
        row = 0
        for key, val in order_dict.items():
            self.listctrl.InsertItem(row, str(row+1))
            self.listctrl.SetItem(row, 1, val.last_name)
            self.listctrl.SetItem(row, 2, val.first_name)
            self.listctrl.SetItem(row, 3, val.item_name)
            self.listctrl.SetItem(row, 4, val.profile_name)
            row = row + 1


class DelayPanel(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent, wx.ID_ANY)
        label = wx.StaticText(self, wx.ID_ANY, "Delay(ms)")
        self.text = wx.TextCtrl(self, wx.ID_ANY, "0")
        layout = wx.BoxSizer(wx.HORIZONTAL)
        layout.Add(label, flag=wx.RIGHT | wx.LEFT, border=10)
        layout.Add(self.text)
        self.SetSizer(layout)


class InformationPanel(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent, wx.ID_ANY)
        self.label = wx.StaticText(self, wx.ID_ANY, "リストからペースト情報を選択してボタンを押してください", size=(1000, PROFILE_NUM*20))
        layout = wx.BoxSizer(wx.HORIZONTAL)
        layout.Add(self.label, flag=wx.TOP | wx.RIGHT | wx.LEFT, border=10)
        self.SetSizer(layout)


class ButtonPanel(wx.Panel):
    def __init__(self, parent):
        self.info_label = None
        self.delay_text = None
        super().__init__(parent, wx.ID_ANY)
        self.button = wx.Button(self, wx.ID_ANY, "プロファイルごとにブラウザを開いてペースト")
        layout = wx.BoxSizer(wx.HORIZONTAL)
        layout.Add(self.button, flag=wx.RIGHT | wx.LEFT, border=10)
        self.SetSizer(layout)

    def bind(self, listctrl, order_dict):
        self.button.Bind(wx.EVT_BUTTON, lambda event: self.site_paste(listctrl, order_dict))

    def site_paste(self, listctrl, order_dict):
        selected_index_list = []
        index = listctrl.GetFirstSelected()
        selected_index_list.append(index)

        if index < 0:
            self.info_label.SetLabelText("プロファイルが選択されていません")
            return

        while len(selected_index_list) < listctrl.GetSelectedItemCount():
            index = listctrl.GetNextSelected(index)
            selected_index_list.append(index)

        if is_duplicate_profile_selected(listctrl, selected_index_list, order_dict):
            self.info_label.SetLabelText("同一のプロファイルが選択されています")
            return

        self.button.Disable()
        self.info_label.SetLabelText("ペースト処理開始。。。")
        thread_list = []
        for i in selected_index_list:
            id = int(listctrl.GetItem(i, 0).GetText())
            site_paste_thread = SitePasteThread(order_dict[id], self.delay_text.GetValue())
            thread_list.append(site_paste_thread)
            site_paste_thread.start()

        output_msg = ""
        for thread in thread_list:
            thread.join()

        for thread in thread_list:
            width = 0
            for err_msg in thread.error_msg_list:
                width = max(width, len(err_msg))

            if len(thread.error_msg_list) != 0:
                output_msg += "プロファイル[%s]の処理が失敗しました。\n" % thread.order_info.profile_name
                frame = wx.Frame(None, title="ERROR", size=(max(700, 10*width), 150))
                panel = wx.Panel(frame)
                sizer = wx.BoxSizer(wx.VERTICAL)

                text = wx.StaticText(panel, -1, "プロファイル[%s]のペースト処理で、以下のエラーが発生しました。SAVEボタンはクリックされていません。\n" % thread.order_info.profile_name)
                sizer.Add(text, flag=wx.CENTER | wx.ALIGN_CENTER, border=10)
                for err_msg in thread.error_msg_list:
                    text = wx.StaticText(panel, -1, err_msg)
                    sizer.Add(text, flag=wx.CENTER | wx.ALIGN_CENTER, border=10)

                panel.SetSizer(sizer)
                frame.Show()

            else:
                output_msg += "プロファイル[%s]の処理が成功しました。\n" % thread.order_info.profile_name

        self.info_label.SetLabelText(output_msg)
        self.button.Enable()


class OrderInfo:
    def __init__(self):
        self.id = None
        self.item_name = None
        self.item_size = None
        self.first_name = None
        self.last_name = None
        self.mail = None
        self.phone = None
        self.state = None
        self.city = None
        self.address = None
        self.zip_code = None
        self.payment_type = None
        self.card_number = None
        self.month = None
        self.year = None
        self.cvv_number = None
        self.profile_path = None
        self.profile_name = None


class SitePasteThread(threading.Thread):
    def __init__(self, order_info, delay_val):
        super(SitePasteThread, self).__init__()
        self.order_info = order_info
        self.delay_val = delay_val
        self.error_msg_list = []

    def run(self):
        if self.copy_profile_to_wk():
            return
        self.exec_selenium()
        self.copy_profile_to_org()

    def send_keys(self, driver, html_id, target_val):
        driver.find_element_by_xpath(html_id).clear()
        driver.find_element_by_xpath(html_id).send_keys(target_val)

    def select_box(self, driver, html_id, target_val, target_name):
        try:
            Select(driver.find_element_by_xpath(html_id)).select_by_visible_text(target_val)
        except:
            self.error_msg_list.append("[%s]のドロップダウンリストに、[%s]は存在しません。" % (target_name, target_val))

    def copy_profile_to_wk(self):
        target_dir = "./profiles/%s" % self.order_info.profile_name

        try:
            if os.path.exists(target_dir):
                os.system("rmdir /s /q \"%s\"" % target_dir)

            if not os.path.exists(self.order_info.profile_path):
                self.error_msg_list.append("プロファイルフォルダ[%s]は存在しません。" % self.order_info.profile_path)
                return ERROR

            os.system("mkdir \"%s\"" % target_dir)
            os.system("xcopy \"%s\" \"%s/Default\" /e /c /i" % (self.order_info.profile_path, target_dir))
            return SUCCESS
        except:
            self.error_msg_list.append("[%s]のコピーまたは削除に失敗しました。すべてのChromeを閉じてから再実行してください。" % self.order_info.profile_path)
            import traceback
            traceback.print_exc()
            return ERROR

    def copy_profile_to_org(self):
        from_dir = "./profiles/%s/Default" % self.order_info.profile_name
        os.system("xcopy \"%s\" \"%s\" /e /c /i /y" % (from_dir, self.order_info.profile_path))

    def exec_selenium(self):
        try:
            target_dir = "%s\\profiles\\%s" % (os.path.dirname(os.path.abspath(__file__)), self.order_info.profile_name)
            options = webdriver.ChromeOptions()
            options.add_argument("user-data-dir=%s" % target_dir)
            driver = webdriver.Chrome(executable_path=CHROME_DRIVER_PATH, chrome_options=options)
            driver.get(SITE_URL)

            sleep(10)
            sleep(3)
            self.send_keys(driver, HTML_ITEM_NAME_PATH, self.order_info.item_name)
            self.send_keys(driver, HTML_ITEM_SIZE_PATH, self.order_info.item_size)
            self.send_keys(driver, HTML_LAST_NAME_PATH, self.order_info.last_name)
            self.send_keys(driver, HTML_FIRST_NAME_PATH, self.order_info.first_name)
            self.send_keys(driver, HTML_MAIL_PATH, self.order_info.mail)
            self.send_keys(driver, HTML_PHONE_PATH, self.order_info.phone)
            self.select_box(driver, HTML_STATE_PATH, self.order_info.state, "都道府県")
            self.send_keys(driver, HTML_CITY_PATH, self.order_info.city)
            self.send_keys(driver, HTML_ADDRESS_PATH, self.order_info.address)
            self.send_keys(driver, HTML_ZIP_CODE_PATH, self.order_info.zip_code)

            if self.order_info.payment_type == "代金引換":
                if not driver.find_element_by_xpath(HTML_PAYMENT_TYPE_PATH).is_selected():
                    driver.find_element_by_xpath(HTML_PAYMENT_TYPE_PATH).click()
                self.select_box(driver, HTML_CARD_TYPE_PATH, "", "支払い方法")
                driver.find_element_by_xpath(HTML_CARD_NUMBER_PATH).clear()
                self.select_box(driver, HTML_CARD_LIMIT_MONTH_PATH, "", "カード有効期限(月)")
                self.select_box(driver, HTML_CARD_LIMIT_YEAR_PATH, "", "カード有効期限(年)")
                driver.find_element_by_xpath(HTML_CARD_CVV_PATH).clear()
            else:
                if driver.find_element_by_xpath(HTML_PAYMENT_TYPE_PATH).is_selected():
                    driver.find_element_by_xpath(HTML_PAYMENT_TYPE_PATH).click()
                self.select_box(driver, HTML_CARD_TYPE_PATH, self.order_info.payment_type, "支払い方法")
                self.send_keys(driver, HTML_CARD_NUMBER_PATH, self.order_info.card_number)
                self.select_box(driver, HTML_CARD_LIMIT_MONTH_PATH, self.order_info.month, "カード有効期限(月)")
                self.select_box(driver, HTML_CARD_LIMIT_YEAR_PATH, self.order_info.year, "カード有効期限(年)")
                self.send_keys(driver, HTML_CARD_CVV_PATH, self.order_info.cvv_number)

            self.send_keys(driver, HTML_DELAY_PATH, self.delay_val)

            if len(self.error_msg_list) == 0:
                driver.find_element_by_xpath(HTML_SAVE_BUTTON_PATH).click()

            #while True:
            #    sleep(5)

            driver.quit()

        except Exception as e:
            self.error_msg_list.append("想定外エラー。調査が必要です。\n")
            import traceback
            self.error_msg_list.append(traceback.format_exc())
            driver.quit()


def open_paste_window(filepath):
    frame = wx.Frame(None, title="Paste list", size=(720, 550+(PROFILE_NUM*20)))
    frame.panel = wx.Panel(frame)
    order_dict = make_order_dict(filepath)
    list_panel = ListPanel(frame.panel)
    list_panel.list_order(order_dict)

    delay_panel = DelayPanel(frame.panel)
    button_panel = ButtonPanel(frame.panel)
    button_panel.bind(list_panel.listctrl, order_dict)
    info_panel = InformationPanel(frame.panel)
    button_panel.info_label = info_panel.label
    button_panel.delay_text = delay_panel.text
    sizer = wx.BoxSizer(wx.VERTICAL)
    sizer.Add(list_panel)
    sizer.Add(delay_panel)
    sizer.Add(button_panel)
    sizer.Add(info_panel)
    frame.panel.SetSizer(sizer)
    frame.Show()


def site_pate_thread(listctrl, selected_index_list, order_dict):
    selected_profile_dict = {}
    for i in selected_index_list:
        id = int(listctrl.GetItem(i, 0).GetText())
        already_found = selected_profile_dict.get(order_dict[id].profile_name, False)
        if already_found:
            print("duplicate ERROR")
            return
        selected_profile_dict[order_dict[id].profile_name] = True


def is_duplicate_profile_selected(listctrl, selected_index_list, order_dict):
    selected_profile_dict = {}
    for i in selected_index_list:
        id = int(listctrl.GetItem(i, 0).GetText())
        already_found = selected_profile_dict.get(order_dict[id].profile_name, False)
        if already_found:
            return True
        selected_profile_dict[order_dict[id].profile_name] = True

    return False


def make_order_dict(filepath):
    id = 1
    order_dict = {}
    profile_no = 1
    for line in open(filepath, "r", encoding="utf-8_sig"):
        line = line[:-1]
        items = line.split("\t")
        if items[0] == USER_PROFILE_SECTION:
            if items[1] == "":
                break
            order_info_tmpl = OrderInfo()
            order_info_tmpl.last_name = items[2]
            order_info_tmpl.first_name = items[3]
            order_info_tmpl.mail = items[9]
            order_info_tmpl.phone = items[8]
            order_info_tmpl.state = items[5]
            order_info_tmpl.city = items[6]
            order_info_tmpl.address = items[7]
            order_info_tmpl.zip_code = items[4]
            order_info_tmpl.payment_type = items[13]
            order_info_tmpl.card_number = items[14]
            order_info_tmpl.month = "%02d" % int(items[15])
            order_info_tmpl.year = "20%s" % items[16]
            order_info_tmpl.cvv_number = items[17]

        elif items[0] == ITEM_SECTION:
            order_info = copy.deepcopy(order_info_tmpl)
            order_info.id = id
            order_info.item_name = items[4].replace(" ", "").replace("　", "")
            items[2] = items[2].replace(" ", "").replace("　", "")
            if items[2].upper() == SIZE_S:
                order_info.item_size = "Small"
            elif items[2].upper() == SIZE_M:
                order_info.item_size = "Medium"
            elif items[2].upper() == SIZE_L:
                order_info.item_size = "Large"
            elif items[2].upper() == SIZE_XL:
                order_info.item_size = "XLarge"
            else:
                order_info.item_size = items[2]

            if profile_no == PROFILE_NUM:
                order_info.profile_path = "%s/Default" % (PROFILE_DIR)
                order_info.profile_name = "Default"
            else:
                order_info.profile_path = "%s/Profile %d" % (PROFILE_DIR, profile_no)
                order_info.profile_name = "Profile %d" % profile_no

            if profile_no >= PROFILE_NUM:
                profile_no = 1
            else:
                profile_no += 1

            order_dict[id] = order_info
            id += 1

    return order_dict


def read_config():
    global SITE_URL
    global PROFILE_DIR
    global PROFILE_NUM
    for line in open(CONFIG_FILE, "r"):
        items = line.replace("\n", "").split("=")
        if items[0] == SITE_URL_KEY:
            SITE_URL = items[1]
        elif items[0] == PROFILE_DIR_KEY:
            PROFILE_DIR = items[1]
        elif items[0] == PROFILE_NUM_KEY:
            PROFILE_NUM = int(items[1])
        else:
            break


if __name__ == "__main__":
    read_config()
    app = wx.App()
    frm = MyFrame()
    frm.Show()
    app.MainLoop()
