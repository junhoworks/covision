import os
import sys
import time
import threading
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from dotenv import load_dotenv, set_key
from tkinter import Tk, Text, Scrollbar, END, VERTICAL, RIGHT, Y


# GUI 로그 출력 창
class LogWindow:
    def __init__(self):
        self.root = Tk()
        self.root.title("Task Automation Logs")
        self.root.geometry("800x600")

        self.text = Text(self.root, wrap="word", font=("Consolas", 10))
        self.text.pack(side=RIGHT, fill="both", expand=True)

        self.scrollbar = Scrollbar(self.text, orient=VERTICAL, command=self.text.yview)
        self.text.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side=RIGHT, fill=Y)

    def write(self, message):
        def _write():
            self.text.insert(END, message + "\n")
            self.text.see(END)  # Auto-scroll

        self.root.after(0, _write)

    def start(self):
        self.root.mainloop()


log_window = LogWindow()


def print(*args, **kwargs):
    message = " ".join(str(arg) for arg in args)
    log_window.write(message)


# 환경 변수 로드
ENV_FILE_PATH = r"C:\user_data\junho\.env"
load_dotenv(dotenv_path=ENV_FILE_PATH)
USER_ID = os.getenv("USER_ID")
PASSWORD = os.getenv("PASSWORD")
EXCEL_DIR = os.getenv("EXCEL_DIR")


# WebDriver 초기화
def setup_driver(EXCEL_DIR):
    options = Options()
    options.add_argument("--headless")  # (선택 사항) 헤드리스 모드로 실행
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("start-maximized")
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument("user-data-dir=C:\\user_data\\junho")

    # options.add_argument('--start-maximized')
    # options.add_argument("user-data-dir=C:\\user_data\\junho")
    # options.add_experimental_option('detach', True)
    # options.add_experimental_option('excludeSwitches', ['enable-logging'])

    prefs = {
        "download.default_directory": EXCEL_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)

    return webdriver.Chrome(options=options)


def clear_download_folder(EXCEL_DIR):
    print(f"Excel 디렉토리 초기화 ...", end=" ")
    for file in os.listdir(EXCEL_DIR):
        file_path = os.path.join(EXCEL_DIR, file)
        try:
            os.remove(file_path)
        except Exception as e:
            print(f"파일 삭제 실패: {file_path} - {e}")
    print("ok")


def wait_for_element_presence(driver, by, locator, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, locator)))


def wait_for_element_clickable(driver, by, locator, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, locator)))


def wait_for_download_and_rename(EXCEL_DIR, file_keyword, task_name, max_wait_time=30):
    elapsed_time = 0
    downloaded_file = None
    while elapsed_time < max_wait_time:
        files = [f for f in os.listdir(EXCEL_DIR) if file_keyword in f and f.lower().endswith(('.xlsx', '.xls'))]
        if files:
            downloaded_file = files[0]
            break
        time.sleep(1)
        elapsed_time += 1

    if not downloaded_file:
        return f"← {file_keyword} ... (error) 파일 없음"

    old_file_path = os.path.join(EXCEL_DIR, downloaded_file)
    new_file_path = os.path.join(EXCEL_DIR, task_name + os.path.splitext(downloaded_file)[1])

    if os.path.exists(new_file_path):
        os.remove(new_file_path)

    os.rename(old_file_path, new_file_path)

    file_size = os.path.getsize(new_file_path)
    file_size_kb = round(file_size / 1024)
    return "", f"{file_size_kb:,}KB"


def login(driver, user_id, password):
    print("LOGIN ...", end=" ")
    try:
        driver.get('https://gw4j.covision.co.kr')
        id_field = wait_for_element_presence(driver, By.ID, 'id', timeout=10)
        id_field.send_keys(user_id)
        password_field = wait_for_element_presence(driver, By.ID, 'password', timeout=10)
        password_field.send_keys(password)
        login_button = wait_for_element_clickable(driver, By.CLASS_NAME, 'btnLogin.btnInputLogin', timeout=10)
        login_button.click()
        time.sleep(5)
        print("ok")
        return True
    except Exception as e:
        print(f"(error) 로그인 실패: {e}")
        sys.exit(1)


def tasks_action(driver, task_name, file_keyword, search_button_xpath, data_xpath, page_url):
    driver.get(page_url)
    current_data = "0"
    try:
        if search_button_xpath:
            search_button = wait_for_element_clickable(driver, By.XPATH, search_button_xpath)
            search_button.click()

        if data_xpath:
            data_element = wait_for_element_presence(driver, By.XPATH, data_xpath)
            current_data = data_element.text

        previous_data = os.getenv(task_name, "0")
        set_key(ENV_FILE_PATH, task_name, current_data)
        rename_result, file_size_kb = wait_for_download_and_rename(EXCEL_DIR, file_keyword, task_name)

        if rename_result:
            return rename_result
        return f"{file_size_kb} ← {file_keyword} ({current_data} ← {previous_data})"
    except Exception as e:
        return f"(error) 작업 실행 실패: {e}"


def run_task_in_thread(task):
    driver = setup_driver(EXCEL_DIR)
    login(driver, USER_ID, PASSWORD)
    result = tasks_action(driver, task["task_name"], task["file_keyword"], task["search_button_xpath"], task["data_xpath"], task["page_url"])
    driver.quit()
    print(result)


if __name__ == "__main__":
    print("\n\n\n" + "─" * 60)
    
    # 작업 목록
    task_list = [        
        {
            "task_name": "1.수주조회", "file_keyword": "ordered",
            "search_button_xpath": '//*[@id="content"]/div[2]/div/div[1]/div/div[4]/a',
            "data_xpath": '//*[@id="listGrid_AX_gridStatus"]/b',
            "page_url": "https://gw4j.covision.co.kr/bizmnt/layout/bizmnt_orderedList.do?CLSYS=bizmnt&CLMD=user&CLBIZ=BizMnt&menuID=334"
        },
        {
            "task_name": "2.인정실적", "file_keyword": "sales",
            "search_button_xpath": '//*[@id="content"]/div[2]/div/div[1]/div/div[2]/a',
            "data_xpath": '//*[@id="spnSumRecognition"]',            
            "page_url": "https://gw4j.covision.co.kr/bizmnt/layout/bizmnt_salesPerformList.do?CLSYS=bizmnt&CLMD=user&CLBIZ=BizMnt&menuID=403"
        },
        {
            "task_name": "3.영업활동", "file_keyword": "active",
            "search_button_xpath": '//*[@id="btnSearch1"]',
            "data_xpath": '//*[@id="listGrid_AX_gridStatus"]/b',            
            "page_url": "https://gw4j.covision.co.kr/crm/layout/crm_active.do?CLSYS=crm&CLMD=user&CLBIZ=Crm&menuID=70975"
        },
        {
            "task_name": "4.영업기회", "file_keyword": "chance",
            "search_button_xpath": '//*[@id="btnSearch2"]',
            "data_xpath": '//*[@id="listGrid_AX_gridStatus"]/b',
            "page_url": "https://gw4j.covision.co.kr/crm/layout/crm_chanceList.do?CLSYS=crm&CLMD=user&CLBIZ=Crm&menuID=70978"
        },
        {
            "task_name": "5.견적조회", "file_keyword": "Estimate",
            "search_button_xpath": '//*[@id="btnSearch2"]',
            "data_xpath": '//*[@id="listGrid_AX_gridStatus"]/b',
            "page_url": "https://gw4j.covision.co.kr/crm/layout/crm_admin_basic_main.do?CLSYS=crm&CLMD=user&CLBIZ=Crm&menuID=71044&menuCode=BoardMain"
        },
        {
            "task_name": "6.근태조회", "file_keyword": "Attend",
            "search_button_xpath": '',
            "data_xpath": '',
            "page_url": "https://gw4j.covision.co.kr/groupware/layout/attend_AttendUserStatusList.do?CLSYS=attend&CLMD=user&CLBIZ=Attend&menuCode=EasyView"
        },
        {
            "task_name": "7.타임시트", "file_keyword": "프로젝트",
            "search_button_xpath": '',
            "data_xpath": '',
            "page_url": "https://gw4j.covision.co.kr/workreport/workreport/workreportteamproject.do?mnp=9%7C10"
        }        
    ]
    try:
        for task in task_list:
            threading.Thread(target=run_task_in_thread, args=(task,), daemon=True).start()
            time.sleep(1)
    except Exception as e:
        print(f"(error) 프로그램 실행 중 오류: {e}")
    finally:
        log_window.start()
