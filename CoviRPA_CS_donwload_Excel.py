import os
import sys
import time
import bcrypt
import getpass
import win32com.client
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from dotenv import load_dotenv, set_key

# 시작
os.system('cls' if os.name == 'nt' else 'clear')

# 환경 변수 로드
ENV_FILE_PATH = r"C:\user_data\junho\.env"
load_dotenv(dotenv_path=ENV_FILE_PATH)
ENV_SITE_URL = os.getenv("SITE_URL")
ENV_USER_ID = os.getenv("USER_ID")
ENV_PASSWORD = os.getenv("PASSWORD")
ENV_DIRECTORY = os.getenv("EXCEL_DIR")

# Chrome WebDriver를 초기화하고 다운로드 경로 설정
def setup_driver(ENV_DIRECTORY):
    options = Options()
    options.add_argument('--start-maximized')
    options.add_argument("user-data-dir=C:\\user_data\\junho")
    options.add_experimental_option('detach', True)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    # 다운로드 디렉토리 설정
    prefs = {
        "download.default_directory": ENV_DIRECTORY,  # ENV_DIRECTORY 디렉토리 설정
        "download.prompt_for_download": False,    # 다운로드 대화 상자 비활성화
        "download.directory_upgrade": True,       # 기존 다운로드 경로 업데이트 허용
        "safebrowsing.enabled": True,             # 안전 브라우징 활성화
    }
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)  # WebDriver 생성

# 로그인 처리
def process_login():
    try:
        """
        작업순서 : 1.인증, 2.ID 확인 페이지 접속, 3.로그인 페이지 접속
        """
         # 1.인증
        input_password = getpass.getpass("Input password : ")
        hashed_password = bytes(ENV_PASSWORD.strip(), encoding='utf-8')
        time.sleep(1)

        if bcrypt.checkpw(input_password.encode(), hashed_password):
            print("Authentication successful")
        else:
            print("(!)Authentication failed")
            sys.exit()

        # 2.ID 확인 페이지 접속
        driver = setup_driver(ENV_DIRECTORY)  # 드라이버 초기화
        driver.get(ENV_SITE_URL)  # URL 로드
        try:
            composite_field = wait_for_element_presence(driver, By.ID, 'compositeAccount', timeout=3)
            composite_field.send_keys(ENV_USER_ID)

            next_button = wait_for_element_presence(driver, By.CLASS_NAME, 'btnLogin.btnCompositeNext', timeout=5)
            next_button.click()

        except Exception:

            # 3.로그인 페이지 접속
            id_field = wait_for_element_presence(driver, By.ID, 'id', timeout=10)
            id_field.send_keys(ENV_USER_ID)

            password_field = wait_for_element_presence(driver, By.ID, 'password', timeout=10)
            password_field.send_keys(input_password)

            login_button = wait_for_element_presence(driver, By.CLASS_NAME, 'btnLogin.btnInputLogin', timeout=10)
            login_button.click()

            time.sleep(5)  # 로그인 후 메인 페이지 팝업창 로딩 대기 시간
            print("Login successful")
            return driver

    except Exception as e:
        print("(!)Login process failed")
        sys.exit(1)  # 시스템 오류 종료

# 다운로드 디렉토리 초기화
def clear_download_folder():
    try:
        for file in os.listdir(ENV_DIRECTORY):
            file_path = os.path.join(ENV_DIRECTORY, file)
            os.remove(file_path)
        print("Directory initialization successful")

    except Exception as e:
        print("(!)Directory initialization failed")

# 지정된 요소를 대기하고 반환
def wait_for_element_presence(driver, by, locator, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, locator)))

# 클릭 가능한 요소를 대기하고 반환
def wait_for_element_clickable(driver, by, locator, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, locator)))

# 요소 대기 및 반환
# def wait_for_element(driver, by, locator, condition, timeout=30):
#     return WebDriverWait(driver, timeout).until(condition((by, locator)))

# 다운로드 파일 처리
def process_downloaded_file(file_keyword, task_name, row_del_count, col_del_count, max_wait_time=30):
    """
    작업순서 : 1.파일 확인, 2.파일명 변경, # 3.행/열 삭제 및 데이터 건수 계산, 4.xls 변환, 5.파일 사이즈 및 데이터 건수 리턴
    """
    try:
        elapsed_time = 0
        downloaded_file = None
        time.sleep(2)  # 다운로드 대기 시간

        # 1.파일 확인
        while elapsed_time < max_wait_time:
            files = [f for f in os.listdir(ENV_DIRECTORY) if file_keyword in f and f.lower().endswith(('.xlsx', '.xls'))]
            # files = [f for f in os.listdir(ENV_DIRECTORY) if file_keyword in f and not f.endswith('.crdownload')]
            if files:
                downloaded_file = files[0]  # 매칭된 첫 번째 파일
                break
            time.sleep(1)
            elapsed_time += 1

        if not downloaded_file:
            return f"← {file_keyword} ... (!)File does not exist"

        old_path = os.path.join(ENV_DIRECTORY, downloaded_file)
        new_path = os.path.join(ENV_DIRECTORY, task_name + os.path.splitext(downloaded_file)[1])

        # 2.파일명 변경
        if os.path.exists(new_path):
            os.remove(new_path)  # 동일 파일명 삭제
        os.rename(old_path, new_path)

        try:
            # 3.행/열 삭제 및 데이터 건수 계산
            excel = win32com.client.Dispatch("Excel.Application")  # pywin32로 엑셀 파일 열기
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(new_path)
            sheet = wb.Sheets(1)  # 첫 번째 시트 선택
            # time.sleep(2)  # 처리 대기 시간
            # data_rows = sheet.UsedRange.Rows.Count - (row_del_count + 1)  # 데이터 건수 계산(삭제 행, 헤더 행 제외)

            if row_del_count > 0:  # 행 삭제
                for i in range(row_del_count):
                    sheet.Rows(1).Delete()

            if col_del_count > 0:  # 열 삭제
                for i in range(col_del_count):
                    sheet.Columns(1).Delete()

            time.sleep(1)  # 처리 대기 시간
            data_rows = sheet.UsedRange.Rows.Count - 1  # 데이터 건수 계산(삭제 행, 헤더 행 제외)
            wb.SaveAs(new_path, FileFormat=51)  # xlsx 형식으로 저장
            # time.sleep(2)  # 처리 대기 시간
            wb.Close()

            # 4.xls 변환(.xls → .xlsx)
            if new_path.lower().endswith('.xls'):
                converted_file_path = new_path.replace('.xls', '.xlsx')

                wb = excel.Workbooks.Open(new_path)
                sheet = wb.Sheets(1)  # 첫 번째 시트 선택

                # if sheet.UsedRange.Rows.Count <= 1 and sheet.UsedRange.Columns.Count <= 1:
                #     wb.Close(SaveChanges=False)
                #     excel.Quit()
                #     return f"← {file_keyword} : (error) 데이터 없음"

                wb.SaveAs(converted_file_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook
                time.sleep(2)  # 처리 대기 시간
                wb.Close()

                # .xls 파일 삭제
                os.remove(new_path)
                new_path = converted_file_path  # 변환된 파일로 업데이트

        except Exception as e:
            excel.Quit()
            return f"← {downloaded_file} : (!)File conversion processing failed"

        time.sleep(1)  # 처리 대기 시간
        excel.Quit()

        # 5.파일 사이즈 및 데이터 건수 리턴
        file_size = os.path.getsize(new_path)  # 바이트 단위
        file_size_kb = round(file_size / 1024)  # KB 단위로 변환

        return "", f"{file_size_kb:,}KB", f"{data_rows:,}"

    except Exception as e: return f"- {file_keyword} : (!)File rename processing failed"

# 작업 처리
def process_task(task_name, file_keyword, row_del_count, col_del_count, search_button_xpath, page_url):
    driver.get(ENV_SITE_URL + page_url)  # URL 로드
    current_data = "0"
    try:
        if task_name == '6.근태조회':
            """
            작업순서 : 1.엑셀저장 버튼 클릭, 2.iframe으로 전환, 3.select 박스에서 옵션 선택, 4.체크박스 선택 (입출입기록 체크박스), 5.팝업 내 엑셀저장 버튼 클릭, 6.iframe에서 기본 콘텐츠로 복귀
            """
            # 1. 엑셀저장 버튼 클릭
            excel_button = wait_for_element_clickable(driver, By.XPATH, '//*[@id="excelBtn"]')
            excel_button.click()

            # 2.iframe으로 전환
            iframe = wait_for_element_presence(driver, By.ID, 'AttendUserStatusPopup_if')
            driver.switch_to.frame(iframe)

            # 3.select 박스에서 옵션 선택
            select_element = wait_for_element_presence(driver, By.ID, 'excelDownType')
            select = Select(select_element)
            select.select_by_value('M')  # value 속성값으로 선택 (월간)

            # 4.체크박스 선택 (입출입기록 체크박스)
            check_box = wait_for_element_clickable(driver, By.ID, 'dataMode3')
            if not check_box.is_selected():
                check_box.click()

            # 5.팝업 내 엑셀저장 버튼 클릭
            popup_excel_button = wait_for_element_clickable(driver, By.XPATH, '/html/body/div[1]/div/div[2]/a')
            popup_excel_button.click()

            # 6.iframe에서 기본 콘텐츠로 복귀
            driver.switch_to.default_content()

        elif task_name == '7.타임시트':
            """
            작업순서 : 1.달력 아이콘 클릭, 2.(1월) 클릭, 3.확인 버튼 클릭, 4.엑셀 저장 버튼 클릭
            """
            time.sleep(2) # 대기 시간

            # 1.달력 아이콘 클릭
            calendar_button = wait_for_element_clickable(driver, By.XPATH, '//*[@id="inputBasic_AX_EndDate_AX_dateHandle"]')
            calendar_button.click()

            # 2.(1월) 클릭
            select_start_month = wait_for_element_clickable(driver, By.XPATH, '//*[@id="inputBasic_AX_EndDate_AX_displayBox1_AX_1_AX_month"]')
            select_start_month.click()

            # 3.확인 버튼 클릭
            ok_button = wait_for_element_clickable(driver, By.XPATH, '//*[@id="inputBasic_AX_EndDate_AX_closeButton"]')
            ok_button.click()
            time.sleep(3) # 대기 시간

            # 4.엑셀 저장 버튼 클릭
            excel_button = wait_for_element_clickable(driver, By.XPATH, '//*[@id="divFilterBox"]/span[3]/button')
            driver.execute_script("arguments[0].click();", excel_button)  # JavaScript로 클릭

        else:  # 일반 작업
            """
            작업순서 : 1.시작일 초기화 및 설정, 2.검색 버튼 클릭, 3.엑셀저장 버튼 클릭, 4.엑셀 다운로드 팝업 확인 버튼 클릭
            """
            if task_name == '5.견적조회':  # 견적 tab 선택
                tab_element = wait_for_element_clickable(driver, By.XPATH, '//*[@id="content"]/div[2]/div/ul/li[7]/a')
                driver.execute_script("arguments[0].click();", tab_element)

            # 1.시작일 초기화 및 설정
            start_date_field = wait_for_element_presence(driver, By.ID, 'txtStartDt')
            start_date_field.clear()
            time.sleep(1)  # 처리 대기 시간
            start_date_field.send_keys('20080101')

            # 2.검색 버튼 클릭
            search_button = wait_for_element_clickable(driver, By.XPATH, search_button_xpath)
            search_button.click()

            if task_name == '1.수주조회': time.sleep(10) # 1.수주조회 (로딩 시간이 길어 15초 지연 대기)

            # 3.엑셀저장 버튼 클릭
            excel_button = wait_for_element_clickable(driver, By.CLASS_NAME, 'btnTypeDefault.btnExcel.acl-control-v')
            driver.execute_script("arguments[0].click();", excel_button)

            # 4.엑셀 다운로드 팝업 확인 버튼 클릭
            popup_ok = wait_for_element_presence(driver, By.ID, 'popup_ok')
            popup_ok.click()

        # 다운로드 파일 처리
        process_downloaded_file_result, file_size_kb, current_data = process_downloaded_file(file_keyword, task_name, row_del_count, col_del_count)

        # 환경 변수에 엑셀 데이터 건수 가져오기 및 저장
        previous_data = os.getenv(task_name, "0")  # 이전 엑셀 데이터 건수 가져오기
        set_key(ENV_FILE_PATH, task_name, current_data)  # 현재 엑셀 데이터 건수 저장

        if process_downloaded_file_result:  # 다운로드 파일 처리 오류 메시지가 반환된 경우
            return process_downloaded_file_result

        return f"{file_size_kb} ← {file_keyword} ({current_data} ← {previous_data})"  # 성공 메시지 반환

    except Exception as e:
        return "(!)Task execution failed"

# Main
if __name__ == "__main__":

    # 작업 목록
    task_list = [
        {
            "task_name": "1.수주조회", "file_keyword": "ordered", "row_del_count": 0, "col_del_count": 0,
            "search_button_xpath": '//*[@id="content"]/div[2]/div/div[1]/div/div[4]/a',
            "page_url": "bizmnt/layout/bizmnt_orderedList.do?CLSYS=bizmnt&CLMD=user&CLBIZ=BizMnt&menuID=334"
        },
        {
            "task_name": "2.인정실적", "file_keyword": "sales", "row_del_count": 2, "col_del_count": 1,
            "search_button_xpath": '//*[@id="content"]/div[2]/div/div[1]/div/div[2]/a',
            "page_url": "bizmnt/layout/bizmnt_salesPerformList.do?CLSYS=bizmnt&CLMD=user&CLBIZ=BizMnt&menuID=403"
        },
        {
            "task_name": "3.영업활동", "file_keyword": "active", "row_del_count": 2, "col_del_count": 1,
            "search_button_xpath": '//*[@id="btnSearch1"]',
            "page_url": "crm/layout/crm_active.do?CLSYS=crm&CLMD=user&CLBIZ=Crm&menuID=70975"
        },
        {
            "task_name": "4.영업기회", "file_keyword": "chance", "row_del_count": 2, "col_del_count": 1,
            "search_button_xpath": '//*[@id="btnSearch2"]',
            "page_url": "crm/layout/crm_chanceList.do?CLSYS=crm&CLMD=user&CLBIZ=Crm&menuID=70978"
        },
        {
            "task_name": "5.견적조회", "file_keyword": "Estimate", "row_del_count": 2, "col_del_count": 1,
            "search_button_xpath": '//*[@id="btnSearch2"]',
            "page_url": "crm/layout/crm_admin_basic_main.do?CLSYS=crm&CLMD=user&CLBIZ=Crm&menuID=71044&menuCode=BoardMain"
        },
        {
            "task_name": "6.근태조회", "file_keyword": "Attend", "row_del_count": 2, "col_del_count": 1,
            "search_button_xpath": '',
            "page_url": "groupware/layout/attend_AttendUserStatusList.do?CLSYS=attend&CLMD=user&CLBIZ=Attend&menuCode=EasyView"
        },
        {
            "task_name": "7.타임시트", "file_keyword": "프로젝트", "row_del_count": 0, "col_del_count": 0,
            "search_button_xpath": '',
            "page_url": "workreport/workreport/workreportteamproject.do?mnp=9%7C10"
        }
    ]

    # driver = setup_driver(ENV_DIRECTORY)  # 드라이버 초기화
    try:
        print("─" * 70)
        start_time = datetime.now()
        driver = process_login()  # 로그인 처리
        clear_download_folder()  # 다운로드 디렉토리 초기화
        print("─" * 70)

        for task in task_list:
            print(f"{task['task_name']}", end=" ")
            result = process_task(
                                    task["task_name"],
                                    task["file_keyword"],
                                    task["row_del_count"],
                                    task["col_del_count"],
                                    task["search_button_xpath"],
                                    task["page_url"]
                                )
            print(f"{result}")
        driver.quit()  # WebDriver 종료

    finally:
        print("─" * 70)
        end_time = datetime.now()
        elapsed_time = end_time - start_time
        print(f"[시작시간] {start_time.strftime('%H:%M:%S')}")
        print(f"[종료시간] {end_time.strftime('%H:%M:%S')}")
        print(f"[실행시간] {str(elapsed_time).split('.')[0]}")
