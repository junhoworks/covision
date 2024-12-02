import os
import sys
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from dotenv import load_dotenv, set_key


# 환경 변수 로드
ENV_FILE_PATH = r"C:\user_data\junho\.env"
load_dotenv(dotenv_path=ENV_FILE_PATH)
USER_ID = os.getenv("USER_ID")
PASSWORD = os.getenv("PASSWORD")
EXCEL_DIR = os.getenv("EXCEL_DIR")  # EXCEL_DIR 디렉토리 설정


# Chrome WebDriver를 초기화하고 다운로드 경로 설정
def setup_driver(EXCEL_DIR):
    options = Options()
    options.add_argument('--start-maximized')
    options.add_argument("user-data-dir=C:\\user_data\\junho")
    options.add_experimental_option('detach', True)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    # 다운로드 디렉토리 설정
    prefs = {
        "download.default_directory": EXCEL_DIR,  # EXCEL_DIR 디렉토리 설정
        "download.prompt_for_download": False,    # 다운로드 대화 상자 비활성화
        "download.directory_upgrade": True,       # 기존 다운로드 경로 업데이트 허용
        "safebrowsing.enabled": True,             # 안전 브라우징 활성화
    }
    options.add_experimental_option("prefs", prefs)

    # WebDriver 생성
    return webdriver.Chrome(options=options)


# 다운로드 디렉토리 초기화
def clear_download_folder(EXCEL_DIR):
    print(f"Excel 디렉토리 초기화 ...", end = " ")
    for file in os.listdir(EXCEL_DIR):
        file_path = os.path.join(EXCEL_DIR, file)
        os.remove(file_path)
    print("ok")


# 지정된 요소를 대기하고 반환
def wait_for_element_presence(driver, by, locator, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, locator)))


# 클릭 가능한 요소를 대기하고 반환
def wait_for_element_clickable(driver, by, locator, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, locator)))


# 다운로드 파일 확인 및 파일명 변경
def wait_for_download_and_rename(EXCEL_DIR, file_keyword, task_name, max_wait_time=30):
    try:
        elapsed_time = 0
        downloaded_file = None

        # 다운로드 파일명 확인
        while elapsed_time < max_wait_time:
            files = [f for f in os.listdir(EXCEL_DIR) if file_keyword in f and f.lower().endswith(('.xlsx', '.xls'))]
            # files = [f for f in os.listdir(EXCEL_DIR) if file_keyword in f and not f.endswith('.crdownload')]

            if files:
                downloaded_file = files[0]  # 매칭된 첫 번째 파일
                break
            time.sleep(1)  # 1초 대기
            elapsed_time += 1

        if not downloaded_file:
            return f"← {file_keyword} ... (error) 파일 없음"
        
        old_file_path = os.path.join(EXCEL_DIR, downloaded_file)
        new_file_path = os.path.join(EXCEL_DIR, task_name + os.path.splitext(downloaded_file)[1])

        # 파일명 변경
        if os.path.exists(new_file_path):
            os.remove(new_file_path)  # 동일 파일명 삭제

        os.rename(old_file_path, new_file_path)

        # 파일 사이즈 리턴
        file_size = os.path.getsize(new_file_path)  # 바이트 단위
        file_size_kb = round(file_size / 1024)  # KB 단위로 변환

        return "", f"{file_size_kb:,}KB"
    
    except Exception as e: return f"- {file_keyword} : (error) 파일명 변경 실패"


# 로그인
def login(driver, user_id, password):
    print("LOGIN ...", end=" ")
    try:
        driver.get('https://gw4j.covision.co.kr')

        # ID 확인 페이지
        try:
            composite_field = wait_for_element_presence(driver, By.ID, 'compositeAccount', timeout=3)
            composite_field.send_keys(user_id)

            next_button = wait_for_element_presence(driver, By.CLASS_NAME, 'btnLogin.btnCompositeNext', timeout=5)
            next_button.click()

        except Exception:

            # 로그인 페이지     
            id_field = wait_for_element_presence(driver, By.ID, 'id', timeout=10)
            id_field.send_keys(user_id)

            password_field = wait_for_element_presence(driver, By.ID, 'password', timeout=10)
            password_field.send_keys(password)

            login_button = wait_for_element_presence(driver, By.CLASS_NAME, 'btnLogin.btnInputLogin', timeout=10)
            login_button.click()

            time.sleep(5)  # 로그인 후 메인 페이지 팝업창 로딩 대기 시간
            print("ok")
            return True

    except Exception as e:
        print("(error) 로그인 실패")
        sys.exit(1)  # 시스템 오류 종료



# 작업 실행
def tasks_action(driver, task_name, file_keyword, search_button_xpath, data_xpath, page_url):
    driver.get(page_url)  # 페이지 로드
    current_data = "0"
    try:
        if task_name == '6.근태조회':  # 6.근태조회 (예외 작업)
            # 엑셀저장 버튼 클릭
            excel_button = wait_for_element_clickable(driver, By.XPATH, '//*[@id="excelBtn"]')
            excel_button.click()

            # iframe으로 전환
            iframe = wait_for_element_presence(driver, By.ID, 'AttendUserStatusPopup_if')
            driver.switch_to.frame(iframe)

            # select 박스에서 옵션 선택
            select_element = wait_for_element_presence(driver, By.ID, 'excelDownType')
            select = Select(select_element)
            select.select_by_value('M')  # value 속성값으로 선택 (월간)

            # 체크박스 선택 (입출입기록 체크박스)
            check_box = wait_for_element_clickable(driver, By.ID, 'dataMode3')
            if not check_box.is_selected():
                check_box.click()

            # 팝업 내 엑셀저장 버튼 클릭
            popup_excel_button = wait_for_element_clickable(driver, By.XPATH, '/html/body/div[1]/div/div[2]/a')
            popup_excel_button.click()

            # iframe에서 기본 콘텐츠로 복귀
            driver.switch_to.default_content()

        elif task_name == '7.타임시트':  # 7.타임시트 (예외 작업)

            # 1.달력 아이콘 클릭
            calendar_button = wait_for_element_clickable(driver, By.XPATH, '//*[@id="inputBasic_AX_EndDate_AX_dateHandle"]')
            calendar_button.click()

            # 2.(1월) 클릭
            select_start_month = wait_for_element_clickable(driver, By.XPATH, '//*[@id="inputBasic_AX_EndDate_AX_displayBox1_AX_1_AX_month"]')
            select_start_month.click()

            # 3.확인 버튼 클릭
            ok_button = wait_for_element_clickable(driver, By.XPATH, '//*[@id="inputBasic_AX_EndDate_AX_closeButton"]')
            ok_button.click()
            
            time.sleep(5) # 데이터 로드 대기 시간

            # 4. 엑셀 저장 버튼 클릭
            excel_button = wait_for_element_clickable(driver, By.XPATH, '//*[@id="divFilterBox"]/span[3]/button')
            driver.execute_script("arguments[0].click();", excel_button)  # JavaScript로 클릭

            time.sleep(2) # 데이터 로드 대기 시간

        else:  # 1,2,3,4,5.작업
            if task_name == '5.견적조회':
                tab_element = wait_for_element_clickable(driver, By.XPATH, '//*[@id="content"]/div[2]/div/ul/li[7]/a')
                driver.execute_script("arguments[0].click();", tab_element)

            # 1.시작일 초기화 및 설정
            start_date_field = wait_for_element_presence(driver, By.ID, 'txtStartDt')
            start_date_field.clear()
            start_date_field.send_keys('2008.01.01')

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

            # 5.조회 건수 데이터 추출
            data_element = wait_for_element_presence(driver, By.XPATH, data_xpath)
            current_data = data_element.text    

        previous_data = os.getenv(task_name, "0")  # 이전 데이터 가져오기
        set_key(ENV_FILE_PATH, task_name, current_data)  # 현재 데이터 저장

        # 다운로드 파일 확인 및 파일명 변경
        rename_result, file_size_kb = wait_for_download_and_rename(EXCEL_DIR, file_keyword, task_name)

        if rename_result:  # 오류 메시지가 반환된 경우
            return rename_result  # 오류 메시지 반환        

        return f"{file_size_kb} ← {file_keyword} ({current_data} ← {previous_data})"  # 성공 메시지 반환
        
    except Exception as e:
        print("(error) 작업 실행 실패")
        sys.exit(1)  # 시스템 오류 종료


# Main
if __name__ == "__main__":    

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
        print("\n\n\n" + "─" * 60)
        start_time = datetime.now()
        driver = setup_driver(EXCEL_DIR)  # 드라이버 초기화 (브라우저, 다운로드 디렉토리)
        login(driver, USER_ID, PASSWORD)  # 로그인
        clear_download_folder(EXCEL_DIR)  # 다운로드 디렉토리 초기화

        # 작업 지시
        print("─" * 60)        
        for task in task_list:
            print(f"{task['task_name']}", end=" ")
            result = tasks_action(driver, task["task_name"], task["file_keyword"], task["search_button_xpath"], task["data_xpath"], task["page_url"])
            print(f"{result}")

    finally:
        print("─" * 60)
        end_time = datetime.now()
        elapsed_time = end_time - start_time
        print(f"[시작시간] {start_time.strftime('%H:%M:%S')}")
        print(f"[종료시간] {end_time.strftime('%H:%M:%S')}")
        print(f"[실행시간] {str(elapsed_time).split('.')[0]}\n\n\n")
        driver.quit()  # WebDriver 종료
        sys.exit(0)  # 시스템 정상 종료