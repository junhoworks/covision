import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select


# 다운로드 폴더 설정
download_dir = r"C:\Users\jhlee6\OneDrive - Covision\21 CSreport\Excel"


# Chrome WebDriver를 초기화하고 다운로드 경로 설정
def setup_driver(download_dir):
    options = Options()
    options.add_argument('--start-maximized')
    options.add_experimental_option('detach', True)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    # 다운로드 폴더 설정
    prefs = {
        "download.default_directory": download_dir,  # 다운로드 폴더 경로
        "download.prompt_for_download": False,       # 다운로드 대화 상자 비활성화
        "download.directory_upgrade": True,          # 다운로드 경로 업그레이드 허용
        "safebrowsing.enabled": True                 # 안전 브라우징 기능 활성화
    }
    options.add_experimental_option("prefs", prefs)

    # WebDriver 생성
    return webdriver.Chrome(options=options)


# 다운로드 폴더 초기화
def clear_download_folder(download_dir):  
    for file in os.listdir(download_dir):
        file_path = os.path.join(download_dir, file)
        os.remove(file_path)


# 지정된 요소를 대기하고 반환
def wait_for_element_presence(driver, by, locator, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, locator)))


# 클릭 가능한 요소를 대기하고 반환
def wait_for_element_clickable(driver, by, locator, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, locator)))


# 다운로드 파일 확인 및 파일명 변경 (1)
# def wait_for_download_and_rename(download_dir, file_keyword, task_name, max_wait_time=30):
#     try:
#         elapsed_time = 0
#         downloaded_file = None

#         while elapsed_time < max_wait_time:
#             files = [f for f in os.listdir(download_dir) if file_keyword in f and f.lower().endswith(('.xlsx', '.xls'))]
#             # files = [f for f in os.listdir(download_dir) if file_keyword in f and not f.endswith('.crdownload')] # 문제 있음.
          
#             if files:
#                 downloaded_file = files[0]
#                 break
#             time.sleep(1)
#             elapsed_time += 1

#         if not downloaded_file: raise FileNotFoundError(f"- {file_keyword} : (!)파일 없음")

#         old_file_path = os.path.join(download_dir, downloaded_file)
#         new_file_path = os.path.join(download_dir, task_name + os.path.splitext(downloaded_file)[1])

#         if os.path.exists(new_file_path): os.remove(new_file_path)
#         os.rename(old_file_path, new_file_path)

#     except Exception as e: raise Exception(f"- {file_keyword} : (!)파일명 변경 오류")


# 다운로드 파일 확인 및 파일명 변경 (2)
def wait_for_download_and_rename(download_dir, file_keyword, task_name, max_wait_time=30):
    try:
        # 다운로드 파일 확인
        max_wait_time = 30  # 최대 대기 시간 (초)
        elapsed_time = 0
        downloaded_file = None

        while elapsed_time < max_wait_time:
            # 다운로드 폴더에서 파일 검색
            files = [f for f in os.listdir(download_dir) if file_keyword in f and f.lower().endswith(('.xlsx', '.xls'))]
            # files = [f for f in os.listdir(download_dir) if file_keyword in f and not f.endswith('.crdownload')] # 문제 있음.
            if files:
                downloaded_file = files[0]  # 매칭된 첫 번째 파일
                break

            time.sleep(1)  # 1초 대기
            elapsed_time += 1

        if not downloaded_file:
            return f"- {file_keyword} : 실패"

        # 파일명 변경
        old_file_path = os.path.join(download_dir, downloaded_file)
        old_file_extension = os.path.splitext(old_file_path)[1]
        new_file_path = os.path.join(download_dir, f"{task_name}{old_file_extension}")

        if os.path.exists(new_file_path):  # 동일한 이름의 파일이 이미 존재하면 삭제
            os.remove(new_file_path)

        os.rename(old_file_path, new_file_path)

    except Exception as e: raise Exception(f"- {file_keyword} : (!)파일명 변경 오류")    


# 로그인
def login(driver, user_id, password):
    driver.get('https://gw4j.covision.co.kr')
    try:
        # 사용자 ID 입력 (ID 1차 확인 페이지가 있을 경우) - 불필요 시 주석 처리 할것
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'compositeAccount'))
        ).send_keys(user_id)

        # 다음 버튼 클릭 (ID 1차 확인 페이지가 있을 경우) - 불필요 시 주석 처리 할것
        driver.find_element(By.CLASS_NAME, 'btnLogin.btnCompositeNext').click()

        # # 사용자 ID 입력 (ID 1차 확인 페이지가 없을 경우) - 불필요 시 주석 처리 할것
        # WebDriverWait(driver, 10).until(
        #     EC.presence_of_element_located((By.ID, 'id'))
        # ).send_keys(user_id)

        # 비밀번호 입력
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'password'))
        ).send_keys(password)

        # 로그인 버튼 클릭
        driver.find_element(By.CLASS_NAME, 'btnLogin.btnInputLogin').click()
    except Exception as e:
        print(f"(!)로그인 오류")
        driver.quit()


# 작업 실행
def tasks_action(driver, task_name, file_keyword, search_button_xpath, data_xpath, page_url):
    driver.get(page_url)  # 페이지 로드
    extracted_data = 0
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
            extracted_data = data_element.text    

        # 다운로드 파일 확인 및 파일명 변경
        wait_for_download_and_rename(download_dir, file_keyword, task_name)
        return f"- {file_keyword} : 완료 ({extracted_data})"
        
    except Exception as e: return "(!)작업 실행 오류"


# Main
if __name__ == "__main__":
    
    start_time = datetime.now()
    print(f"\n\n\ndownload_folder initialize")
    clear_download_folder(download_dir)  # 다운로드 폴더 초기화
    driver = setup_driver(download_dir)  # 드라이버 초기화 (브라우저, 다운로드 폴더)
    print("-------------------------------------------")

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

        # 로그인
        print("0.LOGIN :", end=" ")
        login(driver, 'jhlee6', 'ewq3412!4520')        
        time.sleep(5)  # 로그인 후 메인 페이지 팝업창 로딩 대기 시간
        print("완료")

        # 작업 지시
        for task in task_list:
            print(f"{task['task_name']}", end=" ")
            result = tasks_action(driver, task["task_name"], task["file_keyword"], task["search_button_xpath"], task["data_xpath"], task["page_url"])
            print(f"{result}")

    finally:
        # WebDriver 종료
        driver.quit()

        # 시간 출력 (시작, 종료, 실행)
        print("-------------------------------------------")
        end_time = datetime.now()
        elapsed_time = end_time - start_time
        print(f"[시작시간] {start_time.strftime('%H:%M:%S')}")
        print(f"[종료시간] {end_time.strftime('%H:%M:%S')}")
        print(f"[실행시간] {str(elapsed_time).split('.')[0]}\n\n\n")