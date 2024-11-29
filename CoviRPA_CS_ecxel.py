from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from datetime import datetime
import time
import os

# [초기화] 드라이버 (브라우저, 다운로드 폴더)
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
    return webdriver.Chrome(options=options)


# [함수] 페이지 로드 대기
def wait_for_page_load(driver, task_name, timeout=20):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    # 페이지 로딩 추가 대기 시간    
    time.sleep(2)      


# [처리] 다운로드 폴더에서 파일명에 키워드가 포함된 파일을 찾아 규칙에 따라 일괄 변경
def rename_files_by_keywords(download_dir, rename_rules):
    try:
        for rule in rename_rules:
            keyword = rule["keyword"]
            task_name = rule["task_name"]

            # 다운로드 폴더에서 파일 검색
            files = [f for f in os.listdir(download_dir) if keyword in f and f.endswith(('.xlsx', '.xls'))]
            if not files:
                print(f"\n[{task_name}] '{keyword}' 키워드를 포함한 파일을 찾을 수 없습니다.")
                continue

            # 가장 최근 파일 선택 (키워드와 매칭되는 파일 중 생성 시점이 가장 늦은 파일)
            latest_file = max([os.path.join(download_dir, f) for f in files], key=os.path.getctime)

            # 새 파일 이름 생성
            file_extension = os.path.splitext(latest_file)[1]
            new_file_path = os.path.join(download_dir, f"{task_name}{file_extension}")

            # 기존 동일 이름 파일이 있으면 삭제
            if os.path.exists(new_file_path): os.remove(new_file_path)

            # 파일 이름 변경
            os.rename(latest_file, new_file_path)
    except Exception as e: print("Error")

# [처리] 로그인
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
        print(f"로그인 중 오류 발생: {e}")
        driver.quit()


# [처리] 작업 실행
def task_action(driver, task_name, url, search_button_xpath, data_xpath):
    try:
        # 페이지 이동 및 로드 대기
        print(f"{task_name} :", end=" ")
        driver.get(url)
        wait_for_page_load(driver, task_name)
 
        if task_name == '6.근태조회': # 6.근태조회 (예외 작업)

            # 엑셀저장 버튼 클릭
            excel_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="excelBtn"]'))
            )
            excel_button.click()

            # iframe으로 전환
            iframe = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'AttendUserStatusPopup_if'))  # iframe의 ID 또는 다른 적절한 locator
            )
            driver.switch_to.frame(iframe)  # iframe으로 전환

            # select 박스에서 옵션 선택
            select_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'excelDownType'))  # select 박스의 ID
            )
            select = Select(select_element)  # Select 객체 생성
            select.select_by_value('M')  # value 속성값으로 선택 (월간)

            # 체크박스 선택 (입출입기록 체크박스)
            check_box = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'dataMode3'))  # 체크박스 ID
            )
            if not check_box.is_selected():  # 이미 선택된 상태가 아니면 클릭
                check_box.click()

            # 팝업 내 엑셀저장 버튼 클릭
            popup_excel_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/a'))
            )
            popup_excel_button.click()

            # iframe에서 기본 콘텐츠로 복귀
            driver.switch_to.default_content()

            # 작업 완료 출력
            print("OK")

        elif task_name == '7.타임시트': # 7.타임시트 (예외 작업)

            # 달력 선택
            calendar_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="inputBasic_AX_EndDate_AX_dateHandle"]'))
            )
            calendar_button.click()

            # 1월 선택
            select_start_month = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="inputBasic_AX_EndDate_AX_displayBox1_AX_1_AX_month"]'))
            )
            select_start_month.click()

            # OK 버튼 클릭
            ok_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="inputBasic_AX_EndDate_AX_closeButton"]'))
            )
            ok_button.click()
            
            # 데이터 로드 대기 시간
            time.sleep(3)

            # 엑셀 저장 버튼 클릭
            excel_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="divFilterBox"]/span[3]/button'))
            )
            driver.execute_script("arguments[0].click();", excel_button)  # JavaScript로 클릭

            # 작업 완료 출력
            print("OK")

        else: # 1,2,3,4,5.작업

            # 5.견적조회 (탭 클릭)
            if task_name == '5.견적조회':
                tab_element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="content"]/div[2]/div/ul/li[7]/a'))
                )
                driver.execute_script("arguments[0].click();", tab_element)

            # 시작일 초기화 및 설정
            start_date_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'txtStartDt'))
            )
            start_date_field.clear()
            start_date_field.send_keys('2008.01.01')

            # 검색 버튼 클릭
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, search_button_xpath))
            )
            search_button.click()

            if task_name == '1.수주조회': time.sleep(15) # 1.수주조회 (로딩 시간이 길어 15초 지연 대기)

            # 엑셀저장 버튼 클릭
            excel_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, 'btnTypeDefault.btnExcel.acl-control-v'))
            )
            driver.execute_script("arguments[0].click();", excel_button)

            # 엑셀 다운로드 팝업 확인 버튼 클릭
            popup_ok = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'popup_ok'))
            )
            popup_ok.click()
            
            # 조회 결과 데이터 추출
            data_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, data_xpath))
            )
            extracted_data = data_element.text
            print(f"{extracted_data}")

        # 파일 다운로드 대기 시간
        time.sleep(3)
        return None

    except Exception as e:
        print("ERROR")
        return None


# [처리] MAIN
if __name__ == "__main__":
    # 시작 작업
    start_time = datetime.now()
    print(f"\n\n\n#시작시간 : {start_time.strftime('%H:%M:%S')}")

    # 정보 설정
    USER_ID = 'jhlee6'
    PASSWORD = 'ewq3412!4520'

    # 다운로드 폴더 설정
    download_dir = r"C:\Users\jhlee6\OneDrive - Covision\21 CSreport\Excel"

    # 드라이버 초기화 (브라우저, 다운로드 폴더)
    driver = setup_driver(download_dir)

    # 작업 대상 설정 (작업이름, URL, 검색버튼 XPATH, 데이터 XPATH)
    download_tasks = [
        {
            "task_name": "1.수주조회",
            "url": "https://gw4j.covision.co.kr/bizmnt/layout/bizmnt_orderedList.do?CLSYS=bizmnt&CLMD=user&CLBIZ=BizMnt&menuID=334",
            "search_button_xpath": '//*[@id="content"]/div[2]/div/div[1]/div/div[4]/a',
            "data_xpath": '//*[@id="listGrid_AX_gridStatus"]/b',
            "keyword": "ordered"
        },
        {
            "task_name": "2.인정실적",
            "url": "https://gw4j.covision.co.kr/bizmnt/layout/bizmnt_salesPerformList.do?CLSYS=bizmnt&CLMD=user&CLBIZ=BizMnt&menuID=403",
            "search_button_xpath": '//*[@id="content"]/div[2]/div/div[1]/div/div[2]/a',
            "data_xpath": '//*[@id="spnSumRecognition"]',
            "keyword": "salesPerform"
        },
        {
            "task_name": "3.영업활동",
            "url": "https://gw4j.covision.co.kr/crm/layout/crm_active.do?CLSYS=crm&CLMD=user&CLBIZ=Crm&menuID=70975",
            "search_button_xpath": '//*[@id="btnSearch1"]',
            "data_xpath": '//*[@id="listGrid_AX_gridStatus"]/b',
            "keyword": "active"
        },
        {
            "task_name": "4.영업기회",
            "url": "https://gw4j.covision.co.kr/crm/layout/crm_chanceList.do?CLSYS=crm&CLMD=user&CLBIZ=Crm&menuID=70978",
            "search_button_xpath": '//*[@id="btnSearch2"]',
            "data_xpath": '//*[@id="listGrid_AX_gridStatus"]/b',
            "keyword": "chance"
        },
        {
            "task_name": "5.견적조회",
            "url": "https://gw4j.covision.co.kr/crm/layout/crm_admin_basic_main.do?CLSYS=crm&CLMD=user&CLBIZ=Crm&menuID=71044&menuCode=BoardMain",
            "search_button_xpath": '//*[@id="btnSearch2"]',
            "data_xpath": '//*[@id="listGrid_AX_gridStatus"]/b',
            "keyword": "basicEstimate"
        },
        {
            "task_name": "6.근태조회",
            "url": "https://gw4j.covision.co.kr/groupware/layout/attend_AttendUserStatusList.do?CLSYS=attend&CLMD=user&CLBIZ=Attend&menuCode=EasyView",
            "search_button_xpath": '',
            "data_xpath": '',
            "keyword": "AttendUserStatusMonthInfo"
        },
        {
            "task_name": "7.타임시트",
            "url": "https://gw4j.covision.co.kr/workreport/workreport/workreportteamproject.do?mnp=9%7C10",
            "search_button_xpath": '',
            "data_xpath": '',
            "keyword": "팀원프로젝트분석"
        }        
    ]

    try:
        # 로그인 및 출력
        print("\nLOGIN .......................", end=" ")
        login(driver, USER_ID, PASSWORD)
        time.sleep(5)  # 로그인 후 메인 페이지 팝업창 로딩 대기 시간
        print("OK\n")

        # 작업 지시
        for task in download_tasks:
            task_action(
                driver,
                task["task_name"],
                task["url"],
                task["search_button_xpath"],
                task["data_xpath"]
            )

        # 다운로드 파일 일괄 파일명 변경
        print("\nFILE RENAME .................", end=" ")
        # time.sleep(15)  # 마지막 파일 다운로드 지연시간
        rename_files_by_keywords(download_dir, download_tasks)
        print("OK")

    finally:
        # 마무리 작업
        print("CLOSE .......................", end=" ")
        driver.quit()  #브라우저 종료
        print("OK")

        # 종료 작업
        end_time = datetime.now()
        print(f"\n#종료시간: {end_time.strftime('%H:%M:%S')}")
        elapsed_time = end_time - start_time
        print(f"#실행시간: {str(elapsed_time).split('.')[0]}\n\n\n")