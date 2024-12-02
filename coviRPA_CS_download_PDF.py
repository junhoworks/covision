import os
import sys
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv, set_key


# 환경 변수 로드
ENV_FILE_PATH = r"C:\user_data\junho\.env"
load_dotenv(dotenv_path=ENV_FILE_PATH)
PDF_DIR = os.getenv("PDF_DIR")  # PDF_DIR 디렉토리 설정


# Chrome WebDriver를 초기화하고 다운로드 디렉토리 설정
def setup_driver(PDF_DIR):
    options = Options()
    options.add_argument('--start-maximized')
    options.add_argument("user-data-dir=C:\\user_data\\junho")
    options.add_experimental_option('detach', True)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    # 다운로드 디렉토리 설정
    prefs = {
        "download.default_directory": PDF_DIR,  # PDF_DIR 디렉토리 설정
        "download.prompt_for_download": False,  # 다운로드 대화 상자 비활성화
        "download.directory_upgrade": True,     # 기존 다운로드 경로 업데이트 허용
        "safebrowsing.enabled": True,           # 안전 브라우징 활성화
    }
    options.add_experimental_option("prefs", prefs)

    # WebDriver 생성
    return webdriver.Chrome(options=options)


# 다운로드 디렉토리 초기화
def clear_download_folder(PDF_DIR):
    print(f"PDF 디렉토리 초기화 ...", end = " ") 
    for file in os.listdir(PDF_DIR):
        file_path = os.path.join(PDF_DIR, file)
        os.remove(file_path)
    print("ok")


# 지정된 XPATH를 가진 요소가 DOM에 존재할 때까지 대기하고 반환
def wait_for_element_presence(driver, xpath, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))


# 지정된 XPATH를 가진 요소가 클릭 가능할 때까지 대기하고 반환
def wait_for_element_clickable(driver, xpath, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))


# 다운로드 파일 확인 및 파일명 변경
def wait_for_download_and_rename(PDF_DIR, file_keyword, task_name, max_wait_time=30):
    try:
        elapsed_time = 0
        downloaded_file = None

        while elapsed_time < max_wait_time:
            files = [f for f in os.listdir(PDF_DIR) if file_keyword in f and not f.endswith('.crdownload')]
            # files = [f for f in os.listdir(PDF_DIR) if file_keyword in f and f.lower().f.endswith(('.pdf')]
          
            if files:
                downloaded_file = files[0]
                break
            time.sleep(1)
            elapsed_time += 1

        if not downloaded_file:
            return f"← {file_keyword} : (error) 파일 없음."  

        old_file_path = os.path.join(PDF_DIR, downloaded_file)
        new_file_path = os.path.join(PDF_DIR, task_name + os.path.splitext(downloaded_file)[1])

        # 파일명 변경
        if os.path.exists(new_file_path):
            os.remove(new_file_path)  # 동일 파일명 삭제

        os.rename(old_file_path, new_file_path)

        # 파일 사이즈 리턴
        file_size = os.path.getsize(new_file_path)  # 바이트 단위
        file_size_kb = round(file_size / 1024)  # KB 단위로 변환

        return "", f"{file_size_kb:,}KB"

    except Exception as e: return f"- {file_keyword} : (error) 파일명 변경 실패"


# 작업 실행 : 1.파일 메뉴 > 2.다운로드 메뉴 > 3.PDF 옵션 > 4.내보내기 > 5.다운로드 파일 확인 및 파일명 변경
def tasks_action(driver, task_name, file_keyword, page_url):    
    driver.get(page_url)   # 페이지 로드
    try:
        # 1. 파일 메뉴 클릭
        file_menu = wait_for_element_clickable(driver, "//*[@id='docs-file-menu']")
        file_menu.click()

        # 2. 다운로드 메뉴 클릭
        download_menu = wait_for_element_clickable(driver, "//div[@role='menuitem' and .//span[text()='다운로드']]")
        download_menu.click()

        # 3. PDF 문서 옵션 클릭
        pdf_option = wait_for_element_clickable(driver, "//span[contains(@aria-label, 'PDF')]")
        pdf_option.click()

        # 4. 내보내기 버튼 클릭
        send_button = wait_for_element_clickable(driver, "//span[text()='내보내기']")
        send_button.click()

        # 다운로드 파일 확인 및 파일명 변경
        rename_result, file_size_kb = wait_for_download_and_rename(PDF_DIR, file_keyword, task_name)

        if rename_result:  # 오류 메시지가 반환된 경우
            return rename_result  # 오류 메시지 반환

        return f"{file_size_kb} ← {file_keyword}"  # 성공 메시지 반환
    
    except Exception as e:
        print("(error) 작업 실행 실패")
        sys.exit(1)  # 시스템 오류 종료


# Main
if __name__ == "__main__":

    # 작업 목록
    task_list = [

        # 1.1 조직
        {'task_name': '1.1 조직_OS(조직현황)', 'file_keyword': 'OS(현황)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit#gid=840451947'},

        {'task_name': '1.2 조직_OC(조직도)','file_keyword': 'OC(조직도)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit#gid=1991745721'},

        {'task_name': '1.3 조직_WS(근무일정)','file_keyword': 'WS(근무)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit?gid=1309150508#gid=1309150508'},

        {'task_name': '1.4 조직_WC(근무달력)','file_keyword': 'WC(달력)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit?gid=1781033373#gid=1781033373'},


        # 2.1 CS영업
        {'task_name': '2.1 CS영업_CCR(보고)','file_keyword': 'CCR(보고)',
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=1418377442#gid=1418377442'},

        {'task_name': '2.2 CS영업_CSO(기회)','file_keyword': 'CSO(기회)',
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=1109450762#gid=1109450762'},

        {'task_name': '2.3 CS영업_ACR(활동)','file_keyword': 'ACR(활동)',
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=907484941#gid=907484941'},

        {'task_name': '2.4 CS영업_ACR(견적)','file_keyword': 'ACR(견적)',
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=788203034#gid=788203034'},

        {'task_name': '2.5 CS영업_MCR(수주)','file_keyword': 'MCR(수주)',
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=585653813#gid=585653813'},

        {'task_name': '2.6 CS영업_ACA(배정)','file_keyword': 'ACA(배정)',
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=1491915203#gid=1491915203'},


        # 3.1 CS사업
        {'task_name': '3.1 CS사업_BP(사업계획)','file_keyword': 'BP(계획)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=1685053532#gid=1685053532'},

        {'task_name': '3.2 CS사업_BA(사업실적)','file_keyword': 'BA(실적)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=158339800#gid=158339800'},

        {'task_name': '3.3 CS사업_AW(주간실적)','file_keyword': 'AW(주간)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=338286996#gid=338286996'},

        {'task_name': '3.4 CS사업_AM(월간실적)','file_keyword': 'AM(월간)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=1061543899#gid=1061543899'},

        {'task_name': '3.5 CS사업_MC(월별계약)','file_keyword': 'MC(계약)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=1157221142#gid=1157221142'},

        {'task_name': '3.6 CS사업_MU(투입)','file_keyword': 'MU(투입)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1XSv9GF6bpOaM9ktuYGRF7G3mh0d-OQUkzKSagCh_nvY/edit?gid=1687816592#gid=1687816592'},

        {'task_name': '3.7 CS사업_PU(투입)','file_keyword': 'PU(투입)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1XSv9GF6bpOaM9ktuYGRF7G3mh0d-OQUkzKSagCh_nvY/edit?gid=747947134#gid=747947134'},


        # # 4.1 SCC분석
        {'task_name': '4.1 SCC분석_CA(고객분석)','file_keyword': 'CA(고객)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1zc8Ozo55mplrRfK5ZMqbZhrlLov29WeuF7C5RM_r0ro/edit?gid=1759861482#gid=1759861482'},

        {'task_name': '4.2 SCC분석_MA(담당분석)','file_keyword': 'MA(담당)',
         'page_url': 'https://docs.google.com/spreadsheets/d/1XSv9GF6bpOaM9ktuYGRF7G3mh0d-OQUkzKSagCh_nvY/edit?gid=1200914694#gid=1200914694'}
    ]

    try:
        print("\n\n\n" + "─" * 60)        
        start_time = datetime.now()
        driver = setup_driver(PDF_DIR)  # 드라이버 초기화 (브라우저, 다운로드 디렉토리)
        clear_download_folder(PDF_DIR)  # 다운로드 디렉토리 초기화
        
        # 작업 지시
        print("─" * 60)
        for task in task_list:
            print(f"{task['task_name']}", end=" ")
            result = tasks_action(driver, task["task_name"], task["file_keyword"], task["page_url"])
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

