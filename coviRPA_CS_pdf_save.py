import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# 다운로드 디렉토리 설정
download_dir = r"C:\Users\jhlee6\OneDrive - Covision\21 CSreport\PDF"


# Chrome WebDriver를 초기화하고 다운로드 디렉토리 설정
def setup_driver(download_dir):
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("user-data-dir=C:\\user_data\\junho")  # 사용자 데이터 경로
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    # 다운로드 디렉토리 설정
    prefs = {
        "download.default_directory": download_dir,  # 다운로드 폴더 설정
        "download.prompt_for_download": False,       # 다운로드 대화 상자 비활성화
        "download.directory_upgrade": True,          # 기존 다운로드 경로 업데이트 허용
        "safebrowsing.enabled": True,                # 안전 브라우징 활성화
        # "plugins.always_open_pdf_externally": True   # PDF 파일 외부에서 열기
    }
    options.add_experimental_option("prefs", prefs)

    # WebDriver 생성
    return webdriver.Chrome(options=options)


# 다운로드 폴더 초기화
def clear_download_folder(download_dir):  
    for file in os.listdir(download_dir):
        file_path = os.path.join(download_dir, file)
        os.remove(file_path)


# 지정된 XPATH를 가진 요소가 DOM에 존재할 때까지 대기하고 반환
def wait_for_element_presence(driver, xpath, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))


# 지정된 XPATH를 가진 요소가 클릭 가능할 때까지 대기하고 반환
def wait_for_element_clickable(driver, xpath, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))


# 다운로드 파일 확인 및 파일명 변경
def wait_for_download_and_rename(download_dir, file_keyword, task_name, max_wait_time=30):
    try:
        elapsed_time = 0
        downloaded_file = None

        while elapsed_time < max_wait_time:
            files = [f for f in os.listdir(download_dir) if file_keyword in f and not f.endswith('.crdownload')]
            # files = [f for f in os.listdir(download_dir) if file_keyword in f and f.lower().f.endswith(('.pdf')]
          
            if files:
                downloaded_file = files[0]
                break
            time.sleep(1)
            elapsed_time += 1

        if not downloaded_file: raise FileNotFoundError(f"- {file_keyword} : (!)파일 없음")

        old_file_path = os.path.join(download_dir, downloaded_file)
        new_file_path = os.path.join(download_dir, task_name + os.path.splitext(downloaded_file)[1])

        if os.path.exists(new_file_path): os.remove(new_file_path)
        os.rename(old_file_path, new_file_path)

    except Exception as e: raise Exception(f"- {file_keyword} : (!)파일명 변경 오류")


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

        # 5. 다운로드 파일 확인 및 파일명 변경
        wait_for_download_and_rename(download_dir, file_keyword, task_name)
        return f"- {file_keyword} : 완료"
    
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
        # 작업 지시
        for task in task_list:
            print(f"{task['task_name']}", end=" ")
            result = tasks_action(driver, task["task_name"], task["file_keyword"], task["page_url"])
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
