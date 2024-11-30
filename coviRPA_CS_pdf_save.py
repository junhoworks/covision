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


# Chrome WebDriver를 초기화하고 다운로드 경로 설정
def initialize_webdriver(download_dir):
    """Chrome WebDriver를 초기화하고 다운로드 경로를 설정합니다."""
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("user-data-dir=C:\\user_data\\junho")  # 사용자 데이터 경로
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    # 다운로드 폴더 설정
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)

    # WebDriver 생성
    return webdriver.Chrome(options=options)


# 다운로드 폴더 초기화
def clear_download_folder(directory):  
    print(f"download_folder initialize")
    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)
        os.remove(file_path)


# Google 스프레드시트에서 PDF 파일 다운로드 : 파일 > 다운로드 > PDF > 내보내기
def download_pdf(driver, page_url, file_keyword, loading_time, save_time):        
    driver.get(page_url)
    wait = WebDriverWait(driver, 30)
    time.sleep(loading_time)

    try:
        # 1. 파일 메뉴 클릭
        file_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='docs-file-menu']")))
        file_menu.click()

        # 2. 다운로드 메뉴 클릭
        download_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='menuitem' and .//span[text()='다운로드']]")))
        download_menu.click()
        time.sleep(1)  # 잠시 대기

        # 3. PDF 문서 옵션 클릭
        pdf_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@aria-label, 'PDF')]")))
        pdf_option.click()
        time.sleep(1)  # 잠시 대기

        # 4. 내보내기 버튼 클릭
        send_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='내보내기']")))
        send_button.click()
        time.sleep(save_time)
    except Exception as e: return None


# 파일명 일괄 변경
def rename_files(download_dir, task_list):
    try:
        for task in task_list:
            file_keyword = task["file_keyword"]
            task_name = task["task_name"]

            # 키워드에 해당하는 파일 찾기
            files = [
                f for f in os.listdir(download_dir)
                if file_keyword in f and f.lower().endswith(".pdf")
            ]
            if not files:
                print(f"- {file_keyword} : 파일없음")
                continue

            # 가장 최근 파일 선택
            latest_file = max([os.path.join(download_dir, f) for f in files], key=os.path.getctime)

            # 새 파일 이름 설정
            file_extension = os.path.splitext(latest_file)[1]
            new_file_path = os.path.join(download_dir, f"{task_name}{file_extension}")

            # 기존 파일 삭제 후 이름 변경
            if os.path.exists(new_file_path):
                os.remove(new_file_path)
            os.rename(latest_file, new_file_path)
            time.sleep(1)  # 잠시 대기
    except Exception as e: return None


# Main
if __name__ == "__main__":
    start_time = datetime.now()
    print(f"\n\n\n[시작시간] {start_time.strftime('%H:%M:%S')}")
    clear_download_folder(download_dir)  # 다운로드 폴더 초기화
    print("\n-------------------------------")

    # 작업 목록 정의
    task_list = [

        # 1.1 조직
        {'task_name': '1.1 조직_OS(조직현황)', 'file_keyword': 'OS(현황)', 'loading_time': 10, 'save_time': 4,
         'page_url': 'https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit#gid=840451947'},

        {'task_name': '1.2 조직_OC(조직도)','file_keyword': 'OC(조직도)','loading_time': 5, 'save_time': 3,
         'page_url': 'https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit#gid=1991745721'},

        {'task_name': '1.3 조직_WS(근무일정)','file_keyword': 'WS(근무)','loading_time': 3, 'save_time': 3,
         'page_url': 'https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit?gid=1309150508#gid=1309150508'},

        {'task_name': '1.4 조직_WC(근무달력)','file_keyword': 'WC(달력)','loading_time': 3, 'save_time': 3,
         'page_url': 'https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit?gid=1781033373#gid=1781033373'},


        # 2.1 CS영업
        {'task_name': '2.1 CS영업_CCR(보고)','file_keyword': 'CCR(보고)','loading_time': 8, 'save_time': 4,
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=1418377442#gid=1418377442'},

        {'task_name': '2.2 CS영업_CSO(기회)','file_keyword': 'CSO(기회)','loading_time': 4, 'save_time': 5,
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=1109450762#gid=1109450762'},

        {'task_name': '2.3 CS영업_ACR(활동)','file_keyword': 'ACR(활동)','loading_time': 4, 'save_time': 8,
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=907484941#gid=907484941'},

        {'task_name': '2.4 CS영업_ACR(견적)','file_keyword': 'ACR(견적)','loading_time': 4, 'save_time': 5,
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=788203034#gid=788203034'},

        {'task_name': '2.5 CS영업_MCR(수주)','file_keyword': 'MCR(수주)','loading_time': 4, 'save_time': 4,
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=585653813#gid=585653813'},

        {'task_name': '2.6 CS영업_ACA(배정)','file_keyword': 'ACA(배정)','loading_time': 4, 'save_time': 3,
         'page_url': 'https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=1491915203#gid=1491915203'},


        # 3.1 CS사업
        {'task_name': '3.1 CS사업_BP(사업계획)','file_keyword': 'BP(계획)','loading_time': 12, 'save_time': 4,
         'page_url': 'https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=1685053532#gid=1685053532'},

        {'task_name': '3.2 CS사업_BA(사업실적)','file_keyword': 'BA(실적)','loading_time': 5, 'save_time': 4,
         'page_url': 'https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=158339800#gid=158339800'},

        {'task_name': '3.3 CS사업_AW(주간실적)','file_keyword': 'AW(주간)','loading_time': 5, 'save_time': 5,
         'page_url': 'https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=338286996#gid=338286996'},

        {'task_name': '3.4 CS사업_AM(월간실적)','file_keyword': 'AM(월간)','loading_time': 5, 'save_time': 4,
         'page_url': 'https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=1061543899#gid=1061543899'},

        {'task_name': '3.5 CS사업_MC(월별계약)','file_keyword': 'MC(계약)','loading_time': 5, 'save_time': 4,
         'page_url': 'https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=1157221142#gid=1157221142'},

        {'task_name': '3.6 CS사업_MU(투입)','file_keyword': 'MU(투입)','loading_time': 5, 'save_time': 3,
         'page_url': 'https://docs.google.com/spreadsheets/d/1XSv9GF6bpOaM9ktuYGRF7G3mh0d-OQUkzKSagCh_nvY/edit?gid=1687816592#gid=1687816592'},

        {'task_name': '3.7 CS사업_PU(투입)','file_keyword': 'PU(투입)','loading_time': 5, 'save_time': 3,
         'page_url': 'https://docs.google.com/spreadsheets/d/1XSv9GF6bpOaM9ktuYGRF7G3mh0d-OQUkzKSagCh_nvY/edit?gid=747947134#gid=747947134'},


        # 4.1 SCC분석
        {'task_name': '4.1 SCC분석_CA(고객분석)','file_keyword': 'CA(고객)','loading_time': 10, 'save_time': 5,
         'page_url': 'https://docs.google.com/spreadsheets/d/1zc8Ozo55mplrRfK5ZMqbZhrlLov29WeuF7C5RM_r0ro/edit?gid=1759861482#gid=1759861482'},

        {'task_name': '4.2 SCC분석_MA(담당분석)','file_keyword': 'MA(담당)','loading_time': 8, 'save_time': 3,
         'page_url': 'https://docs.google.com/spreadsheets/d/1XSv9GF6bpOaM9ktuYGRF7G3mh0d-OQUkzKSagCh_nvY/edit?gid=1200914694#gid=1200914694'}
    ]

    driver = initialize_webdriver(download_dir)  # WebDriver 초기화

    try:
        # 작업 수행
        for task in task_list:
            print(f"{task['task_name']}")
            result = download_pdf(driver, task["page_url"], task["file_keyword"], task["loading_time"], task["save_time"])

        # 파일 이름 변경
        print("-------------------------------")
        print("File rename erorr list")
        rename_files(download_dir, task_list)
        print("-------------------------------\n")

        
    finally:
        # WebDriver 종료
        print("WebDriver close")
        driver.quit()

        # 실행 시간 출력
        end_time = datetime.now()
        print(f"[종료시간] {end_time.strftime('%H:%M:%S')}")
        elapsed_time = end_time - start_time
        print(f"[실행시간] {str(elapsed_time).split('.')[0]}\n")
