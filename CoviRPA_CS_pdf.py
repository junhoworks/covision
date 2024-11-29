import os
import pyautogui
import time
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime


# 시작 작업
start_time = datetime.now()
print(f"\n\n\n#시작시간 : {start_time.strftime('%H:%M:%S')}\n")


# 다운로드 경로 설정
download_dir = "C:\\Users\\jhlee6\\OneDrive - Covision\\21 CSreport\\PDF"

# Chrome 옵션 설정
options = Options()
options.add_argument('--start-maximized')  # 창을 최대화 상태로 시작
options.add_argument('user-data-dir=C:\\user_data\\junho')  # 사용자 데이터 경로
options.add_argument('disable-blink-features=AutomationControlled')  # 자동화 탐지 방지
options.add_experimental_option('detach', True)  # 브라우저 창이 닫히지 않도록 설정2.3 CS영업_ACR(활동).pdf


options.add_experimental_option('excludeSwitches', ['enable-logging'])  # 불필요한 로그 제거


# 프린트 대화상자 설정
prefs = {
    "savefile.default_directory": download_dir,
    "print.preview": "false",  # 프린트 미리보기 없이 바로 출력
    "printing.print_preview_sticky": "false",  # 미리보기 설정 비활성화
}
options.add_experimental_option("prefs", prefs)


# ChromeDriver 실행
driver = webdriver.Chrome(options=options)


# 기존 파일이 존재하면 삭제하는 함수
def delete_existing_file(file_name, download_dir):
    file_path = os.path.join(download_dir, file_name)
    if os.path.exists(file_path):
        try:
            os.remove(file_path) # 기존 파일 삭제
        except Exception as e:
            print("ERROR")


# 시트 열고 PDF 저장
def open_and_save_sheet(sheet_url, file_name, load_time, driver):
    driver.get(sheet_url)
    wait = WebDriverWait(driver, 30)  # 최대 30초 대기

    # XPath 정의
    print_button_xpath = '//*[@id="t-print"]/div/div/div'  # 메뉴의 인쇄 버튼 XPath
    next_button_xpath = "//span[text()='다음']"  # 인쇄설정의 다음 버튼 XPath

    try:
        # 작업명
        print(f"{file_name} ...", end=" ")
        
        # 페이지가 완전히 로드될 때까지 기다림
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

        # 추가적인 로딩 대기 시간
        time.sleep(load_time)  # 시트마다 설정된 로딩 타임을 적용

        # 인쇄 버튼 클릭
        print_button = wait.until(EC.element_to_be_clickable((By.XPATH, print_button_xpath)))
        print_button.click()
        time.sleep(2)

        # 다음 버튼 클릭
        next_button_option = wait.until(EC.visibility_of_element_located((By.XPATH, next_button_xpath)))
        next_button_option.click()
        time.sleep(2)

        # 인쇄 대화창에서 Enter키 입력으로 저장
        pyautogui.hotkey('enter')  # Enter 키 입력
        time.sleep(5)

        # 기존에 같은 이름의 파일이 있으면 삭제
        delete_existing_file(file_name, download_dir)

        # 다른 이름으로 저장 대화창이 뜨면 파일 경로 입력 (pyperclip을 사용해 한글 파일명 클립보드에 복사 후 붙여넣기)
        pyperclip.copy(file_name)  # 한글 파일명을 클립보드에 복사
        pyautogui.hotkey('ctrl', 'v')  # 클립보드의 내용을 붙여넣기
        time.sleep(2)  # 경로 입력 후 잠시 대기
        pyautogui.hotkey('enter')  # 저장
        time.sleep(3)  # 저장 버튼 클릭 후 잠시 대기

    except Exception as e:
        print("ERROR")  # 작업중 ERROR
    finally:
        print("OK")


# 시트 URL, 파일명, 로딩 타임 (초 단위) 목록
sheet_info = [
    ('https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit?gid=840451947#gid=840451947', '1.1 조직_OS(조직현황).pdf', 10),
    ('https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit?gid=1991745721#gid=1991745721', '1.2 조직_OC(조직도).pdf', 7),
    ('https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit?gid=1309150508#gid=1309150508', '1.3 조직_WS(근무일정).pdf', 7),
    ('https://docs.google.com/spreadsheets/d/1VrZ2KBT10uGW6b4RZ0t4nxWmwCP6V3etfka3Rhyhh1M/edit?gid=1781033373#gid=1781033373', '1.4 조직_WS(근무달력).pdf', 7),

    ('https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=1418377442#gid=1418377442', '2.1 CS영업_CCR(보고).pdf', 10),
    ('https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=1109450762#gid=1109450762', '2.2 CS영업_CSO(기회).pdf', 5),
    ('https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=907484941#gid=907484941', '2.3 CS영업_ACR(활동).pdf', 5),
    ('https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=788203034#gid=788203034', '2.4 CS영업_ACR(견적).pdf', 5),
    ('https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=585653813#gid=585653813', '2.5 CS영업_MCR(수주).pdf', 5),
    ('https://docs.google.com/spreadsheets/d/174_yFSydS0eCActdZaqslXVBspKTNAbJOvOPpab1_Hw/edit?gid=1491915203#gid=1491915203', '2.6 CS영업_ACA(배정).pdf', 5),

    ('https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=1685053532#gid=1685053532', '3.1 CS사업_BP(사업계획).pdf', 15),
    ('https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=158339800#gid=158339800', '3.2 CS사업_BA(사업실적).pdf', 10),
    ('https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=338286996#gid=338286996', '3.3 CS사업_AW(주간실적).pdf', 15),
    ('https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=1061543899#gid=1061543899', '3.4 CS사업_AM(월간실적).pdf', 10),
    ('https://docs.google.com/spreadsheets/d/1_jTUNxq8KGL7TIjySLb6zEcLFxzrG-sos8KKqCVNRWY/edit?gid=1157221142#gid=1157221142', '3.5 CS사업_MC(월별계약).pdf', 5),
    ('https://docs.google.com/spreadsheets/d/1XSv9GF6bpOaM9ktuYGRF7G3mh0d-OQUkzKSagCh_nvY/edit?gid=1687816592#gid=1687816592', '3.6 CS사업_MU(투입).pdf', 5),
    ('https://docs.google.com/spreadsheets/d/1XSv9GF6bpOaM9ktuYGRF7G3mh0d-OQUkzKSagCh_nvY/edit?gid=747947134#gid=747947134', '3.7 CS사업_PU(투입).pdf', 5),
 
    ('https://docs.google.com/spreadsheets/d/1zc8Ozo55mplrRfK5ZMqbZhrlLov29WeuF7C5RM_r0ro/edit?gid=1759861482#gid=1759861482', '4.1 SCC분석_CA(고객분석).pdf', 10),
    ('https://docs.google.com/spreadsheets/d/1XSv9GF6bpOaM9ktuYGRF7G3mh0d-OQUkzKSagCh_nvY/edit?gid=1200914694#gid=1200914694', '4.2 SCC분석_MA(담당분석).pdf', 10)
]


# 각 시트 URL과 파일명, 로딩 타임에 대해 작업 수행
for sheet_url, file_name, load_time in sheet_info:
    try:
        open_and_save_sheet(sheet_url, file_name, load_time, driver)
        time.sleep(3)  # 각 시트 작업 후 잠시 대기
    except Exception as e:
        print("작업중 ERROR")


# 마무리 작업
print("\nCLOSE .......................", end=" ")
driver.quit()  #브라우저 종료
print("OK")

# 종료 작업
end_time = datetime.now()
print(f"\n#종료시간: {end_time.strftime('%H:%M:%S')}")
elapsed_time = end_time - start_time
print(f"#실행시간: {str(elapsed_time).split('.')[0]}\n\n\n")
