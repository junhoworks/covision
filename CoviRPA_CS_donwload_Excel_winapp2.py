import tkinter as tk
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import threading
import time

class WebDriverApp:
    def __init__(self, root):
        self.root = root
        self.root.title("웹 드라이버 실행")
        self.root.geometry("400x300")

        # 메시지를 표시할 라벨
        self.message_label = tk.Label(root, text="메시지가 여기 표시됩니다.", width=50, height=5)
        self.message_label.pack(pady=20)

        # 실행 버튼
        self.run_button = tk.Button(root, text="실행", command=self.start_browser)
        self.run_button.pack(pady=10)

        # 중지 버튼
        self.stop_button = tk.Button(root, text="중지", command=self.stop_browser, state=tk.DISABLED)
        self.stop_button.pack(pady=10)

        # 웹 드라이버 상태
        self.driver = None
        self.thread = None

    def update_message(self, message):
        """메시지 라벨을 업데이트"""
        self.message_label.config(text=message)

    def start_browser(self):
        """웹 브라우저를 시작하는 함수"""
        self.update_message("웹 드라이버 실행 중...")
        self.run_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        # 새로운 쓰레드에서 웹 드라이버 실행
        self.thread = threading.Thread(target=self.run_browser)
        self.thread.start()

    def run_browser(self):
        """웹 드라이버를 실제로 실행하는 함수"""
        options = Options()
        options.add_argument("--headless")  # 헤드리스 모드 (브라우저 창을 띄우지 않음)
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        try:
            self.driver = webdriver.Chrome(options=options)
            self.update_message("웹 드라이버 시작됨!")

            # 여기에 웹 드라이버로 작업을 추가할 수 있습니다.
            # 예를 들어, 원하는 페이지로 이동
            self.driver.get("https://www.example.com")
            time.sleep(3)

            self.update_message("웹 페이지 로딩 완료!")

            self.update_message("언제까지 Hello World 만 찍을거야?\n 놀지 말고 너도 개발을 해!")

        except Exception as e:
            self.update_message(f"오류 발생: {e}")

    def stop_browser(self):
        """웹 드라이버를 중지하는 함수"""
        if self.driver:
            self.driver.quit()
            self.update_message("웹 드라이버 종료됨.")
        self.run_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = WebDriverApp(root)
    root.mainloop()
