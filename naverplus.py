import sys
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QTextEdit, QPushButton, QVBoxLayout, QHBoxLayout, QFrame, QMessageBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import requests, time, random, pandas as pd, io

class User:
    def __init__(self, nickname, subscription_end, is_active=True):
        self.nickname, self.subscription_end, self.is_active = nickname, subscription_end, is_active

class UserManager:
    def __init__(self):
        self.users = {}
        self.excel_url = "https://docs.google.com/spreadsheets/d/1uVw-x0AlKJ21qo-e0DFlVWpah_cIgTay/export?format=xlsx"
        self.load_users()

    def load_users(self):
        try:
            df = pd.read_excel(io.BytesIO(requests.get(self.excel_url).content))
            for _, row in df.iterrows():
                if all(pd.notna(row[col]) for col in ['닉네임', '구독 만료 날짜', '활성여부']):
                    self.users[row['닉네임']] = User(row['닉네임'],
                                                     pd.to_datetime(row['구독 만료 날짜']).to_pydatetime(),
                                                     row['활성여부'].lower() == '활성')
            print(f"Loaded {len(self.users)} users successfully")
        except Exception as e:
            print(f"Error loading users: {e}")

    def check_user(self, nickname):
        user = self.users.get(nickname)
        if user:
            if user.is_active and user.subscription_end > datetime.now():
                return True
            return "User is inactive" if not user.is_active else "Subscription expired"
        return False

class WorkerThread(QThread):
    update_status = pyqtSignal(str)

    def __init__(self, user_id, user_pw, start_id, custom_message):
        super().__init__()
        self.user_id, self.user_pw, self.start_id, self.custom_message = user_id, user_pw, start_id, custom_message

    def run(self):
        self.update_status.emit("프로세스 시작...")
        driver = webdriver.Chrome()
        self.naver_login(driver)
        buddy_list = self.get_buddy_list()
        for blog_id in buddy_list:
            self.update_status.emit(f'{blog_id}의 블로그 서로이웃 추가 중')
            self.add_neighbor(driver, blog_id)
        driver.quit()
        self.update_status.emit("프로세스 완료")

    def naver_login(self, driver):
        driver.get('https://nid.naver.com/nidlogin.login?mode=form')
        time.sleep(random.randint(2, 6))
        driver.find_element(By.CSS_SELECTOR, '#id').send_keys(self.user_id)
        time.sleep(random.randint(2, 6))
        driver.find_element(By.CSS_SELECTOR, '#pw').send_keys(self.user_pw)
        time.sleep(random.randint(2, 6))
        driver.find_element(By.XPATH, '//*[@id="log.login"]').click()
        time.sleep(30)

    def get_buddy_list(self):
        response = requests.get(f'https://m.blog.naver.com/BuddyList.naver?blogId={self.start_id}')
        return [a.attrs['href'].split('/')[3] for a in BeautifulSoup(response.text, 'html.parser').find_all("a") if 'href' in a.attrs]

    def add_neighbor(self, driver, blog_id):
        driver.get(f'https://m.blog.naver.com/{blog_id}')
        time.sleep(random.randint(2, 6))
        try:
            neighbor_button = driver.find_element(By.XPATH, '//*[@id="root"]/div[4]/div/div[3]/div[1]/button')
            if neighbor_button.text == '이웃추가':
                neighbor_button.click()
                time.sleep(random.randint(2, 6))
                driver.find_element(By.XPATH, '//*[@id="bothBuddyRadio"]').click()
                driver.find_element(By.XPATH, '//*[@id="buddyAddForm"]/fieldset/div/div[2]/div[3]/div/textarea').send_keys(self.custom_message)
                time.sleep(random.randint(2, 6))
                driver.find_element(By.XPATH, '/html/body/ui-view/div[2]/a[2]').click()
                time.sleep(random.randint(2, 6))
        except:
            self.update_status.emit(f'{blog_id} 블로그 이웃 추가 실패')

class UserModeGUI(QWidget):
    def __init__(self, user_manager):
        super().__init__()
        self.user_manager = user_manager
        self.initUI()

    def initUI(self):
        main_layout = QHBoxLayout()
        left_frame, right_frame = QFrame(self), QFrame(self)
        left_layout, right_layout = QVBoxLayout(left_frame), QVBoxLayout(right_frame)

        self.nickname_input, self.id_input, self.pw_input = QLineEdit(), QLineEdit(), QLineEdit()
        self.pw_input.setEchoMode(QLineEdit.Password)
        self.start_id_input, self.custom_message_input = QLineEdit(), QTextEdit()
        self.custom_message_input.setFixedHeight(100)

        for label, widget in [('닉네임', self.nickname_input), ('네이버 ID', self.id_input),
                              ('비밀번호', self.pw_input), ('시작 블로그 ID', self.start_id_input),
                              ('이웃 추가 메시지', self.custom_message_input)]:
            left_layout.addWidget(QLabel(label))
            left_layout.addWidget(widget)

        button_layout = QHBoxLayout()
        start_button, refresh_button = QPushButton('실행'), QPushButton('새로고침')
        start_button.clicked.connect(self.check_user)
        refresh_button.clicked.connect(self.refresh_data)
        button_layout.addWidget(start_button)
        button_layout.addWidget(refresh_button)
        left_layout.addLayout(button_layout)

        right_layout.addWidget(QLabel('작업 진행 상태창'))
        self.status_display = QTextEdit()
        self.status_display.setReadOnly(True)
        right_layout.addWidget(self.status_display)

        main_layout.addWidget(left_frame, 1)
        main_layout.addWidget(right_frame, 1)
        self.setLayout(main_layout)
        self.setWindowTitle('네이버 블로그 서로이웃 자동화')
        self.setGeometry(300, 300, 800, 400)
        self.setStyleSheet("""
            QWidget {font-family: Arial; font-size: 12px;}
            QLabel {font-weight: bold;}
            QLineEdit, QTextEdit {border: 1px solid #BDC3C7; padding: 3px;}
            QPushButton {background-color: #3498DB; color: white; padding: 5px 10px; border: none; border-radius: 3px;}
            QPushButton:hover {background-color: #2980B9;}
        """)

    def refresh_data(self):
        self.user_manager.load_users()
        self.status_display.append("사용자 데이터가 새로고침되었습니다.")

    def check_user(self):
        nickname = self.nickname_input.text()
        if not nickname:
            QMessageBox.warning(self, '오류', '닉네임을 입력해주세요.')
        else:
            result = self.user_manager.check_user(nickname)
            if result is True:
                self.start_process()
            else:
                QMessageBox.warning(self, '오류', result)
                self.disable_inputs()

    def disable_inputs(self):
        for widget in [self.id_input, self.pw_input, self.start_id_input, self.custom_message_input]:
            widget.setEnabled(False)

    def start_process(self):
        self.thread = WorkerThread(self.id_input.text(), self.pw_input.text(),
                                   self.start_id_input.text(), self.custom_message_input.toPlainText())
        self.thread.update_status.connect(self.status_display.append)
        self.thread.start()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = UserModeGUI(UserManager())
    ex.show()
    sys.exit(app.exec_())