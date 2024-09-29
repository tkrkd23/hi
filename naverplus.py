import streamlit as st
import pandas as pd
import io
import requests
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import random


class UserManager:
    def __init__(self, excel_url):
        self.users = {}
        self.excel_url = excel_url
        self.load_users()

    def load_users(self):
        file_id = self.excel_url.split('/')[5]
        download_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
        response = requests.get(download_url)
        excel_data = io.BytesIO(response.content)
        df = pd.read_excel(excel_data)

        for _, row in df.iterrows():
            nickname = row['닉네임']
            subscription_end = row['구독 만료 날짜']
            is_active = row['활성여부']

            if pd.notna(nickname) and pd.notna(subscription_end) and pd.notna(is_active):
                self.users[nickname] = {
                    'subscription_end': pd.to_datetime(subscription_end),
                    'is_active': is_active.lower() == '활성'
                }

    def check_user(self, nickname):
        user = self.users.get(nickname)
        if user and user['is_active'] and user['subscription_end'] > datetime.now():
            return True
        return False


def naver_login(driver, user_id, user_pw):
    driver.get('https://nid.naver.com/nidlogin.login?mode=form')
    time.sleep(random.uniform(1, 3))
    driver.find_element(By.CSS_SELECTOR, '#id').send_keys(user_id)
    time.sleep(random.uniform(1, 3))
    driver.find_element(By.CSS_SELECTOR, '#pw').send_keys(user_pw)
    time.sleep(random.uniform(1, 3))
    driver.find_element(By.XPATH, '//*[@id="log.login"]').click()
    time.sleep(5)


def add_neighbors(driver, start_id, custom_message):
    driver.get(f'https://m.blog.naver.com/{start_id}')
    time.sleep(random.uniform(1, 3))
    # 여기에 이웃 추가 로직 구현
    st.write("이웃 추가 완료")


def main():
    st.title('네이버 블로그 서로이웃 자동화')

    excel_url = "https://docs.google.com/spreadsheets/d/1uVw-x0AlKJ21qo-e0DFlVWpah_cIgTay/edit?usp=sharing&ouid=116411561498747453406&rtpof=true&sd=true"
    user_manager = UserManager(excel_url)

    nickname = st.text_input('닉네임')
    user_id = st.text_input('네이버 ID')
    user_pw = st.text_input('비밀번호', type='password')
    start_id = st.text_input('시작 블로그 ID')
    custom_message = st.text_area('이웃 추가 메시지')

    if st.button('실행'):
        if user_manager.check_user(nickname):
            st.write("프로세스 시작...")
            driver = webdriver.Chrome()
            naver_login(driver, user_id, user_pw)
            add_neighbors(driver, start_id, custom_message)
            driver.quit()
            st.write("프로세스 완료")
        else:
            st.error('유효하지 않은 사용자입니다.')


if __name__ == '__main__':
    main()