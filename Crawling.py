import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

def getreview(number):
    return row.find_element(By.XPATH, f'//*[@id="infoset_reviewContentList"]/div[{number}]').find_element(By.XPATH,f'//*[@id="infoset_reviewContentList"]/div[{number}]/div[1]/div/span/span').text + "," + row.find_element(By.XPATH, f'//*[@id="infoset_reviewContentList"]/div[{number}]').find_element(By.XPATH,f'//*[@id="infoset_reviewContentList"]/div[{number}]/div[1]/div/em[4]').text

# ChromeDriver 경로 설정
chrome_service = Service('C:/Users/user/Downloads/chromedriver-win64(130.0.6723.31)/chromedriver-win64/chromedriver.exe')  # ChromeDriver의 경로를 지정하세요
chrome_options = Options()
chrome_options.add_argument('--disable-gpu')

# 웹드라이버 시작
driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

file_path = 'C:/Users/user/Downloads/베스트셀러.xlsx'  # 실제 파일 경로로 변경하세요

# Excel 파일 읽기
df = pd.read_excel(file_path)

# 두 번째 열 데이터 가져오기 (인덱스는 0부터 시작하므로 1)
second_column_data = df.iloc[:, 2]  # 두 번째 열

for i in second_column_data:
    # 크롤링할 URL
    url = f'https://www.yes24.com/Product/Goods/{i}'  # 크롤링할 웹사이트 URL을 입력하세요
    driver.get(url)

    # 페이지가 로드될 때까지 대기
    time.sleep(10)  # 필요에 따라 대기 시간을 조정하세요

    # 리뷰 내용 출력
    count = 2
    try:
        element = driver.find_element(By.XPATH,'//*[@id="total"]/a/span')
        element.click()
        time.sleep(3)
        with open(f'{i}.csv', mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                # 컬럼명 작성
                writer.writerow(["productcode","review1", "review2", "review3", "review4", "review5"])
                while True:
                    try:
                        reviews = driver.find_elements(By.ID, 'infoset_reviewContentList')
                        for row in reviews:
                                    # 각 셀에서 데이터를 추출
                            productcode = i
                            review1 = getreview(2)
                            review2 = getreview(3)
                            review3 = getreview(4)
                            review4 = getreview(5)
                            review5 = getreview(6)
                            writer.writerow([productcode,review1, review2, review3, review4, review5])
                            if count == 10:
                                driver.find_element(By.XPATH,f'//*[@id="infoset_reviewContentList"]/div[1]/div[1]/div/a[12]').click()
                                count = 1
                                time.sleep(3)
                            else:
                                driver.find_element(By.XPATH,f'//*[@id="infoset_reviewContentList"]/div[1]/div[1]/div/a[{count+2}]').click()
                                count+=1
                                time.sleep(3)
                    except Exception as e:
                        print(e)
                        break
    except  Exception as e:
        print(e)
        continue
driver.quit()