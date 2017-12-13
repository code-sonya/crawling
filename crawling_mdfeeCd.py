# -*- coding: utf-8 -*-

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import re
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import xlrd
from generate_csv import Generate_csv

def add_lst(code):
        global failed_lst
        failed_lst.append(code)

class OpendataCrawler(object):
    def __init__(self, url, md_code, data_btn ,directory , by):
        self.url = url
        self.md_code = md_code
        self.data_btn = data_btn
        #self.count = count
        self.directory = directory
        self.by = by

    def crawl_data(self):
        for code in self.md_code:
            try:
                url = 'http://opendata.hira.or.kr/op/opc/olapDiagBhvInfo.do'
                chromedriver = './chromedriver'
                driver = webdriver.Chrome(chromedriver)
                driver.get(url)

                # main : 메인 창
                main = driver.current_window_handle
                driver.find_element_by_xpath('//*[@id="searchPopup"]').click()

                # popup : 팝업 창
                popup = driver.window_handles
                driver.switch_to_window(popup.pop())
                driver.implicitly_wait(20)

                # 코드명으로 데이터 조회
                try:
                    driver.find_element_by_xpath('//*[@id="searchWrd1"]').send_keys(code)
                except Exception as e:
                    fail_lst.append(code)

                # 진료행위명칭 클릭
                driver.find_element_by_css_selector('a[id="searchBtn1"]').send_keys("\n")
                driver.implicitly_wait(3)
                driver.find_element_by_xpath('//*[@id="tab1"]/section[2]/table/tbody/tr/td[2]/a').click()

                # 메인 창으로 전환
                driver.switch_to_window(main)
                driver.find_element_by_xpath(self.data_btn).click()

                # iframe : iframe 영역
                iframe = driver.find_element_by_class_name('olapViewFrame')
                driver.switch_to_frame(iframe)
                driver.implicitly_wait(20)


                #iframe
                try:
                    wait = WebDriverWait(driver, 10)

                    ## 진료년월 라디오 버튼
                    radio = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,'#ext-gen1645 > table > tbody > tr > td > div:nth-child(2) > table > tbody > tr > td > ul > li:nth-child(2) > label > input[type="radio"]')))
                    radio = driver.find_element_by_css_selector('#ext-gen1645 > table > tbody > tr > td > div:nth-child(2) > table > tbody > tr > td > ul > li:nth-child(2) > label > input[type="radio"]')
                    driver.execute_script("arguments[0].click();", radio)

                    ### 검색 시작 날짜 선택
                    date_to = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,"#ext-gen1645 > table > tbody > tr > td > div:nth-child(4) > div > input:nth-child(2)")))
                    date_to.click()


                    ## 정규표현식으로 id 찾기 (엘리먼트의 id값이 날자를 기준으로 페이지 로드 마다 계속해서 바뀜)
                    html = driver.page_source
                    soup = BeautifulSoup(html, 'html.parser')
                    id_css = re.search(r'monthpicker_\d+', str(soup))

                    ## 달력에서 년도 선택
                    year_css = '#'+ id_css.group() +' > div > select > option:nth-child(7)'
                    year = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, year_css)))
                    year.click()

                    ## 달력에서 월 선택
                    month_css = '#' + id_css.group() + ' > table > tbody > tr:nth-child(1) > td:nth-child(1)'
                    month = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, month_css)))
                    month.click()


                    ## 조회 버튼
                    search_btn = driver.find_element_by_class_name("dt-btn-search")
                    driver.execute_script("arguments[0].click();", search_btn)

                    ## 데이터가 로드될 때 까지 기다리기
                    try:
                        datagrid = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#panel-1184-body > div.dock_main > div.dock_inner div.m-datagrid-cell')))

                    ## 조회된 데이터가 없는 경우
                    except Exception as e:
                        driver.close()

                    ## 엑셀파일 다운로드
                    # download_excel : 엑셀 다운로드 하는 버튼
                    download_excel = driver.find_element_by_css_selector('#panel-1184-body > div:nth-child(2) > div.dock_inner > div.dock_title.doc_title_normal.mec-report-titlebar > div.dock_title_btnarea > div.dock_button.dock_exp_excel_button')
                    driver.execute_script("arguments[0].click();", download_excel)

                    ## csv 파일명
                    global count
                    if count == 0:
                        file_name = self.directory +self.by+'%28진료년월%29.xls'
                    else:
                        file_name = self.directory +self.by+'%28진료년월%29 ('+str(count)+').xls'

                    ## 엑셀파일이 다운로드될 때 까지 기다리기
                    print ('''Downloading file....{}'''.format(code))
                    while True:
                        try:
                            workbook = xlrd.open_workbook(file_name)
                            count += 1
                            break

                        except Exception as e:
                            continue

                    ## driver 종료
                    driver.implicitly_wait(20)

                    global failed_lst
                    if code in failed_lst:
                        failed_lst.remove(code)
                    driver.close()

                except Exception as e:
                    print (e)
                    driver.close()

            except Exception as e:
                add_lst(code)
                #failed_lst.append(code)
                driver.close()



# institution_btn : 요양기관종별 , location_btn : 요양기관소재지별 , directory : csv 파일 다운로드 되는 위치 (수정 필요)
institution_btn = '/html/body/section[1]/section[2]/div[1]/ul/li[4]'
location_btn = '/html/body/section[1]/section[2]/div[1]/ul/li[5]'
by_institution = '1_진료행위요양기관그룹별현황'
by_location = '1_진료행위요양기관소재지별현황'
url = 'http://opendata.hira.or.kr/op/opc/olapDiagBhvInfo.do'
directory = '/Users/sonya/Downloads/'

# mdfeeCd_lst : 진료행위 코드 리스트
df = pd.read_excel('csv_file/mdfeeCd.xlsx')
mdfeeCd_lst = df['mdfeeCd']
global failed_lst
failed_lst = []
global count
count = 0

# url, md_code, data_btn , count , directory , by):
## crawler_ins1 : 요양기관종별 데이터 크롤링
print ('crawling Data ..............')
crawler_ins1 = OpendataCrawler(url, mdfeeCd_lst, institution_btn, directory , by_institution)
crawler_ins1.crawl_data()

#크롤링 실패한 리스트 가지고 다시 객체 생성 후 데이터 크롤링
print ('failed_lst = ',len(failed_lst))
print ('crawling failed_lst...........')
crawler_ins2 = OpendataCrawler(url, failed_lst, institution_btn , directory , by_institution)
crawler_ins2.crawl_data()

# csv 파일 생성
csv_1 = Generate_csv( count, directory , by_institution)
csv_1.data_lst()

## count 초기화
count = 0

## crawler_loc1 : 요양기관소재지별 데이터 크롤링
print ('crawling Data ..............')
crawler_loc1 = OpendataCrawler(url, mdfeeCd_lst, location_btn, directory , by_location)
crawler_loc1.crawl_data()

#크롤링 실패한 리스트 가지고 다시 객체 생성 후 데이터 크롤링
print ('failed_lst = ',len(failed_lst))
print ('crawling failed_lst...........')
crawler_loc2 = OpendataCrawler(url, failed_lst, institution_btn , directory , by_location)
crawler_loc2.crawl_data()

# csv 파일 생성
csv_2 = Generate_csv( count, directory , by_location)
csv_2.data_lst()
