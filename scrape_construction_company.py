from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import urllib
import chromedriver_binary
import time
import pandas as pd
import openpyxl

url = "https://etsuran.mlit.go.jp/TAKKEN/kensetuKensaku.do"
if __name__ == '__main__':
   # ページを開く
   driver = webdriver.Chrome()
   driver.get(url)

   # 県メニューを検索
   kencode = driver.find_element_by_name('kenCode')

   # 県メニューで東京を入力
   kencode_select_element = Select(kencode)
   kencode_select_element.select_by_value('13')

   # 業種指定
   gyoshu = driver.find_element_by_id('gyosyu')
   gyoshu_select_element = Select(gyoshu)
   gyoshu_select_element.select_by_value('2')

   # 業種種類指定
   gyoshuType = driver.find_element_by_id('gyosyuType')
   gyoshuType_select_element = Select(gyoshuType)
   gyoshuType_select_element.select_by_value('1')

   # 検索結果の数を50に変更
   disp_count = driver.find_element_by_id('dispCount')
   disp_count_select_element = Select(disp_count)
   disp_count_select_element.select_by_value('50')
   # 検索ボタンをクリック
   driver.find_element_by_xpath('//*[@id="input"]/div[6]/div[5]/img').click()

   # 検索画面の個数をカウント
   page_list = driver.find_element_by_id('pageListNo1')
   page_list_number = len(page_list.find_elements_by_tag_name("option"))
   print("検索ページ : " + str(page_list_number))

   data_row = []
   data_all = []

   # テーブルの情報を取得
   # for n in range(page_list_number):
   for n in range(2):
       result_table = driver.find_elements_by_class_name("re_disp")
       tbody = driver.find_element_by_xpath("//*[@id='container_cont']/table/tbody")
       trs = tbody.find_elements_by_tag_name("tr")

       for i in range(1,len(trs)):
            tds = trs[i].find_elements_by_tag_name("td")
            for td in tds:
                print(td.text, end="　")
                data_row.append(td.text)
            print(i)
            data_all.append(data_row)
            data_row = []
       print("-------------------------------------------------")
       time.sleep(1)

       if n == 0:
           page_list = driver.find_element_by_id('pageListNo1')
           page_list_select_element = Select(page_list)
           page_list_select_element.select_by_value("2")
       else:
           page_list = driver.find_element_by_id('pageListNo1')
           page_list_select_element = Select(page_list)
           page_list_select_element.select_by_value(str(n + 1))

   df = pd.DataFrame(data_all, index=list(range(1,len(data_all) + 1)), columns=["No.","許可行政庁","許可番号","商号又は名称","代表者名","営業所名","所在地"])
   print(df)
   df.to_excel('/Users/Common/Desktop/scraping/test1.xlsx', sheet_name='test1', index = False)
