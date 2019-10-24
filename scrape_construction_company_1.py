import numpy
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
   # ページを開く~
   driver = webdriver.Chrome()
   driver.get(url)

   # 県メニューを検索
   kencode = driver.find_element_by_name('kenCode')

   # 県メニューで東京を入力
   kencode_select_element = Select(kencode)
   kencode_select_element.select_by_value('13')

   # # 業種指定
   # gyoshu = driver.find_element_by_id('gyosyu')
   # gyoshu_select_element = Select(gyoshu)
   # gyoshu_select_element.select_by_value('2')
   #
   # # 業種種類指定
   # gyoshuType = driver.find_element_by_id('gyosyuType')
   # gyoshuType_select_element = Select(gyoshuType)
   # gyoshuType_select_element.select_by_value('1')

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

   error = 0
   # テーブルの情報を取得
   for n in range(page_list_number):
   # for n in range(1):

       result_table = driver.find_element_by_class_name("re_disp")
       tbody = result_table.find_element_by_tag_name("tbody")
       trs = tbody.find_elements_by_tag_name("tr")

       for i in range(1,len(trs)):
            try:
                result_table = driver.find_element_by_class_name("re_disp")
                tbody = driver.find_element_by_xpath("//*[@id='container_cont']/table/tbody")
                trs = tbody.find_elements_by_tag_name("tr")
                tds = trs[i].find_elements_by_tag_name("td")

                for td in tds:
                    print(td.text, end="　")
                    data_row.append(td.text)

                try:
                    company_name = tds[3].find_element_by_tag_name("a")
                    company_name.click()

                    valid_period = driver.find_element_by_class_name("re_summ_5").find_element_by_tag_name("tbody").find_elements_by_tag_name("tr")[0].find_elements_by_tag_name("td")[0]
                    # principle_name = driver.find_element_by_class_name("re_summ").find_element_by_tag_name("tbody").find_elements_by_tag_name("tr")[2].find_element_by_tag_name("td")
                    phone_number = driver.find_element_by_class_name("re_summ").find_element_by_tag_name("tbody").find_elements_by_tag_name("tr")[4].find_element_by_tag_name("td")


                    print(phone_number.text, end=" ")
                    print(valid_period.text)

                    data_row.append(phone_number.text)
                    data_row.append(valid_period.text)
                    driver.back()

                except:
                    error += 1
                    time.sleep(5)
                    if error >= 10:
                        break
                    print("error :" + str(error))
                    pass

                data_all.append(data_row)
                data_row = []
            except:
                error += 1
                time.sleep(5)
                print("error :" + str(error))
                pass


       print("-------------------------------------------------")


       if n == 0:
           page_list = driver.find_element_by_id('pageListNo1')
           page_list_select_element = Select(page_list)
           page_list_select_element.select_by_value("2")
           print("合計error数 :" + str(error))
       else:
           time.sleep(5)
           page_list = driver.find_element_by_id('pageListNo1')
           page_list_select_element = Select(page_list)
           page_list_select_element.select_by_value(str(n + 1))
           print("合計error数 :" + str(error))


   df = pd.DataFrame(data_all, index=list(range(1,len(data_all) + 1)), columns=["No.","許可行政庁","許可番号","商号又は名称","代表者名","営業所名","所在地","電話番号","許可の有効期間"])
   print(df)
   df.to_excel('/Users/nakazawatoga/Desktop/scrape_construction_company/test1.xlsx', sheet_name='test1', index = False)
