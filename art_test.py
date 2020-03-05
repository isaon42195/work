#!/usr/bin/env python
# coding: utf-8

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By    

import datetime
import numpy as np
import pandas as pd
from pandas import DataFrame, Series
import openpyxl
import glob
import os
import sys
import configparser
import logging
import codecs
import bs4 as bs
from bs4 import BeautifulSoup
import re
import csv

 
# ログイン処理
def login() :
    # driver = webdriver.Chrome()
    # IPSのホームを開く、ログイン画面にリダイレクトされる
    # driver.get("http://wptg.oraclecorp.com/pls/htmldb/f?p=101:1:2256983443082162")
    driver.implicitly_wait(10)
    driver.get("https://ips.oraclecorp.com/ords/f?p=108:3:12053727619584::NO:3::")
    wait = WebDriverWait(driver, 60)
    title = wait.until(expected_conditions.title_contains("Query All Projects"))
    print(title)
    #driver.find_element_by_id("sso_username").send_keys(sso_username)
    #sleep(15)
    #print(driver.find_element_by_id("ssopassword").is_displayed())
    #print(psd.text)
    # submit
    #driver.find_element_by_id("sso_username").send_keys(Keys.ENTER)

def login_t() :
    # driver = webdriver.Chrome()
    # IPSのホームを開く、ログイン画面にリダイレクトされる
    # driver.get("http://wptg.oraclecorp.com/pls/htmldb/f?p=101:1:2256983443082162")
    driver.implicitly_wait(10)
    driver.get("https://ips.oraclecorp.com/ords/f?p=108:3:12053727619584::NO:3::")
    #wait = WebDriverWait(driver, 60)
    driver.find_element_by_id("sso_username").send_keys("isao.nakamura@oracle.com")
    driver.find_element_by_id("ssopassword").send_keys("IS83on273")
    #sleep(15)
    #print(driver.find_element_by_id("ssopassword").is_displayed())
    #print(psd.text)
    # submit
    driver.find_element_by_id("sso_username").send_keys(Keys.ENTER)

    #title = wait.until(expected_conditions.title_contains("Query All Projects"))
    #print(title)

# ログイン処理　ユーザがuser/passwordを入力
def loginw() :
    # driver = webdriver.Chrome()
    # 　IPSのホームを開く、ログイン画面にリダイレクトされる
    # driver.get("http://wptg.oraclecorp.com/pls/htmldb/f?p=101:1:2256983443082162")
    driver.implicitly_wait(10)
    driver.get("https://ips.us.oracle.com/ords/f?p=101:1:16926983381439:::::")

# ファイルを削除
def rfile(path) :
    file_list = glob.glob(path)

    for file in file_list:
        print("remove：{0}".format(file))
        os.remove(file)

# スクリーンショット
def scshots(ponum, prd_name):
    filename = "/scshots/" + prd_name + "_" + ponum + '_' + timestamp + ".png"
    driver.get_screenshot_as_file(filename)

# Procurement でのreceiving
def receive(ponum, p_amount, prd, es_sow) :
    driver.refresh()
    # driver.implicitly_wait(50)
    proc = "https://eeho.fa.us2.oraclecloud.com/fscmUI/faces/FndOverview?fnd=%3B%3B%3B%3Bfalse%3B256%3B%3B%3B&fndGlobalItemNodeId=itemNode_my_information_self_service_receipts&_afrLoop=4739405550090837&_afrWindowMode=0&_afrWindowId=112bfxuov1&_adf.ctrl-state=bsi7c28t5_594&_afrFS=16&_afrMT=screen&_afrMFW=1361&_afrMFH=633&_afrMFDW=1366&_afrMFDH=768&_afrMFC=8&_afrMFCI=0&_afrMFM=0&_afrMFR=96&_afrMFG=0&_afrMFS=0&_afrMFO=0"
    driver.get(proc)
    driver.refresh()

    # POを検索、サブルーチン化する
    eby = driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value10::content")
    eby.clear()
    bu = driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value40::content")
    #bu.send_keys("ORCL IE-iProc")
    WebDriverWait(driver, 25).until(expected_conditions.presence_of_element_located((By.ID, '_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value40::content')))   

    select = Select(bu)
    select.select_by_visible_text("ORCL IE-iProc")

    bu = driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value30::content")
    WebDriverWait(driver, 25).until(expected_conditions.presence_of_element_located((By.ID, '_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value30::content')))   
    select = Select(bu)
    #select.clear()
    select.select_by_visible_text("Any time")

    # ここは変数
    po = driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value50::content")
    po.clear()
    po.send_keys(ponum)
    driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId::search").click()
    driver.find_element_by_xpath("//div[@id='_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:AT1:_ATp:QrRsId::db']/table/tbody/tr/td[2]/div/table/tbody/tr/td[5]").click()

    driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:AT1:_ATp:ReceiveItemsReceiveButtonId").click()
    
    # xpathでQuantity フィールドを指定（idは変化するため）
    element = driver.find_element_by_xpath("//span/span/input")
    element.click()
    element.clear()
    element.send_keys(p_amount)
    #element.send_keys(Keys.ENTER)
    
    driver.find_element_by_xpath("//span[contains(.,'Submit')]").click()
    
    #ここでrcp_rec()を呼び出し、Receipt#を記録
    rcp_rec(ponum, es_sow)
    
    driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:1:ap1:ReceiptConfirmationDialogId::_ttxt").click()
    
    #scshots(ponum, prd)
    
    

# Finapp 上でのreceiving
def finapp_rcv(url, r_amount, prd):
    proc_fin = "https://ips.us.oracle.com/ords/f?p=700:1:7797688992995:::::"
    proc_po = url

    driver.get(proc_fin)
    driver.get(proc_po)

#20200213S
    WebDriverWait(driver, 25).until(expected_conditions.presence_of_element_located((By.ID, 'P36_RECEIVABLE'))) 
    driver.find_element_by_id("P36_RECEIVABLE").send_keys(str(r_amount))
    driver.find_element_by_id("P36_RECEIVABLE").send_keys(str(Keys.ENTER))
    # scshots(ponum, prd)
    WebDriverWait(driver, 25).until(expected_conditions.presence_of_element_located((By.ID, 'B687103398259645923'))) 
    WebDriverWait(driver, 5)
    driver.find_element_by_id("B687103398259645923").click()

    #alert = driver.find_elements_by_class_name("t-Alert-title")
    #alert = driver.find_element_by_xpath("//*[@id='t_Alert_Success']/div/div[2]/div/h2")
    #print(alert)
    #WebDriverWait(driver, 25).until(expected_conditions.presence_of_element_located((By.CLASS_NAME, 't-Alert-title'))) 
    #driver.refresh()
#20200213E

# Finapp 上でのreceiving
def finapp_rcv_test(url, r_amount, prd):
    proc_fin = "https://ips.us.oracle.com/ords/f?p=700:1:7797688992995:::::"
    proc_po = url

    driver.get(proc_fin)
    driver.get(proc_po)

#20200213S
    WebDriverWait(driver, 25).until(expected_conditions.presence_of_element_located((By.ID, 'P36_RECEIVABLE'))) 
    driver.find_element_by_id("P36_RECEIVABLE").send_keys(str(r_amount))
    driver.find_element_by_id("P36_RECEIVABLE").send_keys(str(Keys.ENTER))
    
    # scshots(ponum, prd)
    #driver.refresh()
    WebDriverWait(driver, 25).until(expected_conditions.presence_of_element_located((By.ID, 'B687103398259645923'))) 
    #driver.find_element_by_id("B687103398259645923").send_keys(str(Keys.ENTER))
    driver.find_element_by_id("B687103398259645923").click() 
    driver.implicitly_wait(100) 

    driver.refresh()
    WebDriverWait(driver, 25).until(expected_conditions.presence_of_element_located((By.ID, 'P36_AMT_TO_RECEIVE_DISPLAY'))) 
    #if driver.find_element_by_xpath("/html/body/form/div[1]/div/div[2]/span[1]/div/div/div/div[2]/div/h2") :
    #    print('Good Job!')


# Finapp でPOとSOWのテーブルをそれぞれ作成し、csvに保存
def es_sow_f(proc_prd):
    rfile(downloads + "estimates*.csv")
    proc_fin = "https://ips.us.oracle.com/ords/f?p=700:1:7797688992995:::::"
    # proc_prd = "http://wptg.oraclecorp.com/pls/htmldb/f?p=700:11:6925941450150::NO::P0_PROJECT_ID,P0_SSOT_VERSION_ID,P0_BP_ID:47044,7993,237293"

    driver.get(proc_fin)
    driver.get(proc_prd)

    # driver.refresh()
    driver.find_element_by_id("estimates_report_actions_button").click()
    driver.find_element_by_id("estimates_report_actions_menu_14i").click()
    driver.find_element_by_id("estimates_report_download_CSV").click()
    driver.find_element_by_id("estimates_report_download_CSV").send_keys(Keys.ESCAPE)

    logger.log(20, 'Estimates csv download.')
    #print('Processing : Estimates csv download.')
    driver.refresh()

    rfile(downloads + "statements_of_work*.csv")
    driver.find_element_by_id("sow_report_actions_button").click()
    driver.find_element_by_id("sow_report_actions_menu_14i").click()
    driver.find_element_by_id("sow_report_download_CSV").click()
    driver.find_element_by_id("sow_report_download_CSV").send_keys(Keys.ESCAPE)
    logger.log(20, 'SOW cvs downloaded.')
    #print('Processing : SOW cvs download.')
    driver.refresh()

def es_sow_x(proc_prd):
    rfile(downloads + "estimates*.csv")
    proc_fin = "https://ips.us.oracle.com/ords/f?p=700:1:7797688992995:::::"
    # proc_prd = "http://wptg.oraclecorp.com/pls/htmldb/f?p=700:11:6925941450150::NO::P0_PROJECT_ID,P0_SSOT_VERSION_ID,P0_BP_ID:47044,7993,237293"

    driver.get(proc_fin)
    driver.get(proc_prd)

    logger.log(20, 'es_sow_x')
    driver.refresh()
    #driver.find_element_by_id("estimates_report_actions_button").click()
    driver.find_element_by_xpath("//div[@id='estimates_report_toolbar_controls']/div[3]/div[2]/button").click()
    #driver.find_element_by_id("estimates_report_actions_menu_14i").click()
    driver.find_element_by_xpath("//button[@id='estimates_report_actions_menu_14i']").click()
    #driver.find_element_by_id("estimates_report_download_CSV").click()
    driver.find_element_by_xpath("//a[@id='estimates_report_download_CSV']/span").click()
    #driver.find_element_by_id("estimates_report_download_CSV").send_keys(Keys.ESCAPE)
    
    logger.log(20, 'Estimates csv download.')
    #print('Processing : Estimates csv download.')
    driver.refresh()

    rfile(downloads + "statements_of_work*.csv")
    #driver.find_element_by_id("sow_report_actions_button").click()
    driver.find_element_by_xpath("//button[@id='sow_report_actions_button']").click()
    #driver.find_element_by_id("sow_report_actions_menu_14i").click()
    driver.find_element_by_xpath("//button[@id='sow_report_actions_menu_14i']").click()
    #driver.find_element_by_id("sow_report_download_CSV").click()
    driver.find_element_by_xpath("//a[@id='sow_report_download_CSV']/span").click()
    #driver.find_element_by_id("sow_report_download_CSV").send_keys(Keys.ESCAPE)

    logger.log(20, 'SOW cvs downloaded.')
    #print('Processing : SOW cvs download.')
    driver.refresh()


# Confirmation Dialogから Receipt#を取得して記録
def rcp_rec(ponum, es_sow):
    rcp_txt = driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:1:ap1:MsTxId").get_attribute("textContent")
    # rcp_txt = "You created the following receipt numbers: 23371."
    rcp_num = rcp_txt.replace("You created the following receipt numbers: ", "").replace(".", "")
    logger.log(20, 'Receipt# : ' + rcp_num)

    #es_sow['Recept#'] = "N/A"
    es_sow.loc[es_sow['PO Number'] == ponum, 'Receipt#'] = str(rcp_num)
    
# PO# の確認
def po_chk(ponum) :
    # driver.implicitly_wait(20)
    proc = "https://eeho.fa.us2.oraclecloud.com/fscmUI/faces/FndOverview?fnd=%3B%3B%3B%3Bfalse%3B256%3B%3B%3B&fndGlobalItemNodeId=itemNode_my_information_self_service_receipts&_afrLoop=4739405550090837&_afrWindowMode=0&_afrWindowId=112bfxuov1&_adf.ctrl-state=bsi7c28t5_594&_afrFS=16&_afrMT=screen&_afrMFW=1361&_afrMFH=633&_afrMFDW=1366&_afrMFDH=768&_afrMFC=8&_afrMFCI=0&_afrMFM=0&_afrMFR=96&_afrMFG=0&_afrMFS=0&_afrMFO=0"
    driver.get(proc)
    driver.refresh()  
    
    eby = driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value10::content")
    eby.clear()
    bu = driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value40::content")
    #bu.send_keys("ORCL IE-iProc")
    select = Select(bu)
    select.select_by_visible_text("ORCL IE-iProc")
    #driver.refresh()  

#20200213S
    WebDriverWait(driver, 25).until(expected_conditions.presence_of_element_located((By.ID, '_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value30::content')))   
    bu = driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value30::content")
#20200213E
    select = Select(bu)
    #select.clear()
    select.select_by_visible_text("Any time")

    po = driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId:value50::content")
    po.clear()
    po.send_keys(ponum)
    ## 77 2line
    target = driver.find_element_by_xpath("//*[@id='_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId::search']")
    driver.execute_script("arguments[0].click();", target)
    #driver.find_element_by_id("_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:QryId::search").click()
    
    driver.find_element_by_xpath("//div[@id='_FOpt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:ap1:AT1:_ATp:QrRsId::db']/table/tbody/tr/td[2]/div/table/tbody/tr/td[5]").click()


# Receiving の例外処理
def rcv_fail(isfin, ponum, es_sow, out_csv):
    #es_sow['Recept#'] = "N/A"
    if isfin == 'proc' :
        es_sow.loc[es_sow['PO Number'] == ponum, 'Receipt#'] = 'failed'
        #scshots(ponum, prd_name)
        es_sow.to_csv(out_csv)
    else :
        es_sow.loc[es_sow['PO Number'] == ponum, 'FinApp'] = 'failed'
        #scshots(ponum, prd_name)
        es_sow.to_csv(out_csv)


# レシービングを実行するための準備。FinAppでEstimateとSOWのテーブルをダウンロードし、マージする。
def create_table4_rcv() :
    try :
        es_sow_x(proc_prd)
        logger.log(20, 'CVS download done.')
        # print('Processing : CSV download done. ')
    except NoSuchElementException :
        logger.log(100, 'Failed to download csv file.')
        # print('Error : failed to dowload csv files.')
        sys.exit()

    # csvをDataFrameに読み込みマージ
    estimates = pd.read_csv(downloads + 'estimates.csv', thousands=',')
    estimates = estimates.drop(['Estimate Scenario', 'Vendor', 'Req Number', 'Status', 'Delta (USD)'], axis=1)
    estimates['Amt. Received (Local)']  = estimates['Amt. Received (Local)'] .replace('-', 0)
    # Rcvがない場合の対処
    estimates.to_csv(downloads + 'estimates.csv')
    estimates = pd.read_csv(downloads + 'estimates.csv', thousands=',')
        
    sow = pd.read_csv(downloads + "statements_of_work.csv", thousands=',', engine='python')

    sow = sow.drop(['Language', 'Vendor', 'Status', 'Appr. Value (USD)', 'Currency code', 'Extra comment', 'Req Number'], axis=1)
    es_sow = pd.merge(estimates, sow, on='PO Number')
    es_sow['to_Rcv'] = es_sow['Value (Local)_y'] - es_sow['Amt. Received (Local)']
    es_sow.to_csv(downloads + 'es_sow.csv')
    logger.log(20, 'es_sow merge done.')
    # print('Processing : es_sow merge done.')
    return es_sow
    

# PO# の存在、レシーブ可能かをチェック
def ponum_isexist(es_sow) :
    for ponum, amount in zip(es_sow['PO Number'], es_sow['to_Rcv']):
        if ('IE' in str(ponum) and amount > 0):
            try :
                po_chk(ponum)
                logger.log(20, ponum + ' exists')
                # print(ponum + ' exists')
            except NoSuchElementException :
                logger.log(100, ponum + ' does not exists or unreceivable.')
                es_sow.loc[es_sow['PO Number'] == ponum, 'Receipt#'] = 'failed'
                go_fa = False
                #print('Error : ' + ponum + ' do not exist or fully received.')
                #sys.exit()

#各estimate へのurlリストを作成
def get_url_table(proc_prd, prd_name):
    driver.get(proc_prd)
    html = driver.page_source
    columnname = ["URL", "Vendor", "Language", "Req Number", "Status", "PO Number", "Amt. Received (Local)", "Amt. Received (USD)", "Value (Local)", "Value(USD)", "Appr. Value", "Delta usd", "Currency Code"]

    soup = bs.BeautifulSoup(html, 'lxml')
    parsed_table = soup.find_all('table')[0] 
    data = [[td.a['href'] if td.find('a') else 
             ''.join(td.stripped_strings)
             for td in row.find_all('td')]
            for row in parsed_table.find_all('tr')]
    df = pd.DataFrame(data[1:], columns=columnname)
    df["URL"] = "https://ips.us.oracle.com/ords/" + df["URL"] 
    df.to_csv(prd_name + '.csv')  

# bs4 でテーブル作成
def get_es_sow_html(es_or_sow, proc_prd):
    proc_fin = "https://ips.us.oracle.com/ords/f?p=700:1:7797688992995:::::"
    driver.get(proc_fin)
    driver.get(proc_prd)
    html = driver.page_source
    html = html.replace('$', '').replace('€', '').replace('¥', '').replace(',', '')
    
    #with open('sss.html', 'w') as f:
    #    f.write(html)

    bsObj = BeautifulSoup(html, "html.parser")

    table = bsObj.findAll("table", {"aria-label": re.compile(es_or_sow)})[0]
    rows = table.findAll("tr")

    csvFile = open(es_or_sow + ".csv", 'wt', newline='', encoding='utf-8')
    writer = csv.writer(csvFile)

    try:
        for row in rows:
            csvRow = []
            for cell in row.findAll(['td', 'th']):
                csvRow.append(cell.get_text())
            writer.writerow(csvRow)
    finally:
        csvFile.close()

# bs4で作成したes_sow
def merge_es_sow_html(prd_name, proc_prd):
    get_es_sow_html('Estimates', proc_prd)
    get_es_sow_html('Statements of Work', proc_prd)
    es_df_html = pd.read_csv("Estimates.csv", header=1)
    es_df_html = es_df_html[es_df_html['StatusAscending'] == 'Approved']
    es_df_html = es_df_html.fillna(0)
    es_df_html = es_df_html[es_df_html['PO Number'].str.contains('IE', na=False)]
    es_df_html['Amt. Received (Local)'] = es_df_html['Amt. Received (Local)'].astype(float)
    #print(es_df_html.dtypes)
    #es_df_html.to_csv('OCI.csv')
    sow_df_html = pd.read_csv("Statements of Work.csv")
    sow_df_html = sow_df_html.fillna(0)
    es_sow = pd.merge(es_df_html, sow_df_html, on='PO Number')
    es_sow.to_csv(prd_name + '_es_sow.csv')
    #print(es_sow.dtypes)
    es_sow['to_Rcv'] = es_sow['Value (Local)_y'] - es_sow['Amt. Received (Local)']
    es_sow['to_Rcv'] = round(es_sow['to_Rcv'], 2)
    es_sow.to_csv(prd_name + '_es_sow.csv')
    #print(es_sow.dtypes)
    if (es_sow.duplicated(subset='PO Number').any()):
        logger.log(10, 'Error : There are duplicated PO#s in FinApp')
        sys.exit()
        
    return es_sow

# FinApp のrcvをチェック
def fin_rcv_chk(out_csv, proc_prd):
    es_sow = pd.read_csv(out_csv)
    get_es_sow_html('Estimates', proc_prd)
    es_df_html = pd.read_csv("Estimates.csv", header=1)
    es_df_html = es_df_html[es_df_html['StatusAscending'] == 'Approved']
    es_df_html = es_df_html.fillna(0)
    es_df_html = es_df_html[es_df_html['PO Number'].str.contains('IE', na=False)]
    es_df_html['Amt. Received (Local)'] = es_df_html['Amt. Received (Local)'].astype(float)
    es_sow_new = pd.merge(es_sow, es_df_html, on='PO Number')
    # es_sow_new = es_sow_new.drop(['Unnamed: 0', 'Link_x', 'StatusAscending_x', 'Delta (USD)_x', 'Link_y', 'LanguageAscending', 'Vendor_y', 'Status', 'Value (USD)_y', 'Appr. Value (USD)_y', 'Currency code_y', 'Value (Local)_y', 'Extra comment', 'Req Number_y', 'Link', 'Vendor', 'Language_y', 'Req Number', 'StatusAscending_y', 'Value (Local)', 'Value (USD)', 'Appr. Value (USD)', 'Delta (USD)_y', 'Currency code'],  axis=1)
    es_sow_new['Diff'] = es_sow_new['Amt. Received (Local)_y'] - es_sow_new['Amt. Received (Local)_x'] - es_sow_new['to_Rcv']
    es_sow['Diff'] = es_sow_new['Diff'].astype(int)
    es_sow['Rate'] =es_sow['Value (Local)_x'] / es_sow['Value (USD)_x']
    es_sow['to_Rcv(USD)'] = es_sow['to_Rcv'] / es_sow['Rate'] 
    es_sow['Receipt#'] = es_sow['Receipt#'].astype(object)
    es_sow.loc[es_sow['Receipt#'] == 'failed', 'to_Rcv(USD)'] = 0
    for ponum, fail in zip(es_sow['PO Number'], es_sow['Diff']) :
        if (fail > 0):
            logger.log(10, 'Receiving in FinApp might fail : ' + ponum)

    total_r = es_sow['to_Rcv(USD)'].sum()
    #logger.log(20, 'Total receiving amount : $' + str(total_r))
    es_sow.to_csv(out_csv)


# 設定ファイル読み込み。

def config_ini():
    inifile = configparser.ConfigParser()
    inifile.read('config.ini')
    try :
        #sso_username = inifile.get('sso', 'username')
        #ssopassword = inifile.get('sso', 'password')
        downloads = inifile.get('folder', 'downloads')
        prd_name = inifile.get('product', 'name')
        proc_prd = inifile.get('product', 'finapp_url')
        logger.log(20, 'config.ini done. ART Version 1.6')
        #   print('Processing : config.ini done.')
    except configparser.NoOptionError :
        logger.log(100, 'Error in config.ini.')
        #print('Error in config.ini')
        sys.exit()
    return(prd_name, proc_prd)

def config_all():
        finapp_url = pd.read_csv('finapp_url.csv', thousands=',')
        for name, url in zip(finapp_url['name'], finapp_url['url']):
            prd_name, proc_prd = config(name, url)
            rcv_all(prd_name, proc_prd)

def config(name, url):
        prd_name = name
        proc_prd = url
        return (prd_name, proc_prd)

def rcv_all(prd_name, proc_prd):
	# Step 2	
	# Estimate, SOW をダウンロード	
	# es_sow = create_table4_rcv()	
	logger.log(20, 'Receiving ' + prd_name)
	es_sow = merge_es_sow_html(prd_name, proc_prd)	
		
	# URL リストを作成	
	get_url_table(proc_prd, prd_name)	
		
	go_fa = True	
		
	#  Step 3:	
	# PO# がレシーブ可能かチェック	
	# テスト用データ！！	
	#es_sow = pd.read_csv('test_es_sow.csv', thousands=',')	
		
	timestamp = datetime.date.today().isoformat()	
	out_csv = prd_name + '_' + timestamp + '.csv'	
	ponum_isexist(es_sow)	
		
		
	# Step 4	
	# レシービングを実行（Pocurement, FinApp)	
	#finpo_df = pd.read_csv(downloads + 'finapp_po_url.csv', encoding='cp932')	
	finpo_df = pd.read_csv(prd_name + '.csv')	
	fct_df = pd.merge(es_sow, finpo_df, on='PO Number')	
	fct_df.to_csv('url.csv')	
	logger.log(20, 'Line#455')	
	 	
	for ponum, amount, url in zip(fct_df['PO Number'], fct_df['to_Rcv'], fct_df['URL']):	
	    #logger.log(20, 'PO amount is '+str(amount))	
	    if ('IE' in str(ponum) and amount > 0):	
	        try :	
	            #p_amount = str(amount)	
	            p_amount = str(amount).replace(',', '')	
	            logger.log(20, 'Processing : ' + ponum + ' : ' + p_amount)	
	            # print('processing : ' + ponum )	
	            receive(ponum, p_amount, prd_name, es_sow)	
	            logger.log(20, 'Processing : ' + ponum + ' received in Procurement : ' + p_amount)	
	            try :	
	                if (go_fa == True):	
	                    finapp_rcv(url, p_amount, proc_prd)	
	                    logger.log(20, 'Processing : ' + ponum + ' received in FinApp : ' + p_amount)	
	                    loginw()	
	            except NoSuchElementException :	
	                logger.log(100, ponum + ' : receiving failed on FinApp : '  + p_amount)	
	                rcv_fail('finapp', ponum, es_sow, out_csv)	
	            	
	        except NoSuchElementException :	
	            logger.log(100, 'Error : ' + ponum + ' : receiving failed on Procurement : ' + p_amount)	
	            rcv_fail('proc', ponum, es_sow, out_csv)	
	        	
		
	es_sow = es_sow.drop(['Link_x', 'Link_y', 'LanguageAscending', 'Vendor_y', 'Status', 'Value (USD)_y', 'Appr. Value (USD)_y', 'Currency code_y', 'Extra comment', 'Req Number_y'], axis=1)	
	if 'Receipt#' in es_sow.columns: 
	    es_sow = es_sow.dropna(subset=['Receipt#'])	
		
	es_sow.to_csv(out_csv)	
	if 'Receipt#' in es_sow.columns: 	
	    fin_rcv_chk(out_csv, proc_prd)	
	es_sow = pd.read_csv(out_csv)
	if 'to_Rcv(USD)' in es_sow.columns:
		total_r = es_sow['to_Rcv(USD)'].sum()
		logger.log(20, 'Total receiving amount : $' + str(total_r))
	logger.log(20, 'Completed : Receipt number is available : ' + out_csv )	
	#driver.close()	
		

###### Main Step1 ############

# ログの出力名を設定
logger = logging.getLogger('LoggingTest')
# ログレベルの設定
logger.setLevel(10)
# ログのファイル出力先を設定
timestamp = datetime.date.today().isoformat()
fh = logging.FileHandler('auto_rcv_' + timestamp + '.log')
logger.addHandler(fh)
 
# ログのコンソール出力の設定
sh = logging.StreamHandler()
logger.addHandler(sh)
 
# ログの出力形式の設定
formatter = logging.Formatter('%(asctime)s:%(lineno)d:%(levelname)s:%(message)s')
fh.setFormatter(formatter)
sh.setFormatter(formatter)

logger.log(20, 'ART Version 1.83')
logger.log(20, 'Starting Google Chrome. SSO log-on page will appear ...')

options = ChromeOptions()

# ヘッドレスモードを有効にする（次の行をコメントアウトすると画面が表示される）。
#options.add_argument('--headless')
#options.add_argument('--disable-gpu')

# ChromeのWebDriverオブジェクトを作成する。
driver = Chrome(options=options)

login_t()

url = "https://ips.us.oracle.com/ords/f?p=700:36:4564262026917::NO:RP:P36_EST_ID,P0_PROJECT_ID,P36_RECEIVABLE:333236,47788"
r_amount = 0.36
prd = "APEX"

finapp_rcv_test(url, r_amount, prd)
#driver.get(url)

#WebDriverWait(driver, 20)

# ログイン
# driver = webdriver.Chrome()
#login()
#logger.log(20, 'Login done.')
#logger.log(20, 'Loading finapp_url.csv ...')
#config_all()
#logger.log(20, 'Receiving Completed.')
#driver.close()