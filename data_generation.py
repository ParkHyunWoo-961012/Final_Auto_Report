from selenium import webdriver  #셀레늄 : 실제 웹브라우저를 열어서 진행
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import datetime
import pandas as pd
import time
from selenium.webdriver.common.by import By

def data_regeneration():
    def korea_investment_ELS():
        url = "https://securities.koreainvestment.com/main/mall/openels/EdlsInfo.jsp"
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.implicitly_wait(30)

        driver.get(url)
        bsObject = BeautifulSoup(driver.page_source, "html.parser")

        target= bsObject.findAll('table')[4]

        target_ELS_list = []
        ELS_composite_list= []
        ELS_structure_list = []
        ELS_return = []
        ELS_loss = []
        interval_list = []

        for i in target.findAll('td',{'class':"letter_0"}):
            interval_list.append(i.text.replace("\n","").replace("\t",""))
        for idx,rate in enumerate(bsObject.findAll('span',{'class':"product_listUpPoint"})[1:]):
            if idx%2==0:
                ELS_return.append(rate.text)
            else:
                ELS_loss.append(rate.text)
        for idx,i in enumerate(target.findAll("td",{'class':"t_left"})):
            if (idx+1)%3 == 1:
                target_ELS_list.append(i.text)
            if (idx+1)%3 == 2:
                ELS_composite_list.append(i.text)
            if (idx+1)%3 == 0:
                ELS_structure_list.append(i.text)
        target_ELS_list = []
        ELS_composite_list= []
        ELS_structure_list = []
        ELS_return = []
        ELS_loss = []
        interval_list = []

        for i in target.findAll('td',{'class':"letter_0"}):
            interval_list.append(i.text.replace("\n","").replace("\t",""))
        for idx,rate in enumerate(bsObject.findAll('span',{'class':"product_listUpPoint"})[1:]):
            if idx%2==0:
                ELS_return.append(rate.text)
            else:
                ELS_loss.append(rate.text)
        for idx,i in enumerate(target.findAll("td",{'class':"t_left"})):
            if (idx+1)%3 == 1:
                target_ELS_list.append(i.text)
            if (idx+1)%3 == 2:
                ELS_composite_list.append(i.text)
            if (idx+1)%3 == 0:
                ELS_structure_list.append(i.text)

        columns = ["ELS명","기초자산","구조","수익률","최대손실률","청약기간"]
        ELS_df = pd.DataFrame(columns = columns)

        interval_list = interval_list[1:]
        ELS_composite_list = ELS_composite_list[1:]
        target_ELS_list = target_ELS_list[1:]
        ELS_structure_list= ELS_structure_list[1:]

        for i in range(len(target_ELS_list)):
            input_item = {}
            input_item[columns[0]] = target_ELS_list[i].replace("\t","").replace("\n","").replace("비교담기 선택","").rstrip()
            input_item[columns[1]] = ELS_composite_list[i].replace("\n"," ").rstrip().lstrip()
            input_item[columns[2]] = ELS_structure_list[i]
            input_item[columns[3]] = ELS_return[i]
            input_item[columns[4]] = ELS_loss[i]
            input_item[columns[5]] = interval_list[i]
            ELS_df = ELS_df.append(input_item,ignore_index = True)

        url = "https://securities.koreainvestment.com/main/mall/openels/EdlsInfo.jsp"
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.implicitly_wait(30)

        driver.get(url)
        driver.find_element(By.XPATH,'//*[@id="list_style"]/table/thead/tr/th[5]/span/a[2]').click()
        time.sleep(3)

        bsObject = BeautifulSoup(driver.page_source, "html.parser")

        target_ELS_list = []
        ELS_composite_list= []
        ELS_structure_list = []
        ELS_return = []
        ELS_loss = []
        interval_list = []

        target= bsObject.findAll('table')[4]

        for i in target.findAll('td',{'class':"letter_0"}):
            interval_list.append(i.text.replace("\n","").replace("\t",""))
        for idx,rate in enumerate(bsObject.findAll('span',{'class':"product_listUpPoint"})[1:]):
            if idx%2==0:
                ELS_return.append(rate.text)
            else:
                ELS_loss.append(rate.text)
        for idx,i in enumerate(target.findAll("td",{'class':"t_left"})):
            if (idx+1)%3 == 1:
                target_ELS_list.append(i.text)
            if (idx+1)%3 == 2:
                ELS_composite_list.append(i.text)
            if (idx+1)%3 == 0:
                ELS_structure_list.append(i.text)

        for i in range(len(target_ELS_list)):
            input_item = {}
            input_item[columns[0]] = target_ELS_list[i].replace("\t","").replace("\n","").replace("비교담기 선택","").rstrip()
            input_item[columns[1]] = ELS_composite_list[i].replace("\n"," ").rstrip().lstrip()
            input_item[columns[2]] = ELS_structure_list[i]
            input_item[columns[3]] = ELS_return[i]
            input_item[columns[4]] = ELS_loss[i]
            input_item[columns[5]] = interval_list[i]
            ELS_df = ELS_df.append(input_item,ignore_index = True)

        ELS_df['발행사'] = "한국투자증권"
        return ELS_df


    def mirae_asset_ELS():
        url = "https://securities.miraeasset.com/mw/mks/mks4022/r01.do"
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.implicitly_wait(30)

        driver.get(url)
        bsObject = BeautifulSoup(driver.page_source, "html.parser")

        ELS_name = []
        ELS_composite = []
        ELS_structure = []
        ELS_loss = []
        ELS_return = []
        ELS_date = []

        for i in bsObject.findAll("p",{"class" : "tit"}):
            ELS_name.append(i.text)

        for i in bsObject.findAll("div",{"class" : "name"}):
            ELS_composite.append(i.span.text)
            ELS_structure.append(str(i).split("<span>")[2].split("</span>")[0])

        for idx,i in enumerate(bsObject.findAll("em",{"class" : "percent"})):
            if (idx+1)%2 == 1:
                ELS_loss.append(i.text.replace("\n","").replace("\t",""))
            else:
                ELS_return.append(i.text.replace("\n","").replace("\t",""))


        for i in bsObject.findAll("em",{"class" : "date_right"}):
            target = i.text.split("D-")
            if len(target)==1:
                ELS_date.append(datetime.datetime.now().strftime('%Y-%m-%d'))
            else:
                ELS_date.append((datetime.datetime.today()+datetime.timedelta(days=int(target[1]))).strftime("%Y.%m.%d"))


        mirae_asset_ELS = pd.DataFrame()
        mirae_asset_ELS['ELS명'] = ELS_name
        mirae_asset_ELS['기초자산'] = ELS_composite
        mirae_asset_ELS['구조'] = ELS_structure
        mirae_asset_ELS['수익률'] = ELS_return
        mirae_asset_ELS['최대손실률'] = ELS_loss
        mirae_asset_ELS['청약기간'] = ELS_date
        mirae_asset_ELS['발행사'] = "미래에셋증권"
        mirae_asset_ELS.sort_values("수익률")

        return mirae_asset_ELS

    def nh_ELS():
        url = "https://m.nhqv.com/finance/els/els/saleList"
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.implicitly_wait(30)

        driver.get(url)
        bsObject = BeautifulSoup(driver.page_source, "html.parser")
        bsObject.findAll("span",{"class":"txt"})
        bsObject = BeautifulSoup(driver.page_source, "html.parser")

        ELS_name = []
        ELS_loss = []
        ELS_return = []
        ELS_composite = []
        ELS_date = []
        ELS_structure =[]

        ## ELS명
        for i in bsObject.findAll("strong",{"class":"tit h37 disp_b"}):
            ELS_name.append(i.text)

        ## ELS 수익률 손실률
        for idx,i in enumerate(bsObject.findAll("span",{"class" : "info_stx2_1"})):
            if (idx+1)%2 == 1:
                ELS_return.append(i.text)
            else:
                ELS_loss.append(i.text)

        ## ELS 기초자산,청약기간,-,-,구조,-
        for idx,i in enumerate(bsObject.findAll("span",{"class":"txt"})):
            if (idx+1)%6==1:
                ELS_composite.append(i.text)
            if (idx+1)%6==2:
                ELS_date.append(i.text)
            if (idx+1)%6==5:
                ELS_structure.append(i.text)

        NH_els_df = pd.DataFrame()
        NH_els_df['ELS명'] = ELS_name
        NH_els_df['최대손실률'] = ELS_loss
        NH_els_df['수익률'] = ELS_return
        NH_els_df['기초자산'] = ELS_composite[:-1]
        NH_els_df['청약기간'] = ELS_date[:-1]
        NH_els_df['구조'] = ELS_structure
        NH_els_df['발행사'] = "NH투자증권"

        return NH_els_df


    def kiwoom_ELS():
        url = "https://www1.kiwoom.com/wm/edl/es010/edlElsView?dummyVal=0"
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.implicitly_wait(30)

        driver.get(url)
        bsObject = BeautifulSoup(driver.page_source, "html.parser")

        ELS_name = []
        ELS_composite = []
        ELS_return = []
        ELS_loss = []
        ELS_structure = []
        ELS_date =[]

        for idx,i in enumerate(bsObject.find("table").findAll("td")):
            if (idx+1)%10 == 1:
                ELS_name.append(i.text.replace("\n","").replace("\t","").split("제")[1])
            if (idx+1)%10 == 3:
                ELS_composite.append(i.text.replace("\n","").replace("\t",""))
            if (idx+1)%10 == 4:
                ELS_return.append(i.text.replace("\n","").replace("\t",""))
            if (idx+1)%10 == 5:
                ELS_loss.append(i.text.replace("\n","").replace("\t",""))
            if (idx+1)%10 == 6:
                ELS_structure.append(i.text.replace("\n","").replace("\t","").split("(")[1].split(",")[0])
            if (idx+1)%10 == 8:
                ELS_date.append(i.text.replace("\n","").replace("\t","").split("(")[1].split(")")[0])

        kiwoom_ELS = pd.DataFrame()
        kiwoom_ELS['ELS명'] = ELS_name
        kiwoom_ELS['기초자산'] = ELS_composite
        kiwoom_ELS['수익률'] = ELS_return
        kiwoom_ELS['최대손실률'] = ELS_loss
        kiwoom_ELS['구조'] = ELS_structure
        kiwoom_ELS['청약기간'] = ELS_date
        kiwoom_ELS['발행사'] = "키움증권"

        return kiwoom_ELS

    def korea_investment_bond():
        url = "https://www.truefriend.com/main/mall/opendecision/DecisionInfo.jsp?cmd=TF02da010100"
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.implicitly_wait(30)

        driver.get(url)
        bsObject = BeautifulSoup(driver.page_source, "html.parser")

        bond_class = []
        bond_name = []
        credit_list = []
        risk_list = []
        maturity_list = []
        remain_money = []
        표면이율 = []
        매수단가 = []
        세전수익률 = []
        잔존기간 = []
        매수수익률 = []
        세후수익률 = []

        for i in bsObject.findAll('td',{'headers':'tb-c-1-2'}):
            target_str = i.text.replace("\n","").replace("\t","").rstrip().lstrip()
            bond_name.append(target_str.split("신용등급")[0])
            credit_list.append(target_str.split("신용등급 :")[1].split("채권투자분석")[0].lstrip())

        for i in bsObject.findAll('td',{'headers':'tb-c-1-1'}):
            bond_class.append(i.text.replace("\n","").replace("\t","").rstrip().lstrip())

        for i in bsObject.findAll('td',{'headers':'tb-c-1-3'}):
            risk_list.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-1-4'}):
            maturity_list.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-1-5'}):
            remain_money.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-1-6'}):
            표면이율.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-1-7'}):
            매수단가.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-1-8'}):
            세전수익률.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-2-1'}):
            잔존기간.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-2-2'}):
            매수수익률.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-2-3'}):
            세후수익률.append(i.text)

        bond_df = pd.DataFrame()
        bond_df['채권이름'] = bond_name
        bond_df['채권종류'] = bond_class
        bond_df['신용등급'] = credit_list
        bond_df['위험도'] = risk_list
        bond_df['만기일'] = maturity_list
        bond_df['잔존수량(원)'] = remain_money
        bond_df['표면이율'] = 표면이율
        bond_df['매수단가'] = 매수단가
        bond_df['세전수익률'] = 세전수익률
        bond_df['잔존기간'] = 잔존기간
        bond_df['매수수익률'] = 매수수익률
        bond_df['세후수익률'] = 세후수익률

        return bond_df

    def kiwoom_bond():
        url = "https://www.kiwoom.com/wm/bnd/od010/bndOdListView"
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.implicitly_wait(30)


        driver.get(url)
        bsObject = BeautifulSoup(driver.page_source, "html.parser")
        driver.find_element(By.XPATH,'/html/body/main/section/div/div/div[3]/div[1]/ul/li[1]/a').click()
        time.sleep(3)

        bsObject = BeautifulSoup(driver.page_source, "html.parser")
        bond_name = []
        bond_before_tax_return = []
        bond_after_tax_return = []
        bond_expire_date= []
        bond_credit = []

        for i in bsObject.findAll("div",{"class": "fund-title-area"}):
            bond_name.append(i.text)

        for i in bsObject.findAll("td",{"name":"trdeGrt"}):
            bond_before_tax_return.append(i.text)

        for i in bsObject.findAll("td",{"name":"aftGrt"}):
            bond_after_tax_return.append(i.text)

        for i in bsObject.findAll("td",{"name":"exprDt"}):
            bond_expire_date.append(i.text)

        for i in bsObject.findAll("td",{"name":"crdtGrde"}):
            bond_credit.append(i.text)

        kiwoom_bond_df = pd.DataFrame()

        kiwoom_bond_df['채권이름'] = bond_name
        kiwoom_bond_df['세전수익률'] = bond_before_tax_return
        kiwoom_bond_df['세후수익률'] = bond_after_tax_return
        kiwoom_bond_df['만기일'] = bond_expire_date
        kiwoom_bond_df['신용등급'] = bond_credit

        return kiwoom_bond_df


    nh_df = nh_ELS()
    while len(nh_df)==0:
        nh_df = nh_ELS()

    korea_df = korea_investment_ELS()

    nh_df['기초자산'] = nh_df['기초자산'].str.replace(",",", ")
    korea_df['기초자산'] = korea_df['기초자산'].str.rstrip()
    korea_df['기초자산'] = korea_df['기초자산'].str.replace("\t","")
    korea_df['기초자산'] = korea_df['기초자산'].str.replace(" ",", ")


    target_df = kiwoom_ELS()
    target_df = pd.concat([target_df,nh_df])
    target_df = pd.concat([target_df,mirae_asset_ELS()])
    target_df = pd.concat([target_df,korea_df])
    target_df.reset_index(drop=True,inplace=True)

    target_df['기초자산'] = target_df['기초자산'].str.upper()
    target_df['기초자산'] = target_df['기초자산'].str.replace("INC.","")
    target_df['기초자산'] = target_df['기초자산'].str.rstrip()

    target_df['기초자산'] = target_df['기초자산'].str.replace("TESLA","테슬라")
    target_df['기초자산'] = target_df['기초자산'].str.replace("Inc.","")

    target_df['수익률'] = target_df['수익률'].str.replace("최대","")
    target_df['수익률'] = target_df['수익률'].str.replace("연","")
    target_df['수익률'] = target_df['수익률'].str.replace("%","")
    target_df['수익률'] = target_df['수익률'].str.strip()
    target_df['수익률'] =target_df['수익률'].astype(float)

    target_df['최대손실률'] = target_df['최대손실률'].str.replace('최대손실률',"")
    target_df['최대손실률'] = target_df['최대손실률'].str.strip()


    korea_investment_bond_df = korea_investment_bond()
    kiwoom_bond_df = kiwoom_bond()

    def knock_in_show(x):
        if len(x.split("(종가)")) !=1:
            return x.split("(종가)")[0][-2:]
        elif "(No" in x:
            return x.split("(No")[0][-2:]
        elif "NO" in x:
            return x.split(")")[0][-2:]
        elif "No" in x:
            return x.split(",")[0][-2:]

        elif len(x.split("KI")) ==2:
            if len(x.split("KI")[1])==0:
                return x.split("KI")[0][-2:]
            else:
                return x.split("KI")[1][1:3]
        elif len(x.split("KI")) !=1:
            return x.split("KI")[1][:3]
        elif len(x.split("/")) == 2:
            return x.split("/")[1][:2]
        else:
            return x

    target_df['낙인'] = target_df['구조'].apply(lambda x : knock_in_show(x))
    korea_investment_bond_df = korea_investment_bond_df[korea_investment_bond_df['채권종류']=="회사채"]
    kiwoom_bond_df['발행사'] = "키움증권"
    korea_investment_bond_df['발행사'] = "한국투자증권"

    target_bond = pd.concat([kiwoom_bond_df,korea_investment_bond_df[kiwoom_bond_df.columns]])
    target_bond['세후수익률'] = target_bond['세후수익률'].astype(float)

    target_col = list(target_df.columns)
    target_col.remove("발행사")
    target_df = target_df[['발행사']+target_col]

    target_df['ELS명'] = target_df['ELS명'].str.replace("원금비보장종목형","")
    target_df['ELS명'] = target_df['ELS명'].str.replace("원금비보장지수형","")
    target_df['ELS명'] = target_df['ELS명'].str.replace("파생결합증권","")
    def els_name_preprocessing(x):
        target = x.split("(")
        if "ELS" in x:
            return x
        if len(target)!=1:
            return x.split("(")[0]
        else:
            return x
    target_df['ELS명'] = target_df['ELS명'].apply(lambda x : els_name_preprocessing(x))

    unique_item = []
    for i in target_df['기초자산'].str.split(", ").values:
        for x in i:
            if x == "":
                pass
            else:
                unique_item.append(x)
    unique_item = list(set(unique_item))
    remove_item = ["HSCEI","S&P500","KOSPI200","NIKKEI225","EUROSTOXX50"]

    for i in remove_item:
        unique_item.remove(i)

    def base_composite_preprocessing(x,unique):
        for i in unique:
            if i in x:
                return "종목형"
        return "지수형"

    target_df['타입'] = target_df['기초자산'].apply(lambda x : base_composite_preprocessing(x,unique_item))

    target_df.to_csv("./data/ELS모음.csv",index=False)
    target_bond.to_csv("./data/채권모음.csv",index=False)

data_regeneration()

