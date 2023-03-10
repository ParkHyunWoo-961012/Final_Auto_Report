from selenium import webdriver
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

        columns = ["ELS???","????????????","??????","?????????","???????????????","????????????"]
        ELS_df = pd.DataFrame(columns = columns)

        interval_list = interval_list[1:]
        ELS_composite_list = ELS_composite_list[1:]
        target_ELS_list = target_ELS_list[1:]
        ELS_structure_list= ELS_structure_list[1:]

        for i in range(len(target_ELS_list)):
            input_item = {}
            input_item[columns[0]] = target_ELS_list[i].replace("\t","").replace("\n","").replace("???????????? ??????","").rstrip()
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
            input_item[columns[0]] = target_ELS_list[i].replace("\t","").replace("\n","").replace("???????????? ??????","").rstrip()
            input_item[columns[1]] = ELS_composite_list[i].replace("\n"," ").rstrip().lstrip()
            input_item[columns[2]] = ELS_structure_list[i]
            input_item[columns[3]] = ELS_return[i]
            input_item[columns[4]] = ELS_loss[i]
            input_item[columns[5]] = interval_list[i]
            ELS_df = ELS_df.append(input_item,ignore_index = True)

        ELS_df['?????????'] = "??????????????????"
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
        mirae_asset_ELS['ELS???'] = ELS_name
        mirae_asset_ELS['????????????'] = ELS_composite
        mirae_asset_ELS['??????'] = ELS_structure
        mirae_asset_ELS['?????????'] = ELS_return
        mirae_asset_ELS['???????????????'] = ELS_loss
        mirae_asset_ELS['????????????'] = ELS_date
        mirae_asset_ELS['?????????'] = "??????????????????"
        mirae_asset_ELS.sort_values("?????????")

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

        ## ELS???
        for i in bsObject.findAll("strong",{"class":"tit h37 disp_b"}):
            ELS_name.append(i.text)

        ## ELS ????????? ?????????
        for idx,i in enumerate(bsObject.findAll("span",{"class" : "info_stx2_1"})):
            if (idx+1)%2 == 1:
                ELS_return.append(i.text)
            else:
                ELS_loss.append(i.text)

        ## ELS ????????????,????????????,-,-,??????,-
        for idx,i in enumerate(bsObject.findAll("span",{"class":"txt"})):
            if (idx+1)%6==1:
                ELS_composite.append(i.text)
            if (idx+1)%6==2:
                ELS_date.append(i.text)
            if (idx+1)%6==5:
                ELS_structure.append(i.text)

        NH_els_df = pd.DataFrame()
        NH_els_df['ELS???'] = ELS_name
        NH_els_df['???????????????'] = ELS_loss
        NH_els_df['?????????'] = ELS_return
        NH_els_df['????????????'] = ELS_composite[:-1]
        NH_els_df['????????????'] = ELS_date[:-1]
        NH_els_df['??????'] = ELS_structure
        NH_els_df['?????????'] = "NH????????????"

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
                ELS_name.append(i.text.replace("\n","").replace("\t","").split("???")[1])
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
        kiwoom_ELS['ELS???'] = ELS_name
        kiwoom_ELS['????????????'] = ELS_composite
        kiwoom_ELS['?????????'] = ELS_return
        kiwoom_ELS['???????????????'] = ELS_loss
        kiwoom_ELS['??????'] = ELS_structure
        kiwoom_ELS['????????????'] = ELS_date
        kiwoom_ELS['?????????'] = "????????????"

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
        ???????????? = []
        ???????????? = []
        ??????????????? = []
        ???????????? = []
        ??????????????? = []
        ??????????????? = []

        for i in bsObject.findAll('td',{'headers':'tb-c-1-2'}):
            target_str = i.text.replace("\n","").replace("\t","").rstrip().lstrip()
            bond_name.append(target_str.split("????????????")[0])
            credit_list.append(target_str.split("???????????? :")[1].split("??????????????????")[0].lstrip())

        for i in bsObject.findAll('td',{'headers':'tb-c-1-1'}):
            bond_class.append(i.text.replace("\n","").replace("\t","").rstrip().lstrip())

        for i in bsObject.findAll('td',{'headers':'tb-c-1-3'}):
            risk_list.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-1-4'}):
            maturity_list.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-1-5'}):
            remain_money.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-1-6'}):
            ????????????.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-1-7'}):
            ????????????.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-1-8'}):
            ???????????????.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-2-1'}):
            ????????????.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-2-2'}):
            ???????????????.append(i.text)

        for i in bsObject.findAll('td',{'headers':'tb-c-2-3'}):
            ???????????????.append(i.text)

        bond_df = pd.DataFrame()
        bond_df['????????????'] = bond_name
        bond_df['????????????'] = bond_class
        bond_df['????????????'] = credit_list
        bond_df['?????????'] = risk_list
        bond_df['?????????'] = maturity_list
        bond_df['????????????(???)'] = remain_money
        bond_df['????????????'] = ????????????
        bond_df['????????????'] = ????????????
        bond_df['???????????????'] = ???????????????
        bond_df['????????????'] = ????????????
        bond_df['???????????????'] = ???????????????
        bond_df['???????????????'] = ???????????????

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

        kiwoom_bond_df['????????????'] = bond_name
        kiwoom_bond_df['???????????????'] = bond_before_tax_return
        kiwoom_bond_df['???????????????'] = bond_after_tax_return
        kiwoom_bond_df['?????????'] = bond_expire_date
        kiwoom_bond_df['????????????'] = bond_credit

        return kiwoom_bond_df


    nh_df = nh_ELS()
    while len(nh_df)==0:
        nh_df = nh_ELS()

    korea_df = korea_investment_ELS()

    nh_df['????????????'] = nh_df['????????????'].str.replace(",",", ")
    korea_df['????????????'] = korea_df['????????????'].str.rstrip()
    korea_df['????????????'] = korea_df['????????????'].str.replace("\t","")
    korea_df['????????????'] = korea_df['????????????'].str.replace(" ",", ")


    target_df = kiwoom_ELS()
    target_df = pd.concat([target_df,nh_df])
    target_df = pd.concat([target_df,mirae_asset_ELS()])
    target_df = pd.concat([target_df,korea_df])
    target_df.reset_index(drop=True,inplace=True)

    target_df['????????????'] = target_df['????????????'].str.upper()
    target_df['????????????'] = target_df['????????????'].str.replace("INC.","")
    target_df['????????????'] = target_df['????????????'].str.rstrip()

    target_df['????????????'] = target_df['????????????'].str.replace("TESLA","?????????")
    target_df['????????????'] = target_df['????????????'].str.replace("Inc.","")

    target_df['?????????'] = target_df['?????????'].str.replace("??????","")
    target_df['?????????'] = target_df['?????????'].str.replace("???","")
    target_df['?????????'] = target_df['?????????'].str.replace("%","")
    target_df['?????????'] = target_df['?????????'].str.strip()
    target_df['?????????'] = target_df['?????????'].str.replace('??????????????? ????????? ??????',"")
    target_df = target_df[target_df['?????????'] != ""]
    target_df['?????????'] =target_df['?????????'].astype(float)

    target_df['???????????????'] = target_df['???????????????'].str.replace('???????????????',"")
    target_df['???????????????'] = target_df['???????????????'].str.strip()


    korea_investment_bond_df = korea_investment_bond()
    kiwoom_bond_df = kiwoom_bond()

    def knock_in_show(x):
        if len(x.split("(??????)")) !=1:
            return x.split("(??????)")[0][-2:]
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

    target_df['??????'] = target_df['??????'].apply(lambda x : knock_in_show(x))
    korea_investment_bond_df = korea_investment_bond_df[korea_investment_bond_df['????????????']=="?????????"]
    kiwoom_bond_df['?????????'] = "????????????"
    korea_investment_bond_df['?????????'] = "??????????????????"

    target_bond = pd.concat([kiwoom_bond_df,korea_investment_bond_df[kiwoom_bond_df.columns]])

    target_bond['???????????????'] = target_bond['???????????????'].astype(float)

    target_col = list(target_df.columns)
    target_col.remove("?????????")
    target_df = target_df[['?????????']+target_col]

    target_df['ELS???'] = target_df['ELS???'].str.replace("????????????????????????","")
    target_df['ELS???'] = target_df['ELS???'].str.replace("????????????????????????","")
    target_df['ELS???'] = target_df['ELS???'].str.replace("??????????????????","")

    def els_name_preprocessing(x):
        target = x.split("(")
        if "ELS" in x:
            return x
        if len(target)!=1:
            return x.split("(")[0]
        else:
            return x
    target_df['ELS???'] = target_df['ELS???'].apply(lambda x : els_name_preprocessing(x))

    unique_item = []
    for i in target_df['????????????'].str.split(", ").values:
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
                return "?????????"
        return "?????????"

    target_df['??????'] = target_df['????????????'].apply(lambda x : base_composite_preprocessing(x,unique_item))

    target_df.to_csv("./data/ELS??????.csv",index=False)
    target_bond.to_csv("./data/????????????.csv",index=False)

data_regeneration()

