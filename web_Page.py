import streamlit as st
import pandas as pd
import os
import smtplib
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import datetime

st.set_page_config(page_title='Streamlit 프로토타입 만들기',page_icon='🎈',layout='wide')

st.text('🎈8조 프로젝트')
cus_info = pd.read_excel("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/고객정보.xlsx")
els_df = pd.read_excel("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/ELS모음.xlsx")
bond_df = pd.read_excel("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/채권모음.xlsx")

bond_df['잔존기간(일)'] = (pd.to_datetime(bond_df['만기일']) - datetime.datetime.today()).dt.days
bond_df['잔존기간(일)'] = bond_df['잔존기간(일)'].astype(int)

pb_name = st.text_input(label="PB명", value="PB이름")
input_user_name = st.text_input(label="고객명", value="김세원")
input_birth_day = st.text_input(label="생년월일", value="년도/월/일(예시 : 1996/10/12)")
email_id = st.text_input("네이버이메일주소")

password = st.text_input("Enter a password", type="password")
customer_id = cus_info[cus_info['고객명'] == input_user_name]['이메일'].values[0]

els_df.rename(columns = {"청약기간" : "청약마감일"},inplace=True)
def date_preprocessing(x):
    target_date =  x.split("~")
    if len(target_date) !=1:
        return target_date[1]
    else:
        return x

els_df['청약마감일'] = els_df['청약마감일'].apply(lambda x : date_preprocessing(x))
els_df['ELS명'] = els_df['ELS명'] + "\n(" + els_df['구조']
els_df.drop(columns = ["최대손실률","구조"],inplace=True)

els_sort = st.multiselect("ELS 분류기준",options=['수익률','낙인'])
df_sel_1 = els_df.sort_values(els_sort,ascending=False)
st.dataframe(df_sel_1)

bond_sort = st.selectbox("회사채 분류기준 ",options=['1년이하',"1년이상",'신용등급별최고수익률',"발행사 : 한국투자증권"])
df_sel_2 = bond_df
if bond_sort == "1년이하":
    df_sel_2 = bond_df[bond_df['잔존기간(일)']<=365].sort_values("세후수익률",ascending=False)
elif bond_sort == "1년이상":
    df_sel_2 = bond_df[bond_df['잔존기간(일)']>365].sort_values("세후수익률",ascending=False)
elif bond_sort == "신용등급별최고수익률":
    bond_list = pd.DataFrame(columns = bond_df.columns)
    for i in bond_df.sort_values('신용등급')['신용등급'].unique():
         bond_list = pd.concat([bond_list,bond_df[bond_df['신용등급'] == i].sort_values("세후수익률").tail(1)])
    df_sel_2 = bond_list
else:
    df_sel_2 = bond_df[bond_df['발행사']=="한국투자증권"].sort_values("세후수익률",ascending=False)
st.dataframe(df_sel_2)

if st.button("금융상품 데이터 업데이트"):
    from data_generation import data_regeneration
    data_regeneration()

if st.button("리포트 생성"):
    from report_generation import automatic_report_generate
    automatic_report_generate(input_user_name,pb_name,df_sel_1.head(5),df_sel_2.head(5))

if st.button("Email Send"):
    msg = MIMEMultipart()
    msg['From'] = email_id
    msg['To'] = customer_id
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = Header(s='{0}고객님 리포트 '.format(input_user_name), charset='utf-8')
    body = MIMEText('{0}고객님 리포트 입니다.'.format(input_user_name), _charset='utf-8')
    msg.attach(body)

    filename = '/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/Generated_Report/{0}고객님 리포트.docx'.format(input_user_name)
    attachment = open(filename,'rb')

    part = MIMEBase('application','octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',"attachment", filename= os.path.basename(filename))
    msg.attach(part)

    mailServer = smtplib.SMTP_SSL('smtp.naver.com',465)
    mailServer.login(email_id, password)  # 본인 계정과 비밀번호 사용.
    mailServer.send_message(msg)
    mailServer.quit()
