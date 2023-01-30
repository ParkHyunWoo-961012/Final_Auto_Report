import streamlit as st
import pandas as pd
from market_data_generation import market_data_generate

market_data_generate()

st.set_page_config(page_title='Streamlit 프로토타입 만들기',page_icon='🎈',layout='wide')

st.text('🎈8조 프로젝트')

els_df = pd.read_excel("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/ELS모음.xlsx")
bond_df = pd.read_excel("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/채권모음.xlsx")

st.markdown("ELS")
st.dataframe(els_df.sort_values("수익률"))

st.markdown("회사채")
st.dataframe(bond_df.sort_values("세후수익률"))

input_user_name = st.text_input(label="고객명", value="고객 이름")
input_birth_day = st.text_input(label="생년월일", value="1996/10/12")

if st.button("이메일보내기"):
    con = st.container()
    if input_user_name == "고객 이름":
        con.error("Input your name please~")
    else:
        con.write(f"Hello~ {str(input_user_name)}")

if st.button("금융상품 데이터 업데이트"):
    from data_generation import data_regeneration
    data_regeneration()

if st.button("리포트 생성"):
    from report_generation import automatic_report_generate
    automatic_report_generate(input_user_name)

