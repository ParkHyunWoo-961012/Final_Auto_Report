from docx import Document
from docx.shared import Pt
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
from docx.shared import RGBColor

from docx.enum.style import *
import pandas as pd

def automatic_report_generate(customer_name):
    els_df = pd.read_excel("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/ELS모음.xlsx")
    target_bond = pd.read_excel("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/채권모음.xlsx")

    document = Document('/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/템플릿.docx')
    today_date = datetime.datetime.now().strftime("%Y.%m.%d")

    paragraph3 = document.add_paragraph()

    run = paragraph3.add_run('{0} 고객님 귀하\n'.format(customer_name))
    run.bold = True
    run.underline = True
    paragraph3.add_run(today_date).bold = True

    paragraph3.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    head = document.add_heading(('"이번주는 쉬어갑니다."'), 0)

    head.runs[0].font.size=Pt(35)

    document.add_paragraph("_________________________________________________________________________________________________")

    document.add_heading("금주 추천 금융 상품 ({0} ~ {1})".format(today_date,(datetime.datetime.today()+datetime.timedelta(days=7)).strftime("%Y.%m.%d")), level = 0)

    document.add_heading('회사채', level = 0)

    bond_list = pd.DataFrame(columns = target_bond.columns)
    ## 1. 낙인 낮은 거중에 이자율 높은 것

    els_df['기초자산개수'] = els_df['기초자산'].apply(lambda x : len(x.split(",")))
    ## 2.이자율 가장 높은거
    high_return_ELS = els_df.sort_values('수익률',ascending = False).head(1)
    ##기초자산 3개 중 이자율 높은거
    base_asset3 = els_df[els_df['기초자산개수'] == 3].sort_values("수익률",ascending = False).head(1)
    ##기초자산 2개 중 이자율 높은거
    base_asset2 = els_df[els_df['기초자산개수'] == 2].sort_values("수익률",ascending = False).head(1)

    els_result = pd.DataFrame()
    els_result = pd.concat([els_result,high_return_ELS])
    els_result = pd.concat([els_result,base_asset3])
    els_result = pd.concat([els_result,base_asset2])

    for i in target_bond.sort_values('신용등급')['신용등급'].unique():
        bond_list = pd.concat([bond_list,target_bond[target_bond['신용등급'] == i].sort_values("세후수익률").tail(1)])
    table = document.add_table(bond_list.shape[0]+1, bond_list.shape[1],style = document.styles['Table Grid'])
    target_columns = bond_list.columns

    for r in range(len(table.rows)):
        row = table.rows[r]
        for c in range(len(row.cells)):
            if r == 0:
                cell = row.cells[c]
                cell.text = str(bond_list.columns[c])
            else:
                cell = row.cells[c]
                cell.text = str(bond_list.iloc[r-1][target_columns[c]])
    document.add_paragraph('\n')

    document.add_heading('ELS', level = 0)

    table = document.add_table(els_result.shape[0]+1, els_result.shape[1],style = document.styles['Table Grid'])
    target_columns = els_result.columns

    for r in range(len(table.rows)):
        row = table.rows[r]
        for c in range(len(row.cells)):
            if r == 0:
                cell = row.cells[c]
                cell.text = str(els_result.columns[c])
            else:
                cell = row.cells[c]
                cell.text = str(els_result.iloc[r-1][target_columns[c]])

    document.add_paragraph('\n')
    document.add_paragraph("_________________________________________________________________________________________________")

    document.add_heading('보유종목 Report', level = 0)

    #고객 보유자산 데이터 연결 필요 / 종목별 변동 사유는 어떻게할지 고민

    document.add_paragraph("_________________________________________________________________________________________________")

    document.add_heading('증시', level = 0)
    table = document.add_table(rows = 2, cols = 7)
    table.style = document.styles['Table Grid']
    first_row = table.rows[0].cells
    first_row[0].text = '코스피'
    first_row[1].text = '코스닥'
    first_row[2].text = 'S&P500'
    first_row[3].text = '나스닥'
    first_row[4].text = '국채3년'
    first_row[5].text = '국채10년'
    first_row[6].text = '유가'

    second_row = table.rows[1].cells
    domestic = pd.read_csv("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/국내지수.csv")
    nondomestic = pd.read_csv("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/해외지수.csv")
    wti = pd.read_csv("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/WTI.csv")
    domestic_bond = pd.read_csv("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/data/국채데이터.csv")

    import numpy as np

    domestic['코스피변화율'] = np.round(domestic['Kospi'].pct_change()*100,2)
    domestic['코스닥변화율'] = np.round(domestic['Kosdaq'].pct_change()*100,2)
    nondomestic['나스닥변화율'] = np.round(nondomestic['Nasdaq'].pct_change()*100,2)
    nondomestic['S&P변화율'] = np.round(nondomestic['S&P500'].pct_change()*100,2)

    wti['변화율'] = np.round(wti['Close'].pct_change()*100,2)
    domestic_bond_3year = domestic_bond[domestic_bond['종목명'] == "국채3년"]
    domestic_bond_10year = domestic_bond[domestic_bond['종목명'] == "국채10년"]

    second_row[0].text = str(domestic['코스피변화율'].tail(1).values[0]) +"%"

    def sign_define(x,locate,bond=False):
        if bond:
            if x>0:
                locate.text = "+" + str(x)
            elif x==0:
                locate.text = str(x)
            else:
                locate.text = "-" + str(x)
        else:
            if x>0:
                locate.text = "+" + str(x) + "%"
            elif x==0:
                locate.text = str(x)
            else:
                locate.text = "-" + str(x) + "%"
    sign_define(domestic['코스피변화율'].tail(1).values[0],second_row[0])
    sign_define(domestic['코스닥변화율'].tail(1).values[0],second_row[1])
    sign_define(nondomestic['S&P변화율'].tail(1).values[0],second_row[2])
    sign_define(nondomestic['나스닥변화율'].tail(1).values[0],second_row[3])

    sign_define(domestic_bond_3year['대비'].tail(1).values[0],second_row[4],True)
    sign_define(domestic_bond_10year['대비'].tail(1).values[0],second_row[5],True)
    sign_define(wti['변화율'].tail(1).values[0],second_row[6])

    # second_row[0].text = str(domestic['코스피변화율'].tail(1).values[0]) + "%"
    # second_row[1].text = str(domestic['코스닥변화율'].tail(1).values[0]) + "%"
    # second_row[2].text = str(nondomestic['S&P변화율'].tail(1).values[0]) + "%"
    # second_row[3].text = str(nondomestic['나스닥변화율'].tail(1).values[0]) + "%"
    # second_row[4].text = str(domestic_bond_3year['대비'].tail(1).values[0])
    # second_row[5].text = str(domestic_bond_10year['대비'].tail(1).values[0])
    # second_row[6].text = str(wti['변화율'].tail(1).values[0]) + "%"

    #2행 변동폭 표기 필요
    paragraph3 = document.add_paragraph('\n')
    paragraph3.add_run('김한국, 여의도 영업부/113240@koreainvestment.com').bold = True

    paragraph3.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    document.save("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/Generated_Report/{0}고객님 리포트.docx".format(customer_name))


