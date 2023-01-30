from docx import Document
from docx.shared import Pt
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
import pandas as pd
from docx.shared import RGBColor
from docx.enum.style import *


def automatic_report_generate(customer_name,pb_name,els_df,target_bond):
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

    domestic['Kospi'] = np.round(domestic['Kospi'],2)
    second_row[0].text = str(domestic['코스피변화율'].tail(1).values[0]) +"%"

    def sign_define(x,close,locate,bond=False):
        close = np.round(close,2)
        if bond:
            if x>0:
                locate.text = str(close) + "\n(+" + str(x) + ")"
            elif x==0:
                locate.text = str(close) + "\n(" + str(x) + ")"
            else:
                locate.text = str(close) + "\n(-" + str(x) + ")"
        else:
            if x>0:
                locate.text = str(close) + "\n(+" + str(x) + "%)"
            elif x==0:
                locate.text = str(close) + "\n(" + str(x) + ")"
            else:
                locate.text = str(close) + "\n(-" + str(x) + ")"

    sign_define(domestic['코스피변화율'].tail(1).values[0],domestic['Kospi'].tail(1).values[0] ,second_row[0])
    sign_define(domestic['코스닥변화율'].tail(1).values[0],domestic['Kosdaq'].tail(1).values[0] ,second_row[1])
    sign_define(nondomestic['S&P변화율'].tail(1).values[0],nondomestic['S&P500'].tail(1).values[0],second_row[2])
    sign_define(nondomestic['나스닥변화율'].tail(1).values[0],nondomestic['Nasdaq'].tail(1).values[0],second_row[3])

    sign_define(domestic_bond_3year['대비'].tail(1).values[0],domestic_bond_3year['수익률'].tail(1).values[0],second_row[4],True)
    sign_define(domestic_bond_10year['대비'].tail(1).values[0],domestic_bond_10year['수익률'].tail(1).values[0],second_row[5],True)
    sign_define(wti['변화율'].tail(1).values[0],wti['Close'].tail(1).values[0],second_row[6])

    document.add_paragraph("_________________________________________________________________________________________________")

    document.add_heading("금주 추천 금융 상품 ({0} ~ {1})".format(today_date,(datetime.datetime.today()+datetime.timedelta(days=7)).strftime("%Y.%m.%d")), level = 0)

    document.add_heading('회사채', level = 0)

    bond_list = pd.DataFrame(columns = target_bond.columns)

    # for i in target_bond.sort_values('신용등급')['신용등급'].unique():
    #     bond_list = pd.concat([bond_list,target_bond[target_bond['신용등급'] == i].sort_values("세후수익률").tail(1)])
    table = document.add_table(target_bond.shape[0]+1, target_bond.shape[1],style = document.styles['Table Grid'])

    target_columns = target_bond.columns

    for r in range(len(table.rows)):
        row = table.rows[r]
        for c in range(len(row.cells)):
            if r == 0:
                cell = row.cells[c]
                cell.text = str(target_bond.columns[c])
            else:
                cell = row.cells[c]
                cell.text = str(target_bond.iloc[r-1][target_columns[c]])
    document.add_paragraph('\n')

    document.add_heading('ELS', level = 0)

    table = document.add_table(els_df.shape[0]+1, els_df.shape[1],style = document.styles['Table Grid'])
    table.autofit = True
    target_columns = els_df.columns

    for r in range(len(table.rows)):
        row = table.rows[r]
        for c in range(len(row.cells)):
            if r == 0:
                cell = row.cells[c]
                cell.text = str(els_df.columns[c])
            else:
                cell = row.cells[c]
                cell.text = str(els_df.iloc[r-1][target_columns[c]])

    document.add_paragraph('\n')
    document.add_paragraph("_________________________________________________________________________________________________")

    document.add_heading('보유종목 Report', level = 0)

    #2행 변동폭 표기 필요
    paragraph3 = document.add_paragraph('\n')
    paragraph3.add_run('{0}, 여의도 영업부/113240@koreainvestment.com'.format(pb_name)).bold = True

    paragraph3.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    document.save("/Users/hyunwoo/PycharmProjects/pythonProject/HanTwoProject/8_BoKum/Generated_Report/{0}고객님 리포트.docx".format(customer_name))

