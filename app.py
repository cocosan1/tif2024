import streamlit as st
import pandas as pd
import math
import openpyxl
from io import BytesIO


st.title('TIF集計')

uploaded_file = st.file_uploader("Choose a Excel file", type='xlsx') 
# excelファイルを指定

df = pd.DataFrame()
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name='入力')


df2 = df[['得意先名', '金額', '取引先担当', '売価']]
df2['担当者2'] = df2['得意先名'] + '/' + df2['取引先担当']
with st.expander('df col絞り込み', expanded=False):
    st.write(df2)

s_person = df2.groupby(['担当者2'])['売価'].sum()

df_person = pd.DataFrame(s_person).reset_index()
df_person['売価'] = df_person['売価'].apply(lambda x: math.floor(x))
df_person['QUOカード'] = df_person['売価'].apply(lambda x: math.floor(x/100000))
df_person = df_person[['担当者2', '売価', 'QUOカード']]

with st.expander('df QUO計算', expanded=False):
    st.write(df_person)

df_sort = df_person.sort_values('売価',ascending=False).reset_index(drop=True)
df_sort.index = range(1, len(df_sort) + 1)
st.write(df_sort)

def to_excel(df):
    #メモリ内にバッファを作成し、バイナリデータをその中に書き込む       
    output = BytesIO()
    #dfをバイナリデータ（outputに書き込まれる）として指定されたシート（'Sheet1'）
    # にエクスポート
    df.to_excel(output, index = False, sheet_name='Sheet1')
    #outputに書き込まれたバイナリデータを取得
    processed_data = output.getvalue()

    return processed_data

# 関数実行
df_xlsx = to_excel(df_sort)
# ダウンロードボタン
st.download_button(label='Download Excel file', data=df_xlsx, file_name= 'TIF進捗状況.xlsx')