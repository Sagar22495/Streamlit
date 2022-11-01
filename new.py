import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import openpyxl



st. set_page_config (page_title =' Makerting Available Material Quantity', page_icon =':smiley',layout="wide")

exce_file='Book1.xlsx'
sheet_name='Sheet1'

# url=":bar_chart: Makerting Available Material Quantity"


# def header(url):
#     st.markdown(f'<p style="background-color:orange;color:#33ff33;font-size:24px;border-radius:2%;">{url}</p>',
#                 unsafe_allow_html=True)
st.title(":bar_chart: Makerting Available Material Quantity")
# st.subheader('HR Rect + EE')

wb = openpyxl.load_workbook('Book1.xlsx')
sheet = wb['Sheet1']

E = sheet['E1'].value

Ed = sheet['E2'].value

F = sheet['F1'].value

Fd = sheet['F2'].value

G = sheet['G1'].value

Gd = sheet['G2'].value

c,d,e = st.columns(3,gap="large")

with c:
    st.metric(label="Welcome Kit Available",
              value=E,
              delta=Ed)

with d:
    st.metric(label="Tablet Available",
              value=F,
              delta=Fd)

with e:
    st.metric(label="All Project Brochure Available",
              value=G,
              delta=Gd)


st.subheader(":tshirt: Available Tshirts & Kurtis")
a,b=st.columns(2,gap="large")

with a:
    df= pd.read_excel(exce_file,
                           sheet_name=sheet_name,
                           usecols='R:S',
                           header=0)
    # st.dataframe(df_new)

    pie = px.pie(df,
                 title='Total Available Tshirt In Sizes',
                 values='Qty',
                 names='Tshirt',
                 width=600,
                 height=400
                 )
    st.plotly_chart(pie,use_container_width=True)

with b:
    df_new = pd.read_excel(exce_file,
                       sheet_name=sheet_name,
                       usecols='O:P',
                       header=0)
    # st.dataframe(df_new)

    pie=px.pie(df_new,
               title='Total Available Tshirt And Kurti',
               values='Available Tshirt &Kurti Qty',
               names='Tshirts & Kurti',
               width=600,
               height=400
               )
    st.plotly_chart(pie,use_container_width=True)

st.subheader(":package: Welcome Kit")
col1,col2 = st.columns(2)

with col1:
    df_new = pd.read_excel(exce_file,
                           sheet_name=sheet_name,
                           usecols='U:V',
                           header=0)
    # st.dataframe(df_new)

    pie = px.bar(df_new,
                 title="Contents Available In Welcome Kit",
                 x="Welcome Kit",
                 y="Count",
                 template="plotly_white",
                 width=600,
                 height=400
                 )
    st.plotly_chart(pie,use_container_width=True)









