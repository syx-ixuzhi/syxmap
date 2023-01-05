import os
import datetime
import pandas as pd
import streamlit as st
import numpy as np
from sqlalchemy import create_engine

db = 'postgresql://SyxAdmin:112358@ixuzhi.tpddns.cn:15432/SyxDatabase' #家里数据库

#建立公司数据库连接
def get_engine():
    return create_engine(db)
#读取门店信息
def get_store():
   return read_sql("SELECT * "
                   "FROM store_info "
                   "WHERE class='门店' "
                   "AND POSITION('停用' IN name) = 0 "
                   "AND POSITION('新开门' IN name) = 0 "
                   "AND POSITION('新开店' IN name) = 0 "
                   "ORDER BY CONVERT_TO (name, 'GBK');")

#初始化streamlit
def init_st(title,icon,layout):
    st.set_page_config(page_title=title, page_icon=icon, layout=layout, initial_sidebar_state='auto')
    hide_streamlit_style = """
    <style>
        #MainMenu {
            visibility: hidden;
        }
        footer {
            visibility: hidden;
        }
        .css-18e3th9 {
            padding-top: 2rem;
            padding-bottom: 2rem;
            padding-left: 2rem;
            padding-right: 2rem;
        }
        .info {
            height: 50px;
            background-color: rgb(235, 242, 251);
            border-radius: 5px;
            text-shadow: 1px 1px 2px gray;
            padding: 12px;
        }
        .block-container.css-1gx893w.egzxvld2 {
            margin-top: -50px;
        }
        .main.css-k1vhr4.egzxvld3 {
            margin-top: 20px;
        }
    </style>
    """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True)

#执行pd.read_sql语句
def read_sql(sql):
    engine=get_engine()
    df = pd.read_sql(sql=sql, con=engine)
    engine.dispose()
    return df

#执行execute sql语句
def exec_sql(sql):
    engine=get_engine()
    engine.execute(sql)
    engine.dispose()
